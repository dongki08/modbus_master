using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Modbus.Device; // NModbus4 2.1.0

namespace modbus_master
{
    public partial class MainWindow : Window
    {
        #region Private Fields
        private TcpClient _tcpClient;
        private IModbusMaster _master;
        private volatile bool _isConnected;

        private CancellationTokenSource _pollCts;
        private readonly object _sync = new object();
        private readonly DataTable _table = new DataTable();
        private readonly Stopwatch _sw = new Stopwatch();
        private readonly SemaphoreSlim _pollGate = new SemaphoreSlim(1, 1);
        private readonly SemaphoreSlim _writeGate = new SemaphoreSlim(1, 1);
        
        // 비트 변경 최적화를 위한 디바운싱
        private readonly ConcurrentDictionary<string, CancellationTokenSource> _writeDelayTasks = new ConcurrentDictionary<string, CancellationTokenSource>();
        private volatile bool _isUpdatingFormats = false;

        // Stats
        private long _ok;
        private long _err;
        private double _rttLast;
        private double _rttAvg;

        // UI Resources from XAML
        private readonly Brush _successBrush;
        private readonly Brush _errorBrush;
        #endregion

        #region Constructor & Initialization
        public MainWindow()
        {
            InitializeComponent();
            InitGrid();
            UpdateUiState();
            SetupEventHandlers();

            _successBrush = (Brush)FindResource("SuccessBrush");
            _errorBrush = (Brush)FindResource("ErrorBrush");
        }

        private void SetupEventHandlers()
        {
            dgData.CellEditEnding += DgData_CellEditEnding;
        }

        private void InitGrid()
        {
            _table.Columns.Add("Address", typeof(string));
            _table.Columns.Add("Value", typeof(ushort));
            _table.Columns.Add("ValueHex", typeof(string));
            _table.Columns.Add("ValueHighChar", typeof(string));
            _table.Columns.Add("ValueLowChar", typeof(string));
            _table.Columns.Add("ValueString", typeof(string));
            
            for (int i = 0; i < 16; i++)
            {
                _table.Columns.Add($"Bit{i}", typeof(bool));
            }
            
            dgData.ItemsSource = _table.DefaultView;
            _table.ColumnChanged += Table_ColumnChanged;
        }
        #endregion

        #region Custom Window Chrome
        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2) AdjustWindowSize();
            else DragMove();
        }
        
        private void Minimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
        private void Maximize_Click(object sender, RoutedEventArgs e) => AdjustWindowSize();
        private void Close_Click(object sender, RoutedEventArgs e) => Close();
        
        private void AdjustWindowSize() => 
            WindowState = (WindowState == WindowState.Maximized) ? WindowState.Normal : WindowState.Maximized;
        #endregion

        #region UI Toggle Events
        private void ReadToggle_Checked(object sender, RoutedEventArgs e)
        {
            if (WriteToggle == null || ReadSettingsPanel == null || WriteSettingsPanel == null) return;
            WriteToggle.IsChecked = false;
            ReadSettingsPanel.Visibility = Visibility.Visible;
            WriteSettingsPanel.Visibility = Visibility.Collapsed;
        }

        private void WriteToggle_Checked(object sender, RoutedEventArgs e)
        {
            if (ReadToggle == null || ReadSettingsPanel == null || WriteSettingsPanel == null) return;
            ReadToggle.IsChecked = false;
            ReadSettingsPanel.Visibility = Visibility.Collapsed;
            WriteSettingsPanel.Visibility = Visibility.Visible;
        }
        #endregion

        #region DataGrid Event Handlers
        private async void DgData_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit) return;
            
            if (!(e.EditingElement is TextBox textBox)) return;
            
            var columnHeader = e.Column.Header.ToString();
            var newValue = textBox.Text;
            var row = ((DataRowView)e.Row.Item).Row;
            
            try
            {
                if (TryConvertValue(columnHeader, newValue, out ushort convertedValue))
                {
                    row["Value"] = convertedValue;
                    UpdateRowFormats(row, convertedValue);
                    await WriteRegisterFromRow(row);
                }
                else
                {
                    e.Cancel = true;
                    Log($"[!] Invalid value: {newValue} for {columnHeader}");
                }
            }
            catch (Exception ex)
            {
                e.Cancel = true;
                LogError($"Value conversion failed", ex);
            }
        }

        private bool TryConvertValue(string columnHeader, string value, out ushort result)
        {
            result = 0;
    
            switch (columnHeader)
            {
                case "Dec":
                    return ushort.TryParse(value, out result);
                case "Hex":
                    return ushort.TryParse(value.Replace("0x", ""), System.Globalization.NumberStyles.HexNumber, null, out result);
                case "String":
                    return TryConvertStringToUShort(value, out result);
                default:
                    return false;
            }
        }

        private bool TryConvertStringToUShort(string value, out ushort result)
        {
            result = 0;
            
            if (value.Length > 2) return false;

            byte highByte = 0, lowByte = 0;
            
            if (value.Length >= 1) lowByte = (byte)value[0];
            if (value.Length == 2) 
            {
                highByte = (byte)value[0];
                lowByte = (byte)value[1];
            }
            
            result = (ushort)((highByte << 8) | lowByte);
            return true;
        }

        private async void Table_ColumnChanged(object sender, DataColumnChangeEventArgs e)
        {
            try
            {
                if (!IsBitColumn(e?.Column?.ColumnName)) return;
                if (e.Row == null || _isUpdatingFormats) return;

                var address = e.Row["Address"].ToString();
                await ScheduleDelayedWrite(address, e.Row);
            }
            catch (Exception ex)
            {
                LogError("Error in Table_ColumnChanged", ex);
            }
        }

        private bool IsBitColumn(string columnName) => 
            !string.IsNullOrEmpty(columnName) && columnName.StartsWith("Bit");

        private async Task ScheduleDelayedWrite(string address, DataRow row)
        {
            // 기존 지연 작업 취소
            if (_writeDelayTasks.TryRemove(address, out var existingCts))
            {
                existingCts.Cancel();
                existingCts.Dispose();
            }

            // 새로운 지연 작업 생성
            var newCts = new CancellationTokenSource();
            _writeDelayTasks[address] = newCts;

            try
            {
                await Task.Delay(150, newCts.Token); // 150ms 디바운싱
                
                if (!newCts.Token.IsCancellationRequested)
                {
                    await ExecuteBitChange(row);
                }
            }
            catch (TaskCanceledException)
            {
                // 정상적인 취소
            }
            finally
            {
                _writeDelayTasks.TryRemove(address, out _);
                newCts.Dispose();
            }
        }

        private async Task ExecuteBitChange(DataRow row)
        {
            try
            {
                _isUpdatingFormats = true;

                var newValue = CalculateValueFromBits(row);
                
                await Dispatcher.InvokeAsync(() =>
                {
                    row["Value"] = newValue;
                    UpdateRowDisplayFormats(row, newValue);
                });

                await WriteRegisterFromRow(row);
            }
            finally
            {
                _isUpdatingFormats = false;
            }
        }

        private ushort CalculateValueFromBits(DataRow row)
        {
            ushort value = 0;
            for (int i = 0; i < 16; i++)
            {
                if (row[$"Bit{i}"] is bool bitValue && bitValue)
                {
                    value |= (ushort)(1 << i);
                }
            }
            return value;
        }
        #endregion

        #region Connection Management
        private async void btnToggleConnect_Click(object sender, RoutedEventArgs e)
        {
            if (_isConnected) await DisconnectAsync();
            else await ConnectAsync();
        }

        private async Task ConnectAsync()
        {
            if (!ValidateConnectionInputs()) return;

            try
            {
                await DisconnectAsync();
                
                var ip = txtIp.Text;
                var port = int.Parse(txtPort.Text);
                
                _tcpClient = new TcpClient 
                { 
                    NoDelay = true, 
                    ReceiveTimeout = 2000, 
                    SendTimeout = 2000 
                };
                
                Log($"[*] Connecting to {ip}:{port}...");
                
                var connectTask = _tcpClient.ConnectAsync(ip, port);
                var timeoutTask = Task.Delay(4000);
                
                if (await Task.WhenAny(connectTask, timeoutTask) == timeoutTask)
                {
                    throw new TimeoutException("Connection timeout.");
                }
                
                if (!_tcpClient.Connected)
                {
                    throw new InvalidOperationException("Failed to connect.");
                }
                
                _master = ModbusIpMaster.CreateIp(_tcpClient);
                _master.Transport.ReadTimeout = 1500;
                _master.Transport.WriteTimeout = 1500;
                
                _isConnected = true;
                Log("[+] Connected successfully.");
                SetConnIndicator(true, "Connected");
            }
            catch (Exception ex)
            {
                _isConnected = false;
                LogError("Connection failed", ex);
            }
            finally 
            { 
                UpdateUiState(); 
            }
        }

        private bool ValidateConnectionInputs()
        {
            if (!IPAddress.TryParse(txtIp.Text, out _))
            {
                Log("[!] Invalid IP address");
                return false;
            }
            
            if (!int.TryParse(txtPort.Text, out var port) || port < 1 || port > 65535)
            {
                Log("[!] Invalid port");
                return false;
            }
            
            if (!byte.TryParse(txtUnitId.Text, out _))
            {
                Log("[!] Invalid UnitId");
                return false;
            }
            
            return true;
        }

        private async Task DisconnectAsync()
        {
            StopPolling();
            await CancelAllDelayedWrites();
            
            _master?.Dispose();
            _master = null;
            _tcpClient?.Dispose();
            _tcpClient = null;
            
            if (_isConnected)
            {
                _isConnected = false;
                Log("[-] Disconnected.");
            }
            
            SetConnIndicator(false, "Disconnected");
            UpdateUiState();
        }

        private async Task CancelAllDelayedWrites()
        {
            var tasks = _writeDelayTasks.Values.ToList();
            foreach (var cts in tasks)
            {
                cts.Cancel();
            }
            _writeDelayTasks.Clear();
            
            // 짧은 대기로 진행 중인 작업들이 완료되도록 함
            await Task.Delay(50);
        }
        #endregion

        #region Read Operations
        private async void btnReadOnce_Click(object sender, RoutedEventArgs e) => await ReadOnceAsync();

        private async Task ReadOnceAsync()
        {
            if (!_isConnected || !ValidateReadInputs()) return;

            var unitId = byte.Parse(txtUnitId.Text);
            var start = ushort.Parse(txtStart.Text);
            var count = ushort.Parse(txtCount.Text);

            await _pollGate.WaitAsync();
            try
            {
                _sw.Restart();
                
                var data = await ReadModbusData(unitId, start, count);
                ApplyReadData(start, data);
                
                _sw.Stop();
                OnSuccess(_sw.Elapsed.TotalMilliseconds);
            }
            catch (Exception ex)
            {
                _sw.Stop();
                OnError(_sw.Elapsed.TotalMilliseconds, ex);
                
                if (chkAutoReconnect.IsChecked == true)
                {
                    await TryAutoReconnectAsync();
                }
            }
            finally 
            { 
                _pollGate.Release(); 
            }
        }

        private bool ValidateReadInputs()
        {
            if (!byte.TryParse(txtUnitId.Text, out _))
            {
                Log("[!] Invalid UnitId.");
                return false;
            }
            
            if (!ushort.TryParse(txtStart.Text, out _))
            {
                Log("[!] Invalid Start address.");
                return false;
            }
            
            if (!ushort.TryParse(txtCount.Text, out var count) || count < 1)
            {
                Log("[!] Invalid Count.");
                return false;
            }
            
            return true;
        }

        private async Task<Array> ReadModbusData(byte unitId, ushort start, ushort count)
        {
            switch (cmbFunction.SelectedIndex)
            {
                case 0:
                    return await _master.ReadCoilsAsync(unitId, start, count);
                case 1:
                    return await _master.ReadInputsAsync(unitId, start, count);
                case 2:
                    return await _master.ReadHoldingRegistersAsync(unitId, start, count);
                case 3:
                    return await _master.ReadInputRegistersAsync(unitId, start, count);
                default:
                    throw new InvalidOperationException("Invalid function selection");
            }
        }

        private void ApplyReadData(ushort start, Array data)
        {
            lock (_sync)
            {
                EnsureRows(start, (ushort)data.Length);
        
                for (int i = 0; i < data.Length; i++)
                {
                    ushort value;
                    var dataValue = data.GetValue(i);
            
                    if (dataValue is bool b)
                        value = b ? (ushort)1 : (ushort)0;
                    else if (dataValue is ushort us)
                        value = us;
                    else
                        value = (ushort)0;
            
                    _table.Rows[i]["Value"] = value;
                    UpdateRowFormats(_table.Rows[i], value);
                }
            }
        }

        private void EnsureRows(ushort start, ushort length)
        {
            if (_table.Rows.Count != length)
            {
                _table.Clear();
                for (ushort i = 0; i < length; i++)
                {
                    var row = _table.NewRow();
                    row["Address"] = $"{start + i}";
                    _table.Rows.Add(row);
                }
            }
        }

        private void UpdateRowFormats(DataRow row, ushort value)
        {
            if (_isUpdatingFormats) return;
            
            _isUpdatingFormats = true;
            try
            {
                _table.ColumnChanged -= Table_ColumnChanged;
                UpdateRowDisplayFormats(row, value);
                UpdateRowBits(row, value);
                _table.ColumnChanged += Table_ColumnChanged;
            }
            finally
            {
                _isUpdatingFormats = false;
            }
        }

        private void UpdateRowDisplayFormats(DataRow row, ushort value)
        {
            row["ValueHex"] = $"0x{value:X4}";
            
            var (highByte, lowByte) = ((byte)(value >> 8), (byte)(value & 0xFF));
            
            row["ValueHighChar"] = GetPrintableChar(highByte);
            row["ValueLowChar"] = GetPrintableChar(lowByte);
            row["ValueString"] = GetStringRepresentation(highByte, lowByte);
        }

        private void UpdateRowBits(DataRow row, ushort value)
        {
            for (int i = 0; i < 16; i++)
            {
                row[$"Bit{i}"] = (value & (1 << i)) != 0;
            }
        }

        private string GetPrintableChar(byte b) => 
            (b >= 32 && b <= 126) ? ((char)b).ToString() : ".";

        private string GetStringRepresentation(byte highByte, byte lowByte)
        {
            if (highByte == 0)
                return GetPrintableChar(lowByte).Replace(".", "");
            if (lowByte == 0)
                return GetPrintableChar(highByte).Replace(".", "");
            
            var highChar = GetPrintableChar(highByte).Replace(".", "");
            var lowChar = GetPrintableChar(lowByte).Replace(".", "");
            return highChar + lowChar;
        }
        #endregion

        #region Polling Operations
        private async void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if (_pollCts != null)
            {
                Log("[!] Polling is already running.");
                return;
            }
            
            if (!int.TryParse(txtPeriod.Text, out var period) || period < 50)
            {
                Log("[!] Period must be >= 50ms.");
                return;
            }
            
            _pollCts = new CancellationTokenSource();
            Log($"[*] Polling started (period: {period}ms).");
            UpdateUiState();
            
            await Task.Run(() => PollLoop(period, _pollCts.Token), _pollCts.Token);
        }

        private async Task PollLoop(int period, CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                await Dispatcher.InvokeAsync(ReadOnceAsync);
                
                try 
                { 
                    await Task.Delay(period, ct); 
                } 
                catch (TaskCanceledException) 
                { 
                    break; 
                }
            }
        }
        
        private void btnStop_Click(object sender, RoutedEventArgs e) => StopPolling();

        private void StopPolling()
        {
            if (_pollCts != null)
            {
                _pollCts.Cancel();
                _pollCts.Dispose();
                _pollCts = null;
                Log("[*] Polling stopped.");
                UpdateUiState();
            }
        }
        #endregion

        #region Write Operations
        private async void btnWrite_Click(object sender, RoutedEventArgs e)
        {
            if (!_isConnected || !ValidateWriteInputs()) return;

            var unitId = byte.Parse(txtUnitId.Text);
            var addr = ushort.Parse(txtWriteAddr.Text);

            try
            {
                await ExecuteWriteOperation(unitId, addr);
            }
            catch (Exception ex)
            {
                LogError("Write operation failed", ex);
                if (chkAutoReconnect.IsChecked == true)
                {
                    await TryAutoReconnectAsync();
                }
            }
        }

        private bool ValidateWriteInputs()
        {
            if (!byte.TryParse(txtUnitId.Text, out _))
            {
                Log("[!] Invalid UnitId.");
                return false;
            }
            
            if (!ushort.TryParse(txtWriteAddr.Text, out _))
            {
                Log("[!] Invalid write address.");
                return false;
            }
            
            return true;
        }

        private async Task ExecuteWriteOperation(byte unitId, ushort addr)
        {
            switch (cmbWriteType.SelectedIndex)
            {
                case 0: // Coil
                    if (TryParseBool(txtWriteValue.Text, out var coilVal))
                    {
                        await _master.WriteSingleCoilAsync(unitId, addr, coilVal);
                        Log($"[+] Wrote coil {addr} = {coilVal}");
                    }
                    else
                    {
                        Log("[!] Coil value must be 0/1 or true/false.");
                    }
                    break;
                    
                case 1: // Single Register
                    if (ushort.TryParse(txtWriteValue.Text, out var regVal))
                    {
                        await _master.WriteSingleRegisterAsync(unitId, addr, regVal);
                        Log($"[+] Wrote register {addr} = {regVal}");
                    }
                    else
                    {
                        Log("[!] Register value must be 0-65535.");
                    }
                    break;
                    
                case 2: // Multiple Registers
                    await WriteMultipleRegisters(unitId, addr);
                    break;
            }
        }

        private async Task WriteRegisterFromRow(DataRow row)
        {
            if (!_isConnected || cmbFunction.SelectedIndex != 2) return;
            if (!byte.TryParse(txtUnitId.Text, out var unitId)) return;
            if (!ushort.TryParse(row["Address"].ToString(), out var addr)) return;

            var value = (ushort)row["Value"];
            
            await _writeGate.WaitAsync();
            try
            {
                await _master.WriteSingleRegisterAsync(unitId, addr, value);
                Log($"[+] Auto-Wrote Register {addr} = {value}");
            }
            catch (Exception ex)
            {
                LogError($"Auto-Write failed for address {addr}", ex);
                if (chkAutoReconnect.IsChecked == true)
                {
                    await TryAutoReconnectAsync();
                }
            }
            finally
            {
                _writeGate.Release();
            }
        }

        private async Task WriteMultipleRegisters(byte unitId, ushort startAddr)
        {
            if (dgData.SelectedItems.Count == 0)
            {
                Log("[!] No rows selected for multiple write.");
                return;
            }
            
            var values = dgData.SelectedItems.Cast<DataRowView>()
                .Select(item => (ushort)item.Row["Value"])
                .ToArray();
                
            await _master.WriteMultipleRegistersAsync(unitId, startAddr, values);
            Log($"[+] Wrote {values.Length} registers from {startAddr}.");
        }

        private void btnUseSelected_Click(object sender, RoutedEventArgs e)
        {
            if (dgData.SelectedItem is DataRowView view)
            {
                txtWriteAddr.Text = view["Address"].ToString();
                txtWriteValue.Text = view["Value"].ToString();
            }
        }

        private void dgData_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (dgData.SelectedItem is DataRowView view)
            {
                txtWriteAddr.Text = view["Address"].ToString();
                txtWriteValue.Text = view["Value"].ToString();
                Log($"[*] Selected address={txtWriteAddr.Text}, value={txtWriteValue.Text} for writing.");
            }
        }
        #endregion

        #region Statistics & UI Updates
        private void OnSuccess(double ms)
        {
            _rttLast = ms;
            _ok++;
            _rttAvg = (_rttAvg * (_ok - 1) + ms) / _ok;
            UpdateStats();
        }

        private void OnError(double ms, Exception ex)
        {
            _rttLast = ms;
            _err++;
            UpdateStats();
            LogError("Read error (Connection may be lost)", ex);
        }

        private void UpdateStats()
        {
            Dispatcher.Invoke(() =>
            {
                lblRttLast.Text = _rttLast.ToString("F1");
                lblRttAvg.Text = _ok > 0 ? _rttAvg.ToString("F1") : "-";
                lblOk.Text = _ok.ToString();
                lblErr.Text = _err.ToString();
            });
        }

        private void UpdateUiState()
        {
            Dispatcher.Invoke(() =>
            {
                bool isPolling = _pollCts != null;
                btnToggleConnect.Content = _isConnected ? "🚀 Disconnect" : "🚀 Connect";
                bool canInteract = _isConnected && !isPolling;
                
                btnReadOnce.IsEnabled = canInteract;
                btnStart.IsEnabled = _isConnected && !isPolling;
                btnStop.IsEnabled = isPolling;
                btnWrite.IsEnabled = _isConnected;
                
                txtIp.IsEnabled = !_isConnected;
                txtPort.IsEnabled = !_isConnected;
                txtUnitId.IsEnabled = !_isConnected;
            });
        }
        
        private void SetConnIndicator(bool on, string text)
        {
            Dispatcher.Invoke(() =>
            {
                elConn.Fill = on ? _successBrush : _errorBrush;
                lblConn.Text = text;
            });
        }
        #endregion

        #region Utility Methods
        private bool TryParseBool(string text, out bool value)
        {
            text = text?.Trim().ToLower();
            if (text == "1" || text == "true")
            {
                value = true;
                return true;
            }
            if (text == "0" || text == "false")
            {
                value = false;
                return true;
            }
            value = false;
            return false;
        }

        private void Log(string msg)
        {
            Dispatcher.Invoke(() =>
            {
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss.fff}] {msg}{Environment.NewLine}");
                txtLog.ScrollToEnd();
            });
        }

        private void LogError(string title, Exception ex) => 
            Log($"[-] {title}: {ex.Message}");

        private async Task TryAutoReconnectAsync()
        {
            if (_isConnected) return;
            
            Log("[*] Connection lost. Attempting auto-reconnect...");
            await DisconnectAsync();
            await Task.Delay(1000);
            await ConnectAsync();
        }
        #endregion

        #region Export & UI Actions
        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ModbusMaster");
                Directory.CreateDirectory(dir);
                var file = Path.Combine(dir, $"snapshot_{DateTime.Now:yyyyMMdd_HHmmss}.csv");
                
                var sw = new StreamWriter(file, false, Encoding.UTF8);
                
                // CSV 헤더
                var headers = new List<string> { "Address", "Dec", "Hex", "HighChar", "LowChar", "String" };
                headers.AddRange(Enumerable.Range(0, 16).Reverse().Select(i => $"Bit{i}"));
                sw.WriteLine(string.Join(",", headers));
                
                // 데이터 행
                foreach (DataRow row in _table.Rows)
                {
                    var values = new List<string>
                    {
                        row["Address"].ToString(),
                        row["Value"].ToString(),
                        row["ValueHex"].ToString(),
                        row["ValueHighChar"].ToString(),
                        row["ValueLowChar"].ToString(),
                        row["ValueString"].ToString()
                    };
                    
                    values.AddRange(Enumerable.Range(0, 16).Reverse().Select(i => row[$"Bit{i}"].ToString()));
                    sw.WriteLine(string.Join(",", values));
                }
                
                Log($"[+] Exported data to: {file}");
            }
            catch (Exception ex)
            {
                LogError("Export failed", ex);
            }
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            _table.Clear();
            _ok = _err = 0;
            _rttLast = _rttAvg = 0;
            UpdateStats();
            Log("[*] Data grid and stats cleared.");
        }

        private void btnCopyLog_Click(object sender, RoutedEventArgs e) => 
            Clipboard.SetText(txtLog.Text);

        private void btnClearLog_Click(object sender, RoutedEventArgs e) => 
            txtLog.Clear();
        #endregion

        #region Cleanup
        protected override async void OnClosing(CancelEventArgs e)
        {
            StopPolling();
            await DisconnectAsync();
            base.OnClosing(e);
        }
        #endregion
    }
}