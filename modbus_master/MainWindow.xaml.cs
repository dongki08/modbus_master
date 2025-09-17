using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Modbus.Device; // NModbus4 2.1.0

namespace modbus_master
{
    public partial class MainWindow : Window
    {
        #region Unit Tab Data Model
        public class UnitTabData : INotifyPropertyChanged
        {
            private string _header;
            private byte _unitId;
            private TcpClient _tcpClient;
            private IModbusMaster _master;
            private bool _isConnected;
            private DataTable _table;
            private CancellationTokenSource _pollCts;
            private long _ok;
            private long _err;
            private double _rttLast;
            private double _rttAvg;

            public string Header 
            { 
                get => _header; 
                set { _header = value; OnPropertyChanged(nameof(Header)); }
            }
            
            public byte UnitId 
            { 
                get => _unitId; 
                set { _unitId = value; OnPropertyChanged(nameof(UnitId)); }
            }
            
            public TcpClient TcpClient 
            { 
                get => _tcpClient; 
                set { _tcpClient = value; OnPropertyChanged(nameof(TcpClient)); }
            }
            
            public IModbusMaster Master 
            { 
                get => _master; 
                set { _master = value; OnPropertyChanged(nameof(Master)); }
            }
            
            public bool IsConnected 
            { 
                get => _isConnected; 
                set { _isConnected = value; OnPropertyChanged(nameof(IsConnected)); }
            }
            
            public DataTable Table 
            { 
                get => _table; 
                set { _table = value; OnPropertyChanged(nameof(Table)); }
            }
            
            public CancellationTokenSource PollCts 
            { 
                get => _pollCts; 
                set { _pollCts = value; OnPropertyChanged(nameof(PollCts)); }
            }
            
            public long Ok 
            { 
                get => _ok; 
                set { _ok = value; OnPropertyChanged(nameof(Ok)); }
            }
            
            public long Err 
            { 
                get => _err; 
                set { _err = value; OnPropertyChanged(nameof(Err)); }
            }
            
            public double RttLast 
            { 
                get => _rttLast; 
                set { _rttLast = value; OnPropertyChanged(nameof(RttLast)); }
            }
            
            public double RttAvg 
            { 
                get => _rttAvg; 
                set { _rttAvg = value; OnPropertyChanged(nameof(RttAvg)); }
            }

            public event PropertyChangedEventHandler PropertyChanged;
            protected virtual void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        #region Private Fields
        private readonly ObservableCollection<UnitTabData> _units = new ObservableCollection<UnitTabData>();
        private UnitTabData _currentUnit;
        
        private readonly object _sync = new object();
        private readonly Stopwatch _sw = new Stopwatch();
        private readonly SemaphoreSlim _pollGate = new SemaphoreSlim(1, 1);
        private readonly SemaphoreSlim _writeGate = new SemaphoreSlim(1, 1);
        
        // 비트 변경 최적화를 위한 디바운싱
        private readonly ConcurrentDictionary<string, CancellationTokenSource> _writeDelayTasks = new ConcurrentDictionary<string, CancellationTokenSource>();
        private volatile bool _isUpdatingFormats = false;

        // UI Resources from XAML
        private readonly Brush _successBrush;
        private readonly Brush _errorBrush;
        #endregion

        #region Constructor & Initialization
        public MainWindow()
        {
            InitializeComponent();
            InitializeUnits();
            UpdateUiState();
            SetupEventHandlers();

            _successBrush = (Brush)FindResource("SuccessBrush");
            _errorBrush = (Brush)FindResource("ErrorBrush");
        }

        private void InitializeUnits()
        {
            // 기본 Unit 추가
            AddNewUnit();
            
        }

        private void SetupEventHandlers()
        {
            dgData.CellEditEnding += DgData_CellEditEnding;
        }

        private DataTable CreateDataTable()
        {
            var table = new DataTable();
            table.Columns.Add("Address", typeof(string));
            table.Columns.Add("Value", typeof(ushort));
            table.Columns.Add("ValueHex", typeof(string));
            table.Columns.Add("ValueHighChar", typeof(string));
            table.Columns.Add("ValueLowChar", typeof(string));
            table.Columns.Add("ValueString", typeof(string));
            
            for (int i = 0; i < 16; i++)
            {
                table.Columns.Add($"Bit{i}", typeof(bool));
            }
            
            table.ColumnChanged += Table_ColumnChanged;
            return table;
        }
        #endregion

        #region Unit Management
        private void btnAddUnit_Click(object sender, RoutedEventArgs e)
        {
            byte unitId;

            // 입력된 Unit ID 확인
            if (!string.IsNullOrEmpty(txtNewUnitId.Text) && byte.TryParse(txtNewUnitId.Text, out unitId))
            {
                // 중복 확인
                if (_units.Any(u => u.UnitId == unitId))
                {
                    Log($"[!] Unit ID {unitId} already exists.");
                    return;
                }
            }
            else
            {
                // 자동 할당
                unitId = GetNextAvailableUnitId();
            }

            var unit = new UnitTabData
            {
                Header = $"Unit {unitId}",
                UnitId = unitId,
                Table = CreateDataTable()
            };

            _units.Add(unit);
    
            // TabControl 코드 제거하고 새로운 토글 버튼 시스템으로 변경
            // tabUnits.SelectedItem = unit;  // 이 줄 삭제
    
            // 토글 버튼 생성 및 추가
            CreateUnitToggleButton(unit);
    
            // 새 Unit을 현재 선택으로 설정
            SelectUnit(unit);

            // 입력 필드 클리어
            txtNewUnitId.Clear();

            Log($"[+] Added Unit {unitId}");
        }

        private void btnRemoveUnit_Click(object sender, RoutedEventArgs e)
        {
            if (_units.Count <= 1)
            {
                Log("[!] At least one unit must remain.");
                return;
            }

            // TabControl 코드 제거하고 현재 선택된 Unit 사용
            // if (tabUnits.SelectedItem is UnitTabData selectedUnit)  // 이 줄 삭제
    
            if (_currentUnit != null)
            {
                RemoveUnit(_currentUnit);
            }
            else
            {
                Log("[!] No unit selected for removal.");
            }
        }

        private void AddNewUnit()
        {
            byte unitId;
    
            // txtNewUnitId에서 입력값 가져오기 (Title Bar의 입력필드)
            if (!string.IsNullOrEmpty(txtNewUnitId.Text) && byte.TryParse(txtNewUnitId.Text, out unitId))
            {
                if (_units.Any(u => u.UnitId == unitId))
                {
                    Log($"[!] Unit ID {unitId} already exists.");
                    return;
                }
            }
            else
            {
                unitId = GetNextAvailableUnitId();
            }
    
            var unit = new UnitTabData
            {
                Header = $"Unit {unitId}",
                UnitId = unitId,
                Table = CreateDataTable()
            };

            _units.Add(unit);
    
            // 토글 버튼 생성 및 추가
            CreateUnitToggleButton(unit);
    
            // 새 Unit을 현재 선택으로 설정
            SelectUnit(unit);
    
            txtNewUnitId.Clear();
            Log($"[+] Added Unit {unitId}");
        }
        
        private void CreateUnitToggleButton(UnitTabData unit)
        {
            var toggleButton = new ToggleButton
            {
                Content = $"Unit {unit.UnitId}",
                Style = (Style)FindResource("SegmentedToggleButton"),
                Tag = unit, // Unit 데이터 연결
                Margin = new Thickness(0, 0, 4, 0),
                MinWidth = 80,
                Height = 32
            };
    
            // 클릭 이벤트 연결
            toggleButton.Checked += UnitToggle_Checked;
    
            // 연결 상태에 따른 시각적 표시 (선택사항)
            UpdateUnitToggleStatus(toggleButton, unit);
    
            spUnitToggles.Children.Add(toggleButton);
        }
        
        private void UpdateUnitToggleStatus(ToggleButton button, UnitTabData unit)
        {
            // 연결 상태에 따른 시각적 피드백
            if (unit.IsConnected)
            {
                button.Content = $"Unit {unit.UnitId} ●"; // 연결됨 표시
            }
            else
            {
                button.Content = $"Unit {unit.UnitId}";
            }
        }
        
        private void UnitToggle_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleButton toggleButton && toggleButton.Tag is UnitTabData unit)
            {
                // 다른 Unit 토글들 해제
                foreach (ToggleButton btn in spUnitToggles.Children.OfType<ToggleButton>())
                {
                    if (btn != toggleButton)
                        btn.IsChecked = false;
                }
        
                SelectUnit(unit);
            }
        }
        
        private void SelectUnit(UnitTabData unit)
        {
            _currentUnit = unit;
    
            // UI 업데이트
            txtUnitId.Text = unit.UnitId.ToString();
            dgData.ItemsSource = unit.Table?.DefaultView;
    
            UpdateConnectionStatus();
            UpdateStats();
            UpdateUiState();
        }

        private void RemoveUnit(UnitTabData unit)
        {
            // 연결 해제
            if (unit.IsConnected)
            {
                DisconnectUnit(unit);
            }

            // 해당 토글 버튼 제거
            var buttonToRemove = spUnitToggles.Children.OfType<ToggleButton>()
                .FirstOrDefault(btn => btn.Tag == unit);
            if (buttonToRemove != null)
            {
                spUnitToggles.Children.Remove(buttonToRemove);
            }

            _units.Remove(unit);
    
            // 새로운 선택 설정
            if (_units.Count > 0)
            {
                var firstUnit = _units.First();
                SelectUnit(firstUnit);
        
                // 첫 번째 토글 버튼 선택
                var firstButton = spUnitToggles.Children.OfType<ToggleButton>().FirstOrDefault();
                if (firstButton != null)
                    firstButton.IsChecked = true;
            }
    
            Log($"[-] Removed Unit {unit.UnitId}");
        }
        
        private byte GetNextAvailableUnitId()
        {
            var usedIds = _units.Select(u => u.UnitId).ToHashSet();
            for (byte i = 1; i <= 255; i++)
            {
                if (!usedIds.Contains(i))
                    return i;
            }
            return 1;
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

        #region View Toggle Events
        private void btnDataView_Checked(object sender, RoutedEventArgs e)
        {
            if (btnLogView == null || DataMonitorPanel == null || EventLogPanel == null) return;
            btnLogView.IsChecked = false;
            DataMonitorPanel.Visibility = Visibility.Visible;
            EventLogPanel.Visibility = Visibility.Collapsed;
        }

        private void btnLogView_Checked(object sender, RoutedEventArgs e)
        {
            if (btnDataView == null || DataMonitorPanel == null || EventLogPanel == null) return;
            btnDataView.IsChecked = false;
            DataMonitorPanel.Visibility = Visibility.Collapsed;
            EventLogPanel.Visibility = Visibility.Visible;
        }
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
            if (e.EditAction != DataGridEditAction.Commit || _currentUnit?.Table == null) return;
            
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
            var key = $"{_currentUnit?.UnitId}_{address}";
            
            // 기존 지연 작업 취소
            if (_writeDelayTasks.TryRemove(key, out var existingCts))
            {
                existingCts.Cancel();
                existingCts.Dispose();
            }

            // 새로운 지연 작업 생성
            var newCts = new CancellationTokenSource();
            _writeDelayTasks[key] = newCts;

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
                _writeDelayTasks.TryRemove(key, out _);
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
            if (_currentUnit == null) return;
            
            if (_currentUnit.IsConnected) 
                await DisconnectCurrentUnitAsync();
            else 
                await ConnectCurrentUnitAsync();
        }

        private async Task ConnectCurrentUnitAsync()
        {
            if (_currentUnit == null || !ValidateConnectionInputs()) return;

            try
            {
                await DisconnectCurrentUnitAsync();
                
                var ip = txtIp.Text;
                var port = int.Parse(txtPort.Text);
                var unitId = byte.Parse(txtUnitId.Text);
                
                _currentUnit.TcpClient = new TcpClient 
                { 
                    NoDelay = true, 
                    ReceiveTimeout = 2000, 
                    SendTimeout = 2000 
                };
                
                Log($"[*] Unit {unitId}: Connecting to {ip}:{port}...");
                
                var connectTask = _currentUnit.TcpClient.ConnectAsync(ip, port);
                var timeoutTask = Task.Delay(4000);
                
                if (await Task.WhenAny(connectTask, timeoutTask) == timeoutTask)
                {
                    throw new TimeoutException("Connection timeout.");
                }
                
                if (!_currentUnit.TcpClient.Connected)
                {
                    throw new InvalidOperationException("Failed to connect.");
                }
                
                _currentUnit.Master = ModbusIpMaster.CreateIp(_currentUnit.TcpClient);
                _currentUnit.Master.Transport.ReadTimeout = 1500;
                _currentUnit.Master.Transport.WriteTimeout = 1500;
                _currentUnit.UnitId = unitId;
                _currentUnit.IsConnected = true;
                
                Log($"[+] Unit {unitId}: Connected successfully.");
                UpdateConnectionStatus();
            }
            catch (Exception ex)
            {
                if (_currentUnit != null)
                    _currentUnit.IsConnected = false;
                LogError($"Unit {_currentUnit?.UnitId}: Connection failed", ex);
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

        private async Task DisconnectCurrentUnitAsync()
        {
            if (_currentUnit == null) return;
            await DisconnectUnit(_currentUnit);
        }

        private async Task DisconnectUnit(UnitTabData unit)
        {
            if (unit == null) return;
            
            StopPollingForUnit(unit);
            await CancelAllDelayedWritesForUnit(unit);
            
            unit.Master?.Dispose();
            unit.Master = null;
            unit.TcpClient?.Dispose();
            unit.TcpClient = null;
            
            if (unit.IsConnected)
            {
                unit.IsConnected = false;
                Log($"[-] Unit {unit.UnitId}: Disconnected.");
            }
            
            if (unit == _currentUnit)
            {
                UpdateConnectionStatus();
                UpdateUiState();
            }
        }

        private async Task CancelAllDelayedWritesForUnit(UnitTabData unit)
        {
            var unitPrefix = $"{unit.UnitId}_";
            var tasks = _writeDelayTasks.Where(kvp => kvp.Key.StartsWith(unitPrefix)).ToList();
            
            foreach (var kvp in tasks)
            {
                kvp.Value.Cancel();
                _writeDelayTasks.TryRemove(kvp.Key, out _);
            }
            
            // 짧은 대기로 진행 중인 작업들이 완료되도록 함
            await Task.Delay(50);
        }

        private void UpdateConnectionStatus()
        {
            var isConnected = _currentUnit?.IsConnected ?? false;
            SetConnIndicator(isConnected, isConnected ? $"Unit {_currentUnit.UnitId} Connected" : "Disconnected");
        }
        #endregion

        #region Read Operations
        private async void btnReadOnce_Click(object sender, RoutedEventArgs e) => await ReadOnceAsync();

        private async Task ReadOnceAsync()
        {
            if (_currentUnit?.IsConnected != true || !ValidateReadInputs()) return;

            var start = ushort.Parse(txtStart.Text);
            var count = ushort.Parse(txtCount.Text);

            await _pollGate.WaitAsync();
            try
            {
                _sw.Restart();
                
                var data = await ReadModbusData(_currentUnit, start, count);
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

        private async Task<Array> ReadModbusData(UnitTabData unit, ushort start, ushort count)
        {
            switch (cmbFunction.SelectedIndex)
            {
                case 0:
                    return await unit.Master.ReadCoilsAsync(unit.UnitId, start, count);
                case 1:
                    return await unit.Master.ReadInputsAsync(unit.UnitId, start, count);
                case 2:
                    return await unit.Master.ReadHoldingRegistersAsync(unit.UnitId, start, count);
                case 3:
                    return await unit.Master.ReadInputRegistersAsync(unit.UnitId, start, count);
                default:
                    throw new InvalidOperationException("Invalid function selection");
            }
        }

        private void ApplyReadData(ushort start, Array data)
        {
            if (_currentUnit?.Table == null) return;
            
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
            
                    _currentUnit.Table.Rows[i]["Value"] = value;
                    UpdateRowFormats(_currentUnit.Table.Rows[i], value);
                }
            }
        }

        private void EnsureRows(ushort start, ushort length)
        {
            if (_currentUnit?.Table == null) return;
            
            if (_currentUnit.Table.Rows.Count != length)
            {
                _currentUnit.Table.Clear();
                for (ushort i = 0; i < length; i++)
                {
                    var row = _currentUnit.Table.NewRow();
                    row["Address"] = $"{start + i}";
                    _currentUnit.Table.Rows.Add(row);
                }
            }
        }

        private void UpdateRowFormats(DataRow row, ushort value)
        {
            if (_isUpdatingFormats) return;
            
            _isUpdatingFormats = true;
            try
            {
                _currentUnit.Table.ColumnChanged -= Table_ColumnChanged;
                UpdateRowDisplayFormats(row, value);
                UpdateRowBits(row, value);
                _currentUnit.Table.ColumnChanged += Table_ColumnChanged;
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
            if (_currentUnit?.PollCts != null)
            {
                Log($"[!] Unit {_currentUnit.UnitId}: Polling is already running.");
                return;
            }
            
            if (!int.TryParse(txtPeriod.Text, out var period) || period < 50)
            {
                Log("[!] Period must be >= 50ms.");
                return;
            }
            
            _currentUnit.PollCts = new CancellationTokenSource();
            Log($"[*] Unit {_currentUnit.UnitId}: Polling started (period: {period}ms).");
            UpdateUiState();
            
            await Task.Run(() => PollLoop(period, _currentUnit.PollCts.Token), _currentUnit.PollCts.Token);
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
            if (_currentUnit == null) return;
            StopPollingForUnit(_currentUnit);
        }

        private void StopPollingForUnit(UnitTabData unit)
        {
            if (unit?.PollCts != null)
            {
                unit.PollCts.Cancel();
                unit.PollCts.Dispose();
                unit.PollCts = null;
                Log($"[*] Unit {unit.UnitId}: Polling stopped.");
                
                if (unit == _currentUnit)
                    UpdateUiState();
            }
        }
        #endregion

        #region Write Operations
        private async void btnWrite_Click(object sender, RoutedEventArgs e)
        {
            if (_currentUnit?.IsConnected != true || !ValidateWriteInputs()) return;

            var addr = ushort.Parse(txtWriteAddr.Text);

            try
            {
                await ExecuteWriteOperation(_currentUnit, addr);
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

        private async Task ExecuteWriteOperation(UnitTabData unit, ushort addr)
        {
            switch (cmbWriteType.SelectedIndex)
            {
                case 0: // Coil
                    if (TryParseBool(txtWriteValue.Text, out var coilVal))
                    {
                        await unit.Master.WriteSingleCoilAsync(unit.UnitId, addr, coilVal);
                        Log($"[+] Unit {unit.UnitId}: Wrote coil {addr} = {coilVal}");
                    }
                    else
                    {
                        Log("[!] Coil value must be 0/1 or true/false.");
                    }
                    break;
                    
                case 1: // Single Register
                    if (ushort.TryParse(txtWriteValue.Text, out var regVal))
                    {
                        await unit.Master.WriteSingleRegisterAsync(unit.UnitId, addr, regVal);
                        Log($"[+] Unit {unit.UnitId}: Wrote register {addr} = {regVal}");
                    }
                    else
                    {
                        Log("[!] Register value must be 0-65535.");
                    }
                    break;
                    
                case 2: // Multiple Registers
                    await WriteMultipleRegisters(unit, addr);
                    break;
            }
        }

        private async Task WriteRegisterFromRow(DataRow row)
        {
            if (_currentUnit?.IsConnected != true || cmbFunction.SelectedIndex != 2) return;
            if (!ushort.TryParse(row["Address"].ToString(), out var addr)) return;

            var value = (ushort)row["Value"];
            
            await _writeGate.WaitAsync();
            try
            {
                await _currentUnit.Master.WriteSingleRegisterAsync(_currentUnit.UnitId, addr, value);
                Log($"[+] Unit {_currentUnit.UnitId}: Auto-Wrote Register {addr} = {value}");
            }
            catch (Exception ex)
            {
                LogError($"Unit {_currentUnit.UnitId}: Auto-Write failed for address {addr}", ex);
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

        private async Task WriteMultipleRegisters(UnitTabData unit, ushort startAddr)
        {
            if (dgData.SelectedItems.Count == 0)
            {
                Log("[!] No rows selected for multiple write.");
                return;
            }
            
            var values = dgData.SelectedItems.Cast<DataRowView>()
                .Select(item => (ushort)item.Row["Value"])
                .ToArray();
                
            await unit.Master.WriteMultipleRegistersAsync(unit.UnitId, startAddr, values);
            Log($"[+] Unit {unit.UnitId}: Wrote {values.Length} registers from {startAddr}.");
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
            if (_currentUnit == null) return;
            
            _currentUnit.RttLast = ms;
            _currentUnit.Ok++;
            _currentUnit.RttAvg = (_currentUnit.RttAvg * (_currentUnit.Ok - 1) + ms) / _currentUnit.Ok;
            UpdateStats();
        }

        private void OnError(double ms, Exception ex)
        {
            if (_currentUnit == null) return;
            
            _currentUnit.RttLast = ms;
            _currentUnit.Err++;
            UpdateStats();
            LogError($"Unit {_currentUnit.UnitId}: Read error (Connection may be lost)", ex);
        }

        private void UpdateStats()
        {
            if (_currentUnit == null) return;
            
            Dispatcher.Invoke(() =>
            {
                lblRttLast.Text = _currentUnit.RttLast.ToString("F1");
                lblRttAvg.Text = _currentUnit.Ok > 0 ? _currentUnit.RttAvg.ToString("F1") : "-";
                lblOk.Text = _currentUnit.Ok.ToString();
                lblErr.Text = _currentUnit.Err.ToString();
            });
        }

        private void UpdateUiState()
        {
            Dispatcher.Invoke(() =>
            {
                bool isConnected = _currentUnit?.IsConnected ?? false;
                bool isPolling = _currentUnit?.PollCts != null;

                btnToggleConnect.Content = isConnected ? "Disconnect" : "Connect";
                btnToggleConnect.Style = isConnected ?
                    (Style)FindResource("DisconnectButton") :
                    (Style)FindResource("PrimaryButton");

                bool canInteract = isConnected && !isPolling;

                btnReadOnce.IsEnabled = canInteract;
                btnStart.IsEnabled = isConnected && !isPolling;
                btnStop.IsEnabled = isPolling;
                btnWrite.IsEnabled = isConnected;

                txtIp.IsEnabled = !isConnected;
                txtPort.IsEnabled = !isConnected;
                txtUnitId.IsEnabled = !isConnected;
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
            if (_currentUnit?.IsConnected == true) return;
            
            Log($"[*] Unit {_currentUnit?.UnitId}: Connection lost. Attempting auto-reconnect...");
            await DisconnectCurrentUnitAsync();
            await Task.Delay(1000);
            await ConnectCurrentUnitAsync();
        }
        #endregion

        #region Export & UI Actions
        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_currentUnit?.Table == null)
            {
                Log("[!] No data to export.");
                return;
            }
            
            try
            {
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ModbusMaster");
                Directory.CreateDirectory(dir);
                var file = Path.Combine(dir, $"snapshot_Unit{_currentUnit.UnitId}_{DateTime.Now:yyyyMMdd_HHmmss}.csv");
                
                using (var sw = new StreamWriter(file, false, Encoding.UTF8))
                {
                    // CSV 헤더
                    var headers = new List<string> { "Address", "Dec", "Hex", "HighChar", "LowChar", "String" };
                    headers.AddRange(Enumerable.Range(0, 16).Reverse().Select(i => $"Bit{i}"));
                    sw.WriteLine(string.Join(",", headers));
                    
                    // 데이터 행
                    foreach (DataRow row in _currentUnit.Table.Rows)
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
                }
                
                Log($"[+] Unit {_currentUnit.UnitId}: Exported data to: {file}");
            }
            catch (Exception ex)
            {
                LogError("Export failed", ex);
            }
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            if (_currentUnit?.Table == null) return;
            
            _currentUnit.Table.Clear();
            _currentUnit.Ok = _currentUnit.Err = 0;
            _currentUnit.RttLast = _currentUnit.RttAvg = 0;
            UpdateStats();
            Log($"[*] Unit {_currentUnit.UnitId}: Data grid and stats cleared.");
        }

        private void btnCopyLog_Click(object sender, RoutedEventArgs e) => 
            Clipboard.SetText(txtLog.Text);

        private void btnClearLog_Click(object sender, RoutedEventArgs e) => 
            txtLog.Clear();
        #endregion

        #region Cleanup
        protected override async void OnClosing(CancelEventArgs e)
        {
            // 모든 유닛 정리
            foreach (var unit in _units.ToList())
            {
                await DisconnectUnit(unit);
            }
            
            // 모든 지연된 쓰기 작업 취소
            foreach (var cts in _writeDelayTasks.Values)
            {
                cts.Cancel();
                cts.Dispose();
            }
            _writeDelayTasks.Clear();
            
            base.OnClosing(e);
        }
        #endregion
    }
}