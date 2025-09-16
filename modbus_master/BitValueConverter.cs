using System;
using System.Globalization;
using System.Windows.Data;

namespace modbus_master
{
    public class BitValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ushort val && parameter is string bitStr && int.TryParse(bitStr, out int bit))
            {
                return (val & (1 << bit)) != 0;
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Note: For ConvertBack to work correctly, the binding must be on a property that can provide the original ushort value.
            // Since we are binding directly to the 'Value' in a DataRowView, we can't get the original value easily here.
            // The proper solution is to use a MultiBinding with a converter that gets both the new bool value and the original DataRowView.
            // However, to avoid further complexity, we will handle the update in a different way, likely back in the UI logic for now.
            // This implementation will not work as intended for TwoWay binding on its own.
            throw new NotImplementedException();
        }
    }
}