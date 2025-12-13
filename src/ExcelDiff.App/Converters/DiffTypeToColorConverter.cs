using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using ExcelDiff.Core.Enums;

namespace ExcelDiff.App.Converters;

public class DiffTypeToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is DiffType diffType)
        {
            return diffType switch
            {
                DiffType.Added => new SolidColorBrush(Color.FromRgb(144, 238, 144)), // LightGreen
                DiffType.Deleted => new SolidColorBrush(Color.FromRgb(240, 128, 128)), // LightCoral
                DiffType.Modified => new SolidColorBrush(Color.FromRgb(255, 255, 0)), // Yellow
                _ => new SolidColorBrush(Colors.White)
            };
        }

        return new SolidColorBrush(Colors.White);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
