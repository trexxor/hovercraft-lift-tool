using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace HovercraftLiftTool;

// Converts a bool (IsEnabled / not solved-for) to a field background brush.
// true  (field is enabled / user-input) → white
// false (field is disabled / solved-for) → light gray (#F2F3F4)
[ValueConversion(typeof(bool), typeof(Brush))]
public class BoolToFieldBrushConverter : IValueConverter
{
    public static readonly BoolToFieldBrushConverter Instance = new();

    private static readonly SolidColorBrush InputBrush  = MakeBrush(0xFF, 0xFF, 0xFF);
    private static readonly SolidColorBrush CalcBrush   = MakeBrush(0xF2, 0xF3, 0xF4);

    private static SolidColorBrush MakeBrush(byte r, byte g, byte b)
    {
        var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
        brush.Freeze();
        return brush;
    }

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is bool b && b ? InputBrush : CalcBrush;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => DependencyProperty.UnsetValue;
}

// Rounds a double to a fixed number of decimal places for display.
// Used as: Text="{Binding Val, Converter={StaticResource Fmt4}}"
[ValueConversion(typeof(double), typeof(string))]
public class RoundingConverter : IValueConverter
{
    public int Decimals { get; set; } = 4;

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is double d)
            return d.ToString("F" + Decimals, culture);
        return value?.ToString() ?? "";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string s && double.TryParse(s, NumberStyles.Any, culture, out double result))
            return result;
        return 0.0;
    }
}
