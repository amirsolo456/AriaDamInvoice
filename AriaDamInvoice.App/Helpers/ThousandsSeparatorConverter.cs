using System.Globalization;
using System.Windows.Data;

namespace AriaDamInvoice.Helpers;

public sealed class ThousandsSeparatorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is null)
            return string.Empty;

        if (decimal.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
            return number.ToString("#,##0", CultureInfo.InvariantCulture);

        return value.ToString() ?? string.Empty;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var text = (value?.ToString() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(text))
            return targetType == typeof(decimal?) ? null! : 0m;

        const string persianDigits = "۰۱۲۳۴۵۶۷۸۹";
        const string arabicDigits = "٠١٢٣٤٥٦٧٨٩";
        var normalized = new string(text.Select(ch =>
        {
            var p = persianDigits.IndexOf(ch);
            if (p >= 0) return (char)('0' + p);
            var a = arabicDigits.IndexOf(ch);
            if (a >= 0) return (char)('0' + a);
            return ch;
        }).Where(ch => ch != ',' && ch != '٬' && !char.IsWhiteSpace(ch)).ToArray());

        if (decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
            return result;

        return 0m;
    }
}