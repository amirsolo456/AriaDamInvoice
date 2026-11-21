using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using AriaDamInvoice.Models;

namespace AriaDamInvoice.Views;

public partial class InvoiceWindow : Window
{
    private readonly ObservableCollection<AnimalLine> _animals = new();
    private readonly PersianCalendar _persianCalendar = new();

    public InvoiceWindow()
    {
        InitializeComponent();
        ContractDateTextBox.Text = FormatPersianDate(DateTime.Today);
        AnimalsGrid.ItemsSource = _animals;
        AddAnimalRow();
        EarTagsTextBox.TextChanged += (_, _) => UpdateEarTagCount();
    }

    private string FormatPersianDate(DateTime date) =>
        $"{_persianCalendar.GetYear(date):0000}/{_persianCalendar.GetMonth(date):00}/{_persianCalendar.GetDayOfMonth(date):00}";

    private static string NormalizeDigits(string value)
    {
        const string persianDigits = "۰۱۲۳۴۵۶۷۸۹";
        const string arabicDigits = "٠١٢٣٤٥٦٧٨٩";
        var chars = value.Trim().ToCharArray();

        for (var i = 0; i < chars.Length; i++)
        {
            var persianIndex = persianDigits.IndexOf(chars[i]);
            if (persianIndex >= 0)
            {
                chars[i] = (char)('0' + persianIndex);
                continue;
            }

            var arabicIndex = arabicDigits.IndexOf(chars[i]);
            if (arabicIndex >= 0)
                chars[i] = (char)('0' + arabicIndex);
        }

        return new string(chars);
    }

    private bool TryGetContractDate(out DateTime gregorianDate)
    {
        gregorianDate = default;
        var parts = NormalizeDigits(ContractDateTextBox.Text)
            .Split(new[] { '/', '-', '.' }, StringSplitOptions.TrimEntries);

        if (parts.Length != 3 ||
            !int.TryParse(parts[0], out var year) ||
            !int.TryParse(parts[1], out var month) ||
            !int.TryParse(parts[2], out var day))
            return false;

        try
        {
            gregorianDate = _persianCalendar.ToDateTime(year, month, day, 0, 0, 0, 0);
            return true;
        }
        catch (ArgumentOutOfRangeException)
        {
            return false;
        }
    }

    private void AddAnimalRow()
    {
        var row = new AnimalLine { RowNumber = _animals.Count + 1 };
        row.PropertyChanged += (_, _) => UpdateGrandTotal();
        _animals.Add(row);
        UpdateGrandTotal();
    }

    private void AddAnimal_Click(object sender, RoutedEventArgs e) => AddAnimalRow();

    private void RemoveAnimal_Click(object sender, RoutedEventArgs e)
    {
        if (AnimalsGrid.SelectedItem is not AnimalLine selected)
            return;

        _animals.Remove(selected);
        for (var i = 0; i < _animals.Count; i++)
            _animals[i].RowNumber = i + 1;

        UpdateGrandTotal();
    }

    private void UpdateGrandTotal()
    {
        GrandTotalText.Text = "جمع کل: " + _animals.Sum(x => x.LineTotal).ToString("N0");
    }

    private void UpdateEarTagCount()
    {
        var count = EarTagsTextBox.Text
            .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Count(x => x.Length > 0);

        EarTagCountText.Text = "تعداد پلاک: " + count.ToString("N0");
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => Close();

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetContractDate(out var contractDate))
        {
            MessageBox.Show(
                "تاریخ شمسی معتبر نیست. تاریخ را به شکل سال/ماه/روز وارد کنید؛ مانند ۱۴۰۵/۰۸/۳۰.",
                "بررسی تاریخ قرارداد",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            ContractDateTextBox.Focus();
            ContractDateTextBox.SelectAll();
            return;
        }

        MessageBox.Show(
            $"تاریخ شمسی قرارداد: {FormatPersianDate(contractDate)}\nاتصال ذخیره‌سازی اکسل و تولید PDF در مرحله بعد انجام می‌شود.",
            "اطلاعات قرارداد",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }
}