using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using AriaDamInvoice.Models;
using AriaDamInvoice.Services;

namespace AriaDamInvoice.Views;

public partial class InvoiceWindow : Window
{
    private readonly ObservableCollection<AnimalLine> _animals = new();
    private readonly ObservableCollection<string> _breedOptions = new();
    private readonly PersianCalendar _persianCalendar = new();
    private readonly SettingsService _settingsService = new();
    private readonly AppSettings _settings;

    public ObservableCollection<string> BreedOptions => _breedOptions;

    public InvoiceWindow()
    {
        InitializeComponent();

        _settings = _settingsService.Load();
        if (_settings.BreedOptions is null || _settings.BreedOptions.Count == 0)
            _settings.BreedOptions = new List<string> { "همو", "هترو", "افشاری" };

        foreach (var breed in _settings.BreedOptions
                     .Where(x => !string.IsNullOrWhiteSpace(x))
                     .Distinct(StringComparer.CurrentCultureIgnoreCase))
            _breedOptions.Add(breed);

        DataContext = this;
        ContractDateTextBox.Text = FormatPersianDate(DateTime.Today);
        if (_settings.NextDocumentNumber < 1000) _settings.NextDocumentNumber = 1000;
        DocumentNumberTextBox.Text = _settings.NextDocumentNumber.ToString(CultureInfo.InvariantCulture);
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

    private bool TryGetDocumentNumber(out int documentNumber)
    {
        var normalized = NormalizeDigits(DocumentNumberTextBox.Text);
        return int.TryParse(normalized, NumberStyles.None, CultureInfo.InvariantCulture, out documentNumber)
               && documentNumber >= 1;
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

    private void AddBreed_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Window
        {
            Title = "افزودن نژاد جدید",
            Width = 360,
            Height = 190,
            ResizeMode = ResizeMode.NoResize,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            FlowDirection = FlowDirection.RightToLeft,
            FontFamily = new FontFamily("B Nazanin"),
            Background = Brushes.White
        };

        var panel = new StackPanel { Margin = new Thickness(18) };
        panel.Children.Add(new TextBlock
        {
            Text = "نام گزینه جدید را وارد کنید:",
            FontSize = 17,
            Margin = new Thickness(0, 0, 0, 8)
        });

        var input = new TextBox
        {
            FontFamily = new FontFamily("B Nazanin"),
            FontSize = 17,
            Padding = new Thickness(7)
        };
        panel.Children.Add(input);

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Left,
            Margin = new Thickness(0, 14, 0, 0)
        };

        var addButton = new Button
        {
            Content = "افزودن",
            MinWidth = 90,
            Margin = new Thickness(0, 0, 8, 0),
            IsDefault = true,
            Background = new SolidColorBrush(Color.FromRgb(23, 107, 87)),
            Foreground = Brushes.White
        };
        var cancelButton = new Button
        {
            Content = "انصراف",
            MinWidth = 90,
            IsCancel = true
        };

        addButton.Click += (_, _) =>
        {
            var value = input.Text.Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                MessageBox.Show(dialog, "نام گزینه را وارد کنید.", "افزودن نژاد",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_breedOptions.Any(x => string.Equals(x, value, StringComparison.CurrentCultureIgnoreCase)))
            {
                MessageBox.Show(dialog, "این گزینه قبلاً ثبت شده است.", "افزودن نژاد",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _breedOptions.Add(value);
            _settings.BreedOptions = _breedOptions.ToList();

            try
            {
                _settingsService.Save(_settings);
                dialog.DialogResult = true;
                dialog.Close();
            }
            catch (Exception ex)
            {
                _breedOptions.Remove(value);
                MessageBox.Show(dialog, $"ذخیره گزینه انجام نشد.\n{ex.Message}", "خطا",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        };

        buttons.Children.Add(addButton);
        buttons.Children.Add(cancelButton);
        panel.Children.Add(buttons);
        dialog.Content = panel;
        dialog.Loaded += (_, _) => input.Focus();
        dialog.ShowDialog();
    }

    private void UpdateGrandTotal()
    {
        GrandTotalText.Inlines.Clear();
        GrandTotalText.Inlines.Add(new Run("جمع کل: ") { FontFamily = new FontFamily("B Nazanin") });
        GrandTotalText.Inlines.Add(new Run(_animals.Sum(x => x.LineTotal).ToString("N0"))
        {
            FontFamily = new FontFamily("B Yekan")
        });
    }

    private void UpdateEarTagCount()
    {
        var count = EarTagsTextBox.Text
            .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Count(x => x.Length > 0);

        EarTagCountText.Inlines.Clear();
        EarTagCountText.Inlines.Add(new Run("تعداد پلاک: ") { FontFamily = new FontFamily("B Nazanin") });
        EarTagCountText.Inlines.Add(new Run(count.ToString("N0"))
        {
            FontFamily = new FontFamily("B Yekan")
        });
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => Close();

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetDocumentNumber(out var documentNumber))
        {
            MessageBox.Show("شماره سند باید یک عدد صحیح مثبت باشد.", "بررسی شماره سند",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            DocumentNumberTextBox.Focus();
            DocumentNumberTextBox.SelectAll();
            return;
        }

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

        AnimalsGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        AnimalsGrid.CommitEdit(DataGridEditingUnit.Row, true);

        var animals = _animals
            .Where(x => x.Count > 0 || x.Weight != 0 ||
                        !string.IsNullOrWhiteSpace(x.AnimalType) ||
                        !string.IsNullOrWhiteSpace(x.Breed) ||
                        x.UnitPrice != 0 ||
                        !string.IsNullOrWhiteSpace(x.Description))
            .ToList();

        if (animals.Count == 0)
        {
            MessageBox.Show(
                "حداقل مشخصات یک ردیف دام را وارد کنید.",
                "اطلاعات قرارداد",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            AnimalsGrid.Focus();
            return;
        }

        var tags = EarTagsTextBox.Text
            .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Where(x => x.Length > 0)
            .ToList();

        try
        {
            var data = new InvoiceData
            {
                ContractDate = FormatPersianDate(contractDate),
                DocumentNumber = documentNumber,
                BuyerName = BuyerNameTextBox.Text.Trim(),
                Phone = PhoneTextBox.Text.Trim(),
                Address = AddressTextBox.Text.Trim(),
                Animals = animals,
                EarTags = tags
            };

            var service = new ExcelContractService(_settings.TemplateFilePath);
            var result = service.Create(data, _settings.OutputDirectory);

            _settings.NextDocumentNumber = checked(documentNumber + 1);
            _settingsService.Save(_settings);

            MessageBox.Show(
                $"قرارداد با موفقیت ذخیره شد.\n\nفایل اکسل:\n{result.ExcelPath}\n\nفایل PDF:\n{result.PdfPath}",
                "صدور قرارداد",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"صدور قرارداد با خطا مواجه شد:\n{ex.Message}",
                "خطا در صدور قرارداد",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }
}