using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using AriaDamInvoice.Models;

namespace AriaDamInvoice.Views;

public partial class InvoiceWindow : Window
{
    private readonly ObservableCollection<AnimalLine> _animals = new();

    public InvoiceWindow()
    {
        InitializeComponent();
        ContractDatePicker.SelectedDate = DateTime.Today;
        AnimalsGrid.ItemsSource = _animals;
        AddAnimalRow();
        EarTagsTextBox.TextChanged += (_, _) => UpdateEarTagCount();
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
        MessageBox.Show(
            "فرم ورود اطلاعات آماده است. اتصال ذخیره‌سازی اکسل و تولید PDF در مرحله بعد انجام می‌شود.",
            "اطلاعات قرارداد",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }
}