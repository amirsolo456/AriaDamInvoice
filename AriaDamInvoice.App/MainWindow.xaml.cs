using System.IO;
using System.Windows;
using AriaDamInvoice.App.Models;
using AriaDamInvoice.App.Services;
using Microsoft.Win32;

namespace AriaDamInvoice.App;

public partial class MainWindow : Window
{
    private readonly SettingsService _settingsService = new();
    private AppSettings _settings;
    private string _selectedFolder = string.Empty;

    public MainWindow()
    {
        InitializeComponent();
        _settings = _settingsService.Load();

        if (Directory.Exists(_settings.OutputDirectory))
            ShowDashboard();
        else
            ShowFolderSetup();
    }

    private void ShowFolderSetup()
    {
        DashboardPanel.Visibility = Visibility.Collapsed;
        SetupPanel.Visibility = Visibility.Visible;
        SelectedFolderText.Text = string.Empty;
        SaveFolderButton.IsEnabled = false;
        _selectedFolder = string.Empty;
    }

    private void ShowDashboard()
    {
        SetupPanel.Visibility = Visibility.Collapsed;
        DashboardPanel.Visibility = Visibility.Visible;
        OutputFolderText.Text = _settings.OutputDirectory;
    }

    private void SelectOutputFolder()
    {
        var dialog = new OpenFolderDialog
        {
            Title = "پوشه محل ذخیره قراردادها را انتخاب کنید",
            Multiselect = false
        };

        if (Directory.Exists(_settings.OutputDirectory))
            dialog.InitialDirectory = _settings.OutputDirectory;

        if (dialog.ShowDialog(this) != true)
            return;

        _selectedFolder = dialog.FolderName;
        SelectedFolderText.Text = _selectedFolder;
        SaveFolderButton.IsEnabled = Directory.Exists(_selectedFolder);
    }

    private void ChooseFolder_Click(object sender, RoutedEventArgs e) => SelectOutputFolder();

    private void SaveFolder_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_selectedFolder) || !Directory.Exists(_selectedFolder))
        {
            MessageBox.Show("ابتدا یک پوشه معتبر انتخاب کنید.", "پوشه ذخیره‌سازی",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SaveSelectedFolderAndShowDashboard();
    }

    private void ChangeFolder_Click(object sender, RoutedEventArgs e)
    {
        SelectOutputFolder();
        if (!string.IsNullOrWhiteSpace(_selectedFolder))
            SaveSelectedFolderAndShowDashboard();
    }

    private void SaveSelectedFolderAndShowDashboard()
    {
        try
        {
            _settings.OutputDirectory = _selectedFolder;
            _settingsService.Save(_settings);
            ShowDashboard();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            MessageBox.Show($"ذخیره محل خروجی انجام نشد.\n{ex.Message}", "خطا",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void NewInvoice_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show("بخش صدور قرارداد در مرحله بعدی توسعه می‌شود.", "صدور قرارداد جدید",
            MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void PreviousInvoices_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show("بخش مشاهده و ویرایش قراردادهای قبلی در مرحله بعدی توسعه می‌شود.", "قراردادهای قبلی",
            MessageBoxButton.OK, MessageBoxImage.Information);
    }
}
