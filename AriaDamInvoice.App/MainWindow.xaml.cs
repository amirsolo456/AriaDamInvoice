using System.IO;
using System.Windows;
using AriaDamInvoice.Models;
using AriaDamInvoice.Services;
using AriaDamInvoice.Views;
using Microsoft.Win32;

namespace AriaDamInvoice;

public partial class MainWindow : Window
{
    private readonly SettingsService _settingsService = new();
    private AppSettings _settings;
    private string _selectedFolder = string.Empty;
    private string _selectedTemplate = string.Empty;

    public MainWindow()
    {
        InitializeComponent();

        _settings = _settingsService.Load();

        if (Directory.Exists(_settings.OutputDirectory) && File.Exists(_settings.TemplateFilePath))
            ShowDashboard();
        else
            ShowFolderSetup();
    }

    private void ShowFolderSetup()
    {
        DashboardPanel.Visibility = Visibility.Collapsed;
        SetupPanel.Visibility = Visibility.Visible;
        SelectedFolderText.Text = string.Empty;
        SelectedTemplateText.Text = string.Empty;
        SaveFolderButton.IsEnabled = false;
        _selectedFolder = string.Empty;
    }

    private void ShowDashboard()
    {
        SetupPanel.Visibility = Visibility.Collapsed;
        DashboardPanel.Visibility = Visibility.Visible;
        OutputFolderText.Text = _settings.OutputDirectory;
        TemplateFileText.Text = _settings.TemplateFilePath;
    }

    private void SelectTemplateFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = "فایل اکسل مرجع قرارداد را انتخاب کنید",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            CheckFileExists = true,
            Multiselect = false
        };

        if (File.Exists(_settings.TemplateFilePath))
            dialog.InitialDirectory = Path.GetDirectoryName(_settings.TemplateFilePath);

        if (dialog.ShowDialog(this) != true)
            return;

        _selectedTemplate = dialog.FileName;
        SelectedTemplateText.Text = _selectedTemplate;
        UpdateSetupSaveButton();
    }

    private void UpdateSetupSaveButton() =>
        SaveFolderButton.IsEnabled = Directory.Exists(_selectedFolder) && File.Exists(_selectedTemplate);

    private void ChooseTemplate_Click(object sender, RoutedEventArgs e) => SelectTemplateFile();

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
        UpdateSetupSaveButton();
    }

    private void ChooseFolder_Click(object sender, RoutedEventArgs e) =>
        SelectOutputFolder();

    private void ChangeTemplate_Click(object sender, RoutedEventArgs e)
    {
        SelectTemplateFile();
        if (string.IsNullOrWhiteSpace(_selectedTemplate)) return;
        try
        {
            _settings.TemplateFilePath = _selectedTemplate;
            _settingsService.Save(_settings);
            TemplateFileText.Text = _settings.TemplateFilePath;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"ذخیره مسیر فایل مرجع انجام نشد.\\n{ex.Message}", "خطا", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SaveFolder_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_selectedFolder) ||
            !Directory.Exists(_selectedFolder) || string.IsNullOrWhiteSpace(_selectedTemplate) || !File.Exists(_selectedTemplate))
        {
            MessageBox.Show(
                "ابتدا پوشه ذخیره‌سازی و فایل اکسل مرجع را انتخاب کنید.",
                "پوشه ذخیره‌سازی",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
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
            _settings.TemplateFilePath = _selectedTemplate;
            _settingsService.Save(_settings);
            ShowDashboard();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            MessageBox.Show(
                $"ذخیره محل خروجی انجام نشد.\n{ex.Message}",
                "خطا",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void NewInvoice_Click(object sender, RoutedEventArgs e)
    {
        var window = new InvoiceWindow
        {
            Owner = this
        };

        window.ShowDialog();
    }

    private void PreviousInvoices_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(
            "بخش مشاهده و ویرایش قراردادهای قبلی در مرحله بعدی توسعه می‌شود.",
            "قراردادهای قبلی",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }
}
