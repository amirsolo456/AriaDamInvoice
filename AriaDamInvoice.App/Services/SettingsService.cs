using System.IO;
using System.Text.Json;
using AriaDamInvoice.Models;

namespace AriaDamInvoice.Services;

public sealed class SettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly string _settingsDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "AriaDamInvoice");

    private string SettingsFilePath => Path.Combine(_settingsDirectory, "settings.json");

    public AppSettings Load()
    {
        try
        {
            if (!File.Exists(SettingsFilePath))
                return new AppSettings();

            var json = File.ReadAllText(SettingsFilePath);
            return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch (Exception ex) when (
            ex is IOException or UnauthorizedAccessException or JsonException)
        {
            return new AppSettings();
        }
    }

    public void Save(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        Directory.CreateDirectory(_settingsDirectory);

        var temporaryPath = SettingsFilePath + ".tmp";
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(temporaryPath, json);

        if (File.Exists(SettingsFilePath))
            File.Replace(temporaryPath, SettingsFilePath, null);
        else
            File.Move(temporaryPath, SettingsFilePath);
    }
}
