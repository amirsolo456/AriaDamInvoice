using System.IO;
using System.Text.Json;
using AriaDamInvoice.App.Models;

namespace AriaDamInvoice.App.Services;

public sealed class SettingsService
{
    private readonly string _settingsDirectory =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AriaDamInvoice");

    private string SettingsFilePath =>
        Path.Combine(_settingsDirectory, "settings.json");

    public AppSettings Load()
    {
        try
        {
            if (!File.Exists(SettingsFilePath))
                return new AppSettings();

            var json = File.ReadAllText(SettingsFilePath);
            return JsonSerializer.Deserialize<AppSettings>(json)
                   ?? new AppSettings();
        }
        catch (JsonException)
        {
            // A damaged settings file should not prevent the app from starting.
            return new AppSettings();
        }
        catch (IOException)
        {
            return new AppSettings();
        }
        catch (UnauthorizedAccessException)
        {
            return new AppSettings();
        }
    }

    public void Save(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        Directory.CreateDirectory(_settingsDirectory);

        var json = JsonSerializer.Serialize(
            settings,
            new JsonSerializerOptions { WriteIndented = true });

        var tempFile = SettingsFilePath + ".tmp";
        File.WriteAllText(tempFile, json);

        // Replace atomically when a previous settings file exists.
        if (File.Exists(SettingsFilePath))
            File.Replace(tempFile, SettingsFilePath, null);
        else
            File.Move(tempFile, SettingsFilePath);
    }
}
