namespace AriaDamInvoice.Models;

public sealed class AppSettings
{
    public string OutputDirectory { get; set; } = string.Empty;

    public string TemplateFilePath { get; set; } = string.Empty;

    public int NextDocumentNumber { get; set; } = 1000;

    public List<string> BreedOptions { get; set; } = new()
    {
        "همو",
        "هترو",
        "افشاری"
    };
}