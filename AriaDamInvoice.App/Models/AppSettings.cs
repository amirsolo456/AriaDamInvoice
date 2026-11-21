namespace AriaDamInvoice.Models;

public sealed class AppSettings
{
    public string OutputDirectory { get; set; } = string.Empty;

    public List<string> BreedOptions { get; set; } = new()
    {
        "همو",
        "هترو",
        "افشاری"
    };
}