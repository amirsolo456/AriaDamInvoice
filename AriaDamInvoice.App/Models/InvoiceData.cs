using AriaDamInvoice.Models;

namespace AriaDamInvoice.Models;

public sealed class InvoiceData
{
    public required string ContractDate { get; init; }
    public required int DocumentNumber { get; init; }
    public string BuyerName { get; init; } = string.Empty;
    public string Phone { get; init; } = string.Empty;
    public string Address { get; init; } = string.Empty;
    public required IReadOnlyList<AnimalLine> Animals { get; init; }
    public required IReadOnlyList<string> EarTags { get; init; }
}