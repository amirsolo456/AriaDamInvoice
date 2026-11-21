using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AriaDamInvoice.Models;

public sealed class AnimalLine : INotifyPropertyChanged
{
    private int _rowNumber;
    private int _count;
    private decimal _weight;
    private string _animalType = string.Empty;
    private string _breed = string.Empty;
    private decimal _unitPrice;
    private string _description = string.Empty;

    public int RowNumber { get => _rowNumber; set => SetField(ref _rowNumber, value); }
    public int Count { get => _count; set { if (SetField(ref _count, value)) OnPropertyChanged(nameof(LineTotal)); } }
    public decimal Weight { get => _weight; set => SetField(ref _weight, value); }
    public string AnimalType { get => _animalType; set => SetField(ref _animalType, value); }
    public string Breed { get => _breed; set => SetField(ref _breed, value); }
    public decimal UnitPrice { get => _unitPrice; set { if (SetField(ref _unitPrice, value)) OnPropertyChanged(nameof(LineTotal)); } }
    public string Description { get => _description; set => SetField(ref _description, value); }
    public decimal LineTotal => Count * UnitPrice;

    public event PropertyChangedEventHandler? PropertyChanged;

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}