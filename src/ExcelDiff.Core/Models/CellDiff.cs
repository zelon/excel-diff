using ExcelDiff.Core.Enums;

namespace ExcelDiff.Core.Models;

public class CellDiff
{
    public CellAddress Address { get; set; }
    public DiffType DiffType { get; set; }
    public string? OldValue { get; set; }
    public string? NewValue { get; set; }

    public CellDiff()
    {
    }

    public CellDiff(CellAddress address, DiffType diffType, string? oldValue, string? newValue)
    {
        Address = address;
        DiffType = diffType;
        OldValue = oldValue;
        NewValue = newValue;
    }
}
