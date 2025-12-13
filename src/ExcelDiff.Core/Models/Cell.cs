namespace ExcelDiff.Core.Models;

public class Cell
{
    public CellAddress Address { get; set; }
    public string Value { get; set; } = string.Empty;
    public string FormattedValue { get; set; } = string.Empty;
    public string? Formula { get; set; }

    public Cell()
    {
    }

    public Cell(CellAddress address, string value)
    {
        Address = address;
        Value = value;
        FormattedValue = value;
    }
}
