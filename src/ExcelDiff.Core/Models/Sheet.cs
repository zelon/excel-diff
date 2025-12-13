namespace ExcelDiff.Core.Models;

public class Sheet
{
    public string Name { get; set; } = string.Empty;
    public int RowCount { get; set; }
    public int ColumnCount { get; set; }
    public Dictionary<CellAddress, Cell> Cells { get; set; } = new();

    public Sheet()
    {
    }

    public Sheet(string name)
    {
        Name = name;
    }
}
