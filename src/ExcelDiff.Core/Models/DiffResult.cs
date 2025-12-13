namespace ExcelDiff.Core.Models;

public class DiffResult
{
    public string SheetName { get; set; } = string.Empty;
    public List<CellDiff> CellDiffs { get; set; } = new();
    public ComparisonStatistics Statistics { get; set; } = new();

    public DiffResult()
    {
    }

    public DiffResult(string sheetName)
    {
        SheetName = sheetName;
    }
}
