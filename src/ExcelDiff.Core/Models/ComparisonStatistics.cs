namespace ExcelDiff.Core.Models;

public class ComparisonStatistics
{
    public int TotalCells { get; set; }
    public int AddedCells { get; set; }
    public int DeletedCells { get; set; }
    public int ModifiedCells { get; set; }
    public int UnchangedCells { get; set; }

    public double ChangePercentage =>
        TotalCells > 0
            ? ((double)(AddedCells + DeletedCells + ModifiedCells) / TotalCells) * 100
            : 0;

    public int ChangedCells => AddedCells + DeletedCells + ModifiedCells;
}
