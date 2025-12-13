using ExcelDiff.Core.Enums;
using ExcelDiff.Core.Models;

namespace ExcelDiff.Core.Services;

public class DiffEngineService : IDiffEngine
{
    public List<DiffResult> CompareExcelFiles(ExcelFile oldFile, ExcelFile newFile)
    {
        var results = new List<DiffResult>();

        var oldSheetNames = oldFile.Sheets.Select(s => s.Name).ToHashSet();
        var newSheetNames = newFile.Sheets.Select(s => s.Name).ToHashSet();

        var commonSheetNames = oldSheetNames.Intersect(newSheetNames).ToList();

        foreach (var sheetName in commonSheetNames.OrderBy(n => n))
        {
            var oldSheet = oldFile.Sheets.First(s => s.Name == sheetName);
            var newSheet = newFile.Sheets.First(s => s.Name == sheetName);

            var result = CompareSheets(oldSheet, newSheet);
            results.Add(result);
        }

        foreach (var sheetName in oldSheetNames.Except(newSheetNames).OrderBy(n => n))
        {
            var oldSheet = oldFile.Sheets.First(s => s.Name == sheetName);
            var result = CreateDeletedSheetResult(oldSheet);
            results.Add(result);
        }

        foreach (var sheetName in newSheetNames.Except(oldSheetNames).OrderBy(n => n))
        {
            var newSheet = newFile.Sheets.First(s => s.Name == sheetName);
            var result = CreateAddedSheetResult(newSheet);
            results.Add(result);
        }

        return results;
    }

    public DiffResult CompareSheets(Sheet oldSheet, Sheet newSheet)
    {
        var result = new DiffResult(oldSheet.Name);

        var allAddresses = new HashSet<CellAddress>(
            oldSheet.Cells.Keys.Union(newSheet.Cells.Keys)
        );

        foreach (var address in allAddresses.OrderBy(a => a.Row).ThenBy(a => a.Column))
        {
            var hasOld = oldSheet.Cells.TryGetValue(address, out var oldCell);
            var hasNew = newSheet.Cells.TryGetValue(address, out var newCell);

            CellDiff diff;

            if (!hasOld && hasNew)
            {
                diff = new CellDiff(address, DiffType.Added, null, newCell!.Value);
                result.Statistics.AddedCells++;
            }
            else if (hasOld && !hasNew)
            {
                diff = new CellDiff(address, DiffType.Deleted, oldCell!.Value, null);
                result.Statistics.DeletedCells++;
            }
            else if (oldCell!.Value != newCell!.Value)
            {
                diff = new CellDiff(address, DiffType.Modified, oldCell.Value, newCell.Value);
                result.Statistics.ModifiedCells++;
            }
            else
            {
                diff = new CellDiff(address, DiffType.Unchanged, oldCell.Value, newCell.Value);
                result.Statistics.UnchangedCells++;
            }

            result.CellDiffs.Add(diff);
        }

        result.Statistics.TotalCells = allAddresses.Count;
        return result;
    }

    private DiffResult CreateDeletedSheetResult(Sheet oldSheet)
    {
        var result = new DiffResult(oldSheet.Name);

        foreach (var cell in oldSheet.Cells.Values.OrderBy(c => c.Address.Row).ThenBy(c => c.Address.Column))
        {
            var diff = new CellDiff(cell.Address, DiffType.Deleted, cell.Value, null);
            result.CellDiffs.Add(diff);
            result.Statistics.DeletedCells++;
        }

        result.Statistics.TotalCells = oldSheet.Cells.Count;
        return result;
    }

    private DiffResult CreateAddedSheetResult(Sheet newSheet)
    {
        var result = new DiffResult(newSheet.Name);

        foreach (var cell in newSheet.Cells.Values.OrderBy(c => c.Address.Row).ThenBy(c => c.Address.Column))
        {
            var diff = new CellDiff(cell.Address, DiffType.Added, null, cell.Value);
            result.CellDiffs.Add(diff);
            result.Statistics.AddedCells++;
        }

        result.Statistics.TotalCells = newSheet.Cells.Count;
        return result;
    }
}
