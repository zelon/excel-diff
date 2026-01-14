using ExcelDiff.Core.Enums;
using ExcelDiff.Core.Models;

namespace ExcelDiff.Core.Services;

public class DiffEngineService : IDiffEngine
{
    public List<DiffResult> CompareExcelFiles(ExcelFile oldFile, ExcelFile newFile, string keyColumn = "")
    {
        var results = new List<DiffResult>();

        var oldSheetNames = oldFile.Sheets.Select(s => s.Name).ToHashSet();
        var newSheetNames = newFile.Sheets.Select(s => s.Name).ToHashSet();

        var commonSheetNames = oldSheetNames.Intersect(newSheetNames).ToList();

        foreach (var sheetName in commonSheetNames.OrderBy(n => n))
        {
            var oldSheet = oldFile.Sheets.First(s => s.Name == sheetName);
            var newSheet = newFile.Sheets.First(s => s.Name == sheetName);

            var result = CompareSheets(oldSheet, newSheet, keyColumn);
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

    public DiffResult CompareSheets(Sheet oldSheet, Sheet newSheet, string keyColumn = "")
    {
        // 기준 컬럼이 없으면 기존 방식으로 비교
        if (string.IsNullOrWhiteSpace(keyColumn))
        {
            return CompareSheetsPositionBased(oldSheet, newSheet);
        }

        // 기준 컬럼이 있으면 키 기반 비교
        return CompareSheetsKeyBased(oldSheet, newSheet, keyColumn);
    }

    private DiffResult CompareSheetsPositionBased(Sheet oldSheet, Sheet newSheet)
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

    private DiffResult CompareSheetsKeyBased(Sheet oldSheet, Sheet newSheet, string keyColumn)
    {
        var result = new DiffResult(oldSheet.Name);
        var keyColIndex = ParseColumnName(keyColumn);

        // 기준 컬럼의 값을 키로 하여 행 번호 매핑
        var oldKeyToRow = BuildKeyToRowMap(oldSheet, keyColIndex);
        var newKeyToRow = BuildKeyToRowMap(newSheet, keyColIndex);

        // 모든 고유 키 수집
        var allKeys = new HashSet<string>(oldKeyToRow.Keys.Union(newKeyToRow.Keys));

        // 각 키별로 행 비교
        foreach (var key in allKeys.OrderBy(k => k))
        {
            var hasOldRow = oldKeyToRow.TryGetValue(key, out var oldRow);
            var hasNewRow = newKeyToRow.TryGetValue(key, out var newRow);

            if (!hasOldRow && hasNewRow)
            {
                // New에만 있는 키 -> 해당 행의 모든 셀을 Added
                AddRowDiffs(result, newSheet, newRow, DiffType.Added, isOld: false);
            }
            else if (hasOldRow && !hasNewRow)
            {
                // Old에만 있는 키 -> 해당 행의 모든 셀을 Deleted
                AddRowDiffs(result, oldSheet, oldRow, DiffType.Deleted, isOld: true);
            }
            else
            {
                // 양쪽에 있는 키 -> 행끼리 셀별 비교
                CompareRowsByKey(result, oldSheet, newSheet, oldRow, newRow);
            }
        }

        return result;
    }

    private void CompareRowsByKey(DiffResult result, Sheet oldSheet, Sheet newSheet, int oldRow, int newRow)
    {
        // 해당 행의 모든 컬럼 수집
        var oldCells = oldSheet.Cells.Where(c => c.Key.Row == oldRow).ToDictionary(c => c.Key.Column, c => c.Value);
        var newCells = newSheet.Cells.Where(c => c.Key.Row == newRow).ToDictionary(c => c.Key.Column, c => c.Value);

        var allColumns = new HashSet<int>(oldCells.Keys.Union(newCells.Keys));

        foreach (var column in allColumns.OrderBy(c => c))
        {
            var hasOld = oldCells.TryGetValue(column, out var oldCell);
            var hasNew = newCells.TryGetValue(column, out var newCell);

            // 새 위치로 주소 생성 (newRow 사용)
            var address = new CellAddress(newRow, column);
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
            result.Statistics.TotalCells++;
        }
    }

    private void AddRowDiffs(DiffResult result, Sheet sheet, int row, DiffType diffType, bool isOld)
    {
        var cells = sheet.Cells.Where(c => c.Key.Row == row).OrderBy(c => c.Key.Column);

        foreach (var cell in cells)
        {
            var diff = diffType == DiffType.Added
                ? new CellDiff(cell.Key, DiffType.Added, null, cell.Value.Value)
                : new CellDiff(cell.Key, DiffType.Deleted, cell.Value.Value, null);

            result.CellDiffs.Add(diff);
            result.Statistics.TotalCells++;

            if (diffType == DiffType.Added)
                result.Statistics.AddedCells++;
            else
                result.Statistics.DeletedCells++;
        }
    }

    private Dictionary<string, int> BuildKeyToRowMap(Sheet sheet, int keyColIndex)
    {
        var keyToRow = new Dictionary<string, int>();

        // 해당 컬럼의 모든 셀 값을 키로 사용
        foreach (var cell in sheet.Cells.Where(c => c.Key.Column == keyColIndex))
        {
            var key = cell.Value.Value ?? string.Empty;
            // 중복 키가 있으면 마지막 행을 사용 (또는 첫 번째 행을 사용하도록 변경 가능)
            keyToRow[key] = cell.Key.Row;
        }

        return keyToRow;
    }

    private int ParseColumnName(string columnName)
    {
        // A=1, B=2, ..., Z=26, AA=27, AB=28, ...
        columnName = columnName.ToUpperInvariant().Trim();
        int result = 0;

        foreach (char c in columnName)
        {
            if (c < 'A' || c > 'Z')
                throw new ArgumentException($"Invalid column name: {columnName}");

            result = result * 26 + (c - 'A' + 1);
        }

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
