using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using ExcelDiff.Core.Enums;
using ExcelDiff.Core.Models;

namespace ExcelDiff.App.ViewModels;

public partial class SheetComparisonViewModel : ObservableObject
{
    [ObservableProperty]
    private string _sheetName = string.Empty;

    [ObservableProperty]
    private ComparisonStatistics _statistics = new();

    [ObservableProperty]
    private double _scrollOffsetX;

    [ObservableProperty]
    private double _scrollOffsetY;

    public ObservableCollection<RowViewModel> OldRows { get; } = new();
    public ObservableCollection<RowViewModel> NewRows { get; } = new();

    public SheetComparisonViewModel()
    {
    }

    public SheetComparisonViewModel(DiffResult diffResult, ExcelFile oldFile, ExcelFile newFile)
    {
        SheetName = diffResult.SheetName;
        Statistics = diffResult.Statistics;

        var oldSheet = oldFile.Sheets.FirstOrDefault(s => s.Name == diffResult.SheetName);
        var newSheet = newFile.Sheets.FirstOrDefault(s => s.Name == diffResult.SheetName);

        BuildRowViewModels(diffResult, oldSheet, newSheet);
    }

    private void BuildRowViewModels(DiffResult diffResult, Sheet? oldSheet, Sheet? newSheet)
    {
        var maxRow = diffResult.CellDiffs.Max(c => c.Address.Row);
        var maxCol = diffResult.CellDiffs.Max(c => c.Address.Column);

        for (int row = 1; row <= maxRow; row++)
        {
            var oldRowVm = new RowViewModel { RowNumber = row };
            var newRowVm = new RowViewModel { RowNumber = row };

            for (int col = 1; col <= maxCol; col++)
            {
                var address = new CellAddress(row, col);
                var diff = diffResult.CellDiffs.FirstOrDefault(d => d.Address == address);

                if (diff != null)
                {
                    var oldCellVm = new CellViewModel
                    {
                        Value = diff.OldValue ?? "",
                        DiffType = diff.DiffType,
                        Address = address
                    };

                    var newCellVm = new CellViewModel
                    {
                        Value = diff.NewValue ?? "",
                        DiffType = diff.DiffType,
                        Address = address
                    };

                    oldRowVm.Cells.Add(oldCellVm);
                    newRowVm.Cells.Add(newCellVm);
                }
                else
                {
                    oldRowVm.Cells.Add(new CellViewModel { Value = "", DiffType = DiffType.Unchanged, Address = address });
                    newRowVm.Cells.Add(new CellViewModel { Value = "", DiffType = DiffType.Unchanged, Address = address });
                }
            }

            OldRows.Add(oldRowVm);
            NewRows.Add(newRowVm);
        }
    }
}

public partial class RowViewModel : ObservableObject
{
    [ObservableProperty]
    private int _rowNumber;

    public ObservableCollection<CellViewModel> Cells { get; } = new();
}

public partial class CellViewModel : ObservableObject
{
    [ObservableProperty]
    private string _value = string.Empty;

    [ObservableProperty]
    private DiffType _diffType;

    [ObservableProperty]
    private CellAddress _address;

    public string ToolTip => DiffType switch
    {
        DiffType.Modified => $"변경됨: {Value}",
        DiffType.Added => "추가됨",
        DiffType.Deleted => "삭제됨",
        _ => Value
    };
}
