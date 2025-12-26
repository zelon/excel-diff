using ExcelDiff.Core.Enums;
using ExcelDiff.Core.Models;
using ExcelDiff.Core.Services;
using FluentAssertions;

namespace ExcelDiff.Tests.Services;

public class DiffEngineServiceTests
{
    private readonly DiffEngineService _diffEngine;

    public DiffEngineServiceTests()
    {
        _diffEngine = new DiffEngineService();
    }

    [Fact]
    public void CompareSheets_ShouldDetectAddedCells()
    {
        // Arrange
        var oldSheet = new Sheet("Sheet1");
        oldSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "A1");

        var newSheet = new Sheet("Sheet1");
        newSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "A1");
        newSheet.Cells[new CellAddress(1, 2)] = new Cell(new CellAddress(1, 2), "B1"); // 추가됨

        // Act
        var result = _diffEngine.CompareSheets(oldSheet, newSheet);

        // Assert
        result.Should().NotBeNull();
        result.Statistics.AddedCells.Should().Be(1);
        result.Statistics.UnchangedCells.Should().Be(1);
        result.CellDiffs.Should().Contain(d => d.DiffType == DiffType.Added && d.NewValue == "B1");
    }

    [Fact]
    public void CompareSheets_ShouldDetectDeletedCells()
    {
        // Arrange
        var oldSheet = new Sheet("Sheet1");
        oldSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "A1");
        oldSheet.Cells[new CellAddress(1, 2)] = new Cell(new CellAddress(1, 2), "B1"); // 삭제될 셀

        var newSheet = new Sheet("Sheet1");
        newSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "A1");

        // Act
        var result = _diffEngine.CompareSheets(oldSheet, newSheet);

        // Assert
        result.Should().NotBeNull();
        result.Statistics.DeletedCells.Should().Be(1);
        result.Statistics.UnchangedCells.Should().Be(1);
        result.CellDiffs.Should().Contain(d => d.DiffType == DiffType.Deleted && d.OldValue == "B1");
    }

    [Fact]
    public void CompareSheets_ShouldDetectModifiedCells()
    {
        // Arrange
        var oldSheet = new Sheet("Sheet1");
        oldSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "OldValue");

        var newSheet = new Sheet("Sheet1");
        newSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "NewValue");

        // Act
        var result = _diffEngine.CompareSheets(oldSheet, newSheet);

        // Assert
        result.Should().NotBeNull();
        result.Statistics.ModifiedCells.Should().Be(1);
        result.Statistics.UnchangedCells.Should().Be(0);

        var modifiedCell = result.CellDiffs.Should().ContainSingle(d => d.DiffType == DiffType.Modified).Subject;
        modifiedCell.OldValue.Should().Be("OldValue");
        modifiedCell.NewValue.Should().Be("NewValue");
    }

    [Fact]
    public void CompareSheets_ShouldDetectUnchangedCells()
    {
        // Arrange
        var oldSheet = new Sheet("Sheet1");
        oldSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "SameValue");

        var newSheet = new Sheet("Sheet1");
        newSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "SameValue");

        // Act
        var result = _diffEngine.CompareSheets(oldSheet, newSheet);

        // Assert
        result.Should().NotBeNull();
        result.Statistics.UnchangedCells.Should().Be(1);
        result.Statistics.ChangedCells.Should().Be(0);
        result.CellDiffs.Should().Contain(d => d.DiffType == DiffType.Unchanged);
    }

    [Fact]
    public void CompareSheets_ShouldCalculateStatisticsCorrectly()
    {
        // Arrange
        var oldSheet = new Sheet("Sheet1");
        oldSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "A1");
        oldSheet.Cells[new CellAddress(1, 2)] = new Cell(new CellAddress(1, 2), "B1");
        oldSheet.Cells[new CellAddress(2, 1)] = new Cell(new CellAddress(2, 1), "A2");

        var newSheet = new Sheet("Sheet1");
        newSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "A1_Modified");
        newSheet.Cells[new CellAddress(1, 3)] = new Cell(new CellAddress(1, 3), "C1_Added");
        newSheet.Cells[new CellAddress(2, 1)] = new Cell(new CellAddress(2, 1), "A2");

        // Act
        var result = _diffEngine.CompareSheets(oldSheet, newSheet);

        // Assert
        result.Statistics.TotalCells.Should().Be(4); // A1, B1, A2, C1
        result.Statistics.ModifiedCells.Should().Be(1); // A1
        result.Statistics.DeletedCells.Should().Be(1); // B1
        result.Statistics.AddedCells.Should().Be(1); // C1
        result.Statistics.UnchangedCells.Should().Be(1); // A2
        result.Statistics.ChangedCells.Should().Be(3); // Modified + Deleted + Added
        result.Statistics.ChangePercentage.Should().BeApproximately(75.0, 0.01);
    }

    [Fact]
    public void CompareExcelFiles_ShouldMatchSheetsByName()
    {
        // Arrange
        var oldFile = new ExcelFile();
        oldFile.Sheets.Add(new Sheet("Sheet1"));
        oldFile.Sheets.Add(new Sheet("Sheet2"));

        var newFile = new ExcelFile();
        newFile.Sheets.Add(new Sheet("Sheet1"));
        newFile.Sheets.Add(new Sheet("Sheet2"));
        newFile.Sheets.Add(new Sheet("Sheet3")); // 새 시트

        // Act
        var results = _diffEngine.CompareExcelFiles(oldFile, newFile);

        // Assert
        results.Should().HaveCount(3); // Sheet1, Sheet2, Sheet3
        results.Should().Contain(r => r.SheetName == "Sheet1");
        results.Should().Contain(r => r.SheetName == "Sheet2");
        results.Should().Contain(r => r.SheetName == "Sheet3");
    }

    [Fact]
    public void CompareExcelFiles_ShouldHandleDeletedSheets()
    {
        // Arrange
        var oldFile = new ExcelFile();
        var oldSheet = new Sheet("DeletedSheet");
        oldSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "Data");
        oldFile.Sheets.Add(oldSheet);

        var newFile = new ExcelFile();
        // DeletedSheet가 없음

        // Act
        var results = _diffEngine.CompareExcelFiles(oldFile, newFile);

        // Assert
        results.Should().HaveCount(1);
        var result = results[0];
        result.SheetName.Should().Be("DeletedSheet");
        result.Statistics.DeletedCells.Should().Be(1);
        result.Statistics.TotalCells.Should().Be(1);
    }

    [Fact]
    public void CompareExcelFiles_ShouldHandleAddedSheets()
    {
        // Arrange
        var oldFile = new ExcelFile();
        // 빈 파일

        var newFile = new ExcelFile();
        var newSheet = new Sheet("AddedSheet");
        newSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "NewData");
        newFile.Sheets.Add(newSheet);

        // Act
        var results = _diffEngine.CompareExcelFiles(oldFile, newFile);

        // Assert
        results.Should().HaveCount(1);
        var result = results[0];
        result.SheetName.Should().Be("AddedSheet");
        result.Statistics.AddedCells.Should().Be(1);
        result.Statistics.TotalCells.Should().Be(1);
    }

    [Fact]
    public void CompareSheets_ShouldHandleEmptySheets()
    {
        // Arrange
        var oldSheet = new Sheet("Empty1");
        var newSheet = new Sheet("Empty2");

        // Act
        var result = _diffEngine.CompareSheets(oldSheet, newSheet);

        // Assert
        result.Should().NotBeNull();
        result.Statistics.TotalCells.Should().Be(0);
        result.Statistics.ChangedCells.Should().Be(0);
        result.CellDiffs.Should().BeEmpty();
    }

    [Fact]
    public void CompareSheets_CellDiffs_ShouldBeSortedByRowThenColumn()
    {
        // Arrange
        var oldSheet = new Sheet("Sheet1");
        oldSheet.Cells[new CellAddress(2, 1)] = new Cell(new CellAddress(2, 1), "A2");
        oldSheet.Cells[new CellAddress(1, 2)] = new Cell(new CellAddress(1, 2), "B1");
        oldSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "A1");

        var newSheet = new Sheet("Sheet1");
        newSheet.Cells[new CellAddress(2, 1)] = new Cell(new CellAddress(2, 1), "A2");
        newSheet.Cells[new CellAddress(1, 2)] = new Cell(new CellAddress(1, 2), "B1");
        newSheet.Cells[new CellAddress(1, 1)] = new Cell(new CellAddress(1, 1), "A1");

        // Act
        var result = _diffEngine.CompareSheets(oldSheet, newSheet);

        // Assert
        result.CellDiffs.Should().HaveCount(3);
        result.CellDiffs[0].Address.Should().Be(new CellAddress(1, 1)); // A1
        result.CellDiffs[1].Address.Should().Be(new CellAddress(1, 2)); // B1
        result.CellDiffs[2].Address.Should().Be(new CellAddress(2, 1)); // A2
    }
}
