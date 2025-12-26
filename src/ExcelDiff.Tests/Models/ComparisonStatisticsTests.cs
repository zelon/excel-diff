using ExcelDiff.Core.Models;
using FluentAssertions;

namespace ExcelDiff.Tests.Models;

public class ComparisonStatisticsTests
{
    [Fact]
    public void ChangePercentage_ShouldCalculateCorrectly()
    {
        // Arrange
        var stats = new ComparisonStatistics
        {
            TotalCells = 100,
            AddedCells = 10,
            DeletedCells = 5,
            ModifiedCells = 15,
            UnchangedCells = 70
        };

        // Act
        var changePercentage = stats.ChangePercentage;

        // Assert
        changePercentage.Should().BeApproximately(30.0, 0.01); // (10+5+15)/100 * 100 = 30%
    }

    [Fact]
    public void ChangePercentage_ShouldReturnZero_WhenTotalCellsIsZero()
    {
        // Arrange
        var stats = new ComparisonStatistics
        {
            TotalCells = 0,
            AddedCells = 0,
            DeletedCells = 0,
            ModifiedCells = 0,
            UnchangedCells = 0
        };

        // Act
        var changePercentage = stats.ChangePercentage;

        // Assert
        changePercentage.Should().Be(0);
    }

    [Fact]
    public void ChangedCells_ShouldSumAllChanges()
    {
        // Arrange
        var stats = new ComparisonStatistics
        {
            AddedCells = 5,
            DeletedCells = 3,
            ModifiedCells = 7
        };

        // Act
        var changedCells = stats.ChangedCells;

        // Assert
        changedCells.Should().Be(15); // 5 + 3 + 7
    }

    [Fact]
    public void ChangePercentage_ShouldBe100_WhenAllCellsChanged()
    {
        // Arrange
        var stats = new ComparisonStatistics
        {
            TotalCells = 50,
            AddedCells = 20,
            DeletedCells = 15,
            ModifiedCells = 15,
            UnchangedCells = 0
        };

        // Act
        var changePercentage = stats.ChangePercentage;

        // Assert
        changePercentage.Should().BeApproximately(100.0, 0.01);
    }

    [Fact]
    public void ChangePercentage_ShouldBe0_WhenNoCellsChanged()
    {
        // Arrange
        var stats = new ComparisonStatistics
        {
            TotalCells = 100,
            AddedCells = 0,
            DeletedCells = 0,
            ModifiedCells = 0,
            UnchangedCells = 100
        };

        // Act
        var changePercentage = stats.ChangePercentage;

        // Assert
        changePercentage.Should().Be(0);
    }
}
