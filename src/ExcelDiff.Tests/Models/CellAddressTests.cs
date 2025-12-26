using ExcelDiff.Core.Models;
using FluentAssertions;

namespace ExcelDiff.Tests.Models;

public class CellAddressTests
{
    [Fact]
    public void CellAddress_ShouldInitializeCorrectly()
    {
        // Arrange & Act
        var address = new CellAddress(5, 10);

        // Assert
        address.Row.Should().Be(5);
        address.Column.Should().Be(10);
    }

    [Fact]
    public void CellAddress_Equals_ShouldReturnTrue_ForSameAddress()
    {
        // Arrange
        var address1 = new CellAddress(3, 7);
        var address2 = new CellAddress(3, 7);

        // Act & Assert
        address1.Equals(address2).Should().BeTrue();
        (address1 == address2).Should().BeTrue();
    }

    [Fact]
    public void CellAddress_Equals_ShouldReturnFalse_ForDifferentAddress()
    {
        // Arrange
        var address1 = new CellAddress(3, 7);
        var address2 = new CellAddress(3, 8);

        // Act & Assert
        address1.Equals(address2).Should().BeFalse();
        (address1 != address2).Should().BeTrue();
    }

    [Fact]
    public void CellAddress_GetHashCode_ShouldBeSame_ForEqualAddresses()
    {
        // Arrange
        var address1 = new CellAddress(2, 5);
        var address2 = new CellAddress(2, 5);

        // Act & Assert
        address1.GetHashCode().Should().Be(address2.GetHashCode());
    }

    [Fact]
    public void CellAddress_ToString_ShouldReturnCorrectFormat()
    {
        // Arrange
        var address = new CellAddress(10, 20);

        // Act
        var result = address.ToString();

        // Assert
        result.Should().Be("R10C20");
    }

    [Fact]
    public void CellAddress_CanBeUsedAsDictionaryKey()
    {
        // Arrange
        var dict = new Dictionary<CellAddress, string>();
        var address1 = new CellAddress(1, 1);
        var address2 = new CellAddress(1, 1);
        var address3 = new CellAddress(2, 2);

        // Act
        dict[address1] = "Value1";
        dict[address2] = "Value2"; // 같은 키이므로 덮어씀
        dict[address3] = "Value3";

        // Assert
        dict.Should().HaveCount(2);
        dict[address1].Should().Be("Value2");
        dict[address3].Should().Be("Value3");
    }
}
