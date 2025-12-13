namespace ExcelDiff.Core.Models;

public readonly struct CellAddress : IEquatable<CellAddress>
{
    public int Row { get; init; }
    public int Column { get; init; }

    public CellAddress(int row, int column)
    {
        Row = row;
        Column = column;
    }

    public bool Equals(CellAddress other)
    {
        return Row == other.Row && Column == other.Column;
    }

    public override bool Equals(object? obj)
    {
        return obj is CellAddress other && Equals(other);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Row, Column);
    }

    public static bool operator ==(CellAddress left, CellAddress right)
    {
        return left.Equals(right);
    }

    public static bool operator !=(CellAddress left, CellAddress right)
    {
        return !left.Equals(right);
    }

    public override string ToString()
    {
        return $"R{Row}C{Column}";
    }
}
