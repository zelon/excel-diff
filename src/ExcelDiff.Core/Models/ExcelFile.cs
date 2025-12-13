namespace ExcelDiff.Core.Models;

public class ExcelFile
{
    public string FilePath { get; set; } = string.Empty;
    public List<Sheet> Sheets { get; set; } = new();
    public DateTime LoadedAt { get; set; } = DateTime.Now;

    public ExcelFile()
    {
    }

    public ExcelFile(string filePath)
    {
        FilePath = filePath;
    }
}
