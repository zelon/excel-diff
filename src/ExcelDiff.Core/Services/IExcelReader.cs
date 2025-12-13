using ExcelDiff.Core.Models;

namespace ExcelDiff.Core.Services;

public interface IExcelReader
{
    ExcelFile ReadExcelFile(string filePath);
    Sheet ReadSheet(string filePath, string sheetName);
}
