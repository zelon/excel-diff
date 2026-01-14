using ExcelDiff.Core.Models;

namespace ExcelDiff.Core.Services;

public interface IDiffEngine
{
    List<DiffResult> CompareExcelFiles(ExcelFile oldFile, ExcelFile newFile, string keyColumn = "");
    DiffResult CompareSheets(Sheet oldSheet, Sheet newSheet, string keyColumn = "");
}
