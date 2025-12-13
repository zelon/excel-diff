using ExcelDiff.Core.Models;

namespace ExcelDiff.Core.Services;

public interface IDiffEngine
{
    List<DiffResult> CompareExcelFiles(ExcelFile oldFile, ExcelFile newFile);
    DiffResult CompareSheets(Sheet oldSheet, Sheet newSheet);
}
