using ClosedXML.Excel;
using ExcelDiff.Core.Models;

namespace ExcelDiff.Core.Services;

public class ExcelReaderService : IExcelReader
{
    public ExcelFile ReadExcelFile(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"Excel 파일을 찾을 수 없습니다: {filePath}");
        }

        var excelFile = new ExcelFile(filePath);

        // 파일이 Excel에서 열려있어도 읽을 수 있도록 FileShare.ReadWrite 옵션 사용
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var workbook = new XLWorkbook(stream);

        foreach (var worksheet in workbook.Worksheets)
        {
            var sheet = ReadWorksheet(worksheet);
            excelFile.Sheets.Add(sheet);
        }

        return excelFile;
    }

    public Sheet ReadSheet(string filePath, string sheetName)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"Excel 파일을 찾을 수 없습니다: {filePath}");
        }

        // 파일이 Excel에서 열려있어도 읽을 수 있도록 FileShare.ReadWrite 옵션 사용
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet(sheetName);

        if (worksheet == null)
        {
            throw new ArgumentException($"시트를 찾을 수 없습니다: {sheetName}");
        }

        return ReadWorksheet(worksheet);
    }

    private Sheet ReadWorksheet(IXLWorksheet worksheet)
    {
        var sheet = new Sheet(worksheet.Name);

        var usedRange = worksheet.RangeUsed();
        if (usedRange == null)
        {
            return sheet;
        }

        sheet.RowCount = usedRange.RowCount();
        sheet.ColumnCount = usedRange.ColumnCount();

        foreach (var cell in usedRange.CellsUsed())
        {
            var address = new CellAddress(cell.Address.RowNumber, cell.Address.ColumnNumber);

            var cellModel = new Cell(address, cell.GetValue<string>())
            {
                FormattedValue = cell.GetFormattedString(),
                Formula = cell.HasFormula ? cell.FormulaA1 : null
            };

            sheet.Cells[address] = cellModel;
        }

        return sheet;
    }
}
