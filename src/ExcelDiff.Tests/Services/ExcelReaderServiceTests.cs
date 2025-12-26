using ExcelDiff.Core.Services;
using ExcelDiff.Tests.Helpers;
using FluentAssertions;

namespace ExcelDiff.Tests.Services;

public class ExcelReaderServiceTests : IDisposable
{
    private readonly string _testDataPath;
    private readonly ExcelReaderService _excelReader;

    public ExcelReaderServiceTests()
    {
        _testDataPath = Path.Combine(Path.GetTempPath(), "ExcelDiffTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(_testDataPath);

        // 테스트용 샘플 파일 생성
        SampleDataGenerator.GenerateSmallSampleFiles(_testDataPath);

        _excelReader = new ExcelReaderService();
    }

    [Fact]
    public void ReadExcelFile_ShouldThrowException_WhenFileDoesNotExist()
    {
        // Arrange
        var nonExistentFile = Path.Combine(_testDataPath, "non_existent.xlsx");

        // Act
        Action act = () => _excelReader.ReadExcelFile(nonExistentFile);

        // Assert
        act.Should().Throw<FileNotFoundException>()
            .WithMessage($"*{nonExistentFile}*");
    }

    [Fact]
    public void ReadExcelFile_ShouldReadAllSheets()
    {
        // Arrange
        var filePath = Path.Combine(_testDataPath, "small_old.xlsx");

        // Act
        var excelFile = _excelReader.ReadExcelFile(filePath);

        // Assert
        excelFile.Should().NotBeNull();
        excelFile.FilePath.Should().Be(filePath);
        excelFile.Sheets.Should().HaveCount(1);
        excelFile.Sheets[0].Name.Should().Be("Sheet1");
    }

    [Fact]
    public void ReadExcelFile_ShouldReadCellValues()
    {
        // Arrange
        var filePath = Path.Combine(_testDataPath, "small_old.xlsx");

        // Act
        var excelFile = _excelReader.ReadExcelFile(filePath);
        var sheet = excelFile.Sheets[0];

        // Assert
        sheet.Cells.Should().HaveCount(6); // 헤더 2개 + 데이터 4개

        // 헤더 확인
        var headerName = sheet.Cells.Values.FirstOrDefault(c => c.Address.Row == 1 && c.Address.Column == 1);
        headerName.Should().NotBeNull();
        headerName!.Value.Should().Be("Name");

        var headerAge = sheet.Cells.Values.FirstOrDefault(c => c.Address.Row == 1 && c.Address.Column == 2);
        headerAge.Should().NotBeNull();
        headerAge!.Value.Should().Be("Age");

        // 데이터 확인
        var aliceName = sheet.Cells.Values.FirstOrDefault(c => c.Address.Row == 2 && c.Address.Column == 1);
        aliceName.Should().NotBeNull();
        aliceName!.Value.Should().Be("Alice");

        var aliceAge = sheet.Cells.Values.FirstOrDefault(c => c.Address.Row == 2 && c.Address.Column == 2);
        aliceAge.Should().NotBeNull();
        aliceAge!.Value.Should().Be("30");
    }

    [Fact]
    public void ReadSheet_ShouldReadSpecificSheet()
    {
        // Arrange
        var filePath = Path.Combine(_testDataPath, "small_old.xlsx");
        var sheetName = "Sheet1";

        // Act
        var sheet = _excelReader.ReadSheet(filePath, sheetName);

        // Assert
        sheet.Should().NotBeNull();
        sheet.Name.Should().Be(sheetName);
        sheet.Cells.Should().NotBeEmpty();
    }

    [Fact]
    public void ReadSheet_ShouldThrowException_WhenSheetDoesNotExist()
    {
        // Arrange
        var filePath = Path.Combine(_testDataPath, "small_old.xlsx");
        var nonExistentSheet = "NonExistentSheet";

        // Act
        Action act = () => _excelReader.ReadSheet(filePath, nonExistentSheet);

        // Assert
        act.Should().Throw<ArgumentException>()
            .WithMessage($"*{nonExistentSheet}*");
    }

    [Fact]
    public void ReadExcelFile_ShouldHandleEmptySheet()
    {
        // Arrange
        var emptyFilePath = Path.Combine(_testDataPath, "empty.xlsx");

        // 빈 Excel 파일 생성
        using (var workbook = new ClosedXML.Excel.XLWorkbook())
        {
            workbook.Worksheets.Add("EmptySheet");
            workbook.SaveAs(emptyFilePath);
        }

        // Act
        var excelFile = _excelReader.ReadExcelFile(emptyFilePath);

        // Assert
        excelFile.Should().NotBeNull();
        excelFile.Sheets.Should().HaveCount(1);
        excelFile.Sheets[0].Cells.Should().BeEmpty();
    }

    public void Dispose()
    {
        // 테스트 후 임시 파일 정리
        if (Directory.Exists(_testDataPath))
        {
            Directory.Delete(_testDataPath, recursive: true);
        }
    }
}
