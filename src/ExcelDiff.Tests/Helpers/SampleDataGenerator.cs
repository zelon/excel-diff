using ClosedXML.Excel;

namespace ExcelDiff.Tests.Helpers;

public static class SampleDataGenerator
{
    public static void GenerateSampleFiles(string testDataPath)
    {
        GenerateOldFile(Path.Combine(testDataPath, "sample_old.xlsx"));
        GenerateNewFile(Path.Combine(testDataPath, "sample_new.xlsx"));
    }

    private static void GenerateOldFile(string filePath)
    {
        using var workbook = new XLWorkbook();

        // Sheet1: 직원 명단
        var sheet1 = workbook.Worksheets.Add("직원명단");
        sheet1.Cell("A1").Value = "사번";
        sheet1.Cell("B1").Value = "이름";
        sheet1.Cell("C1").Value = "부서";
        sheet1.Cell("D1").Value = "직급";
        sheet1.Cell("E1").Value = "연봉";

        sheet1.Cell("A2").Value = "E001";
        sheet1.Cell("B2").Value = "김철수";
        sheet1.Cell("C2").Value = "개발팀";
        sheet1.Cell("D2").Value = "대리";
        sheet1.Cell("E2").Value = 45000000;

        sheet1.Cell("A3").Value = "E002";
        sheet1.Cell("B3").Value = "이영희";
        sheet1.Cell("C3").Value = "기획팀";
        sheet1.Cell("D3").Value = "과장";
        sheet1.Cell("E3").Value = 55000000;

        sheet1.Cell("A4").Value = "E003";
        sheet1.Cell("B4").Value = "박민수";
        sheet1.Cell("C4").Value = "개발팀";
        sheet1.Cell("D4").Value = "사원";
        sheet1.Cell("E4").Value = 35000000;

        sheet1.Cell("A5").Value = "E004";
        sheet1.Cell("B5").Value = "정수진";
        sheet1.Cell("C5").Value = "디자인팀";
        sheet1.Cell("D5").Value = "주임";
        sheet1.Cell("E5").Value = 40000000;

        // Sheet2: 매출 데이터
        var sheet2 = workbook.Worksheets.Add("매출");
        sheet2.Cell("A1").Value = "날짜";
        sheet2.Cell("B1").Value = "제품명";
        sheet2.Cell("C1").Value = "수량";
        sheet2.Cell("D1").Value = "금액";

        sheet2.Cell("A2").Value = "2025-01-01";
        sheet2.Cell("B2").Value = "노트북";
        sheet2.Cell("C2").Value = 10;
        sheet2.Cell("D2").Value = 15000000;

        sheet2.Cell("A3").Value = "2025-01-02";
        sheet2.Cell("B3").Value = "마우스";
        sheet2.Cell("C3").Value = 50;
        sheet2.Cell("D3").Value = 1000000;

        sheet2.Cell("A4").Value = "2025-01-03";
        sheet2.Cell("B4").Value = "키보드";
        sheet2.Cell("C4").Value = 30;
        sheet2.Cell("D4").Value = 2400000;

        // Sheet3: 재고 현황
        var sheet3 = workbook.Worksheets.Add("재고");
        sheet3.Cell("A1").Value = "제품코드";
        sheet3.Cell("B1").Value = "제품명";
        sheet3.Cell("C1").Value = "재고";

        sheet3.Cell("A2").Value = "P001";
        sheet3.Cell("B2").Value = "노트북";
        sheet3.Cell("C2").Value = 100;

        sheet3.Cell("A3").Value = "P002";
        sheet3.Cell("B3").Value = "마우스";
        sheet3.Cell("C3").Value = 500;

        workbook.SaveAs(filePath);
    }

    private static void GenerateNewFile(string filePath)
    {
        using var workbook = new XLWorkbook();

        // Sheet1: 직원 명단 (변경사항 포함)
        var sheet1 = workbook.Worksheets.Add("직원명단");
        sheet1.Cell("A1").Value = "사번";
        sheet1.Cell("B1").Value = "이름";
        sheet1.Cell("C1").Value = "부서";
        sheet1.Cell("D1").Value = "직급";
        sheet1.Cell("E1").Value = "연봉";

        sheet1.Cell("A2").Value = "E001";
        sheet1.Cell("B2").Value = "김철수";
        sheet1.Cell("C2").Value = "개발팀";
        sheet1.Cell("D2").Value = "과장";  // 변경: 대리 → 과장
        sheet1.Cell("E2").Value = 50000000;  // 변경: 45000000 → 50000000

        sheet1.Cell("A3").Value = "E002";
        sheet1.Cell("B3").Value = "이영희";
        sheet1.Cell("C3").Value = "기획팀";
        sheet1.Cell("D3").Value = "과장";
        sheet1.Cell("E3").Value = 55000000;  // 변경 없음

        // E003 삭제됨 (박민수)

        sheet1.Cell("A4").Value = "E004";
        sheet1.Cell("B4").Value = "정수진";
        sheet1.Cell("C4").Value = "디자인팀";
        sheet1.Cell("D4").Value = "대리";  // 변경: 주임 → 대리
        sheet1.Cell("E4").Value = 45000000;  // 변경: 40000000 → 45000000

        // E005 추가됨
        sheet1.Cell("A5").Value = "E005";
        sheet1.Cell("B5").Value = "최동욱";
        sheet1.Cell("C5").Value = "영업팀";
        sheet1.Cell("D5").Value = "대리";
        sheet1.Cell("E5").Value = 48000000;

        // Sheet2: 매출 데이터 (변경사항 포함)
        var sheet2 = workbook.Worksheets.Add("매출");
        sheet2.Cell("A1").Value = "날짜";
        sheet2.Cell("B1").Value = "제품명";
        sheet2.Cell("C1").Value = "수량";
        sheet2.Cell("D1").Value = "금액";

        sheet2.Cell("A2").Value = "2025-01-01";
        sheet2.Cell("B2").Value = "노트북";
        sheet2.Cell("C2").Value = 15;  // 변경: 10 → 15
        sheet2.Cell("D2").Value = 22500000;  // 변경: 15000000 → 22500000

        sheet2.Cell("A3").Value = "2025-01-02";
        sheet2.Cell("B3").Value = "마우스";
        sheet2.Cell("C3").Value = 50;  // 변경 없음
        sheet2.Cell("D3").Value = 1000000;  // 변경 없음

        sheet2.Cell("A4").Value = "2025-01-03";
        sheet2.Cell("B4").Value = "키보드";
        sheet2.Cell("C4").Value = 30;  // 변경 없음
        sheet2.Cell("D4").Value = 2400000;  // 변경 없음

        // 새로운 행 추가
        sheet2.Cell("A5").Value = "2025-01-04";
        sheet2.Cell("B5").Value = "모니터";
        sheet2.Cell("C5").Value = 20;
        sheet2.Cell("D5").Value = 6000000;

        // Sheet3: 재고 현황 (변경사항 포함)
        var sheet3 = workbook.Worksheets.Add("재고");
        sheet3.Cell("A1").Value = "제품코드";
        sheet3.Cell("B1").Value = "제품명";
        sheet3.Cell("C1").Value = "재고";

        sheet3.Cell("A2").Value = "P001";
        sheet3.Cell("B2").Value = "노트북";
        sheet3.Cell("C2").Value = 85;  // 변경: 100 → 85

        sheet3.Cell("A3").Value = "P002";
        sheet3.Cell("B3").Value = "마우스";
        sheet3.Cell("C3").Value = 450;  // 변경: 500 → 450

        // 새로운 제품 추가
        sheet3.Cell("A4").Value = "P003";
        sheet3.Cell("B4").Value = "모니터";
        sheet3.Cell("C4").Value = 80;

        // Sheet4: 새로운 시트 추가
        var sheet4 = workbook.Worksheets.Add("신규시트");
        sheet4.Cell("A1").Value = "카테고리";
        sheet4.Cell("B1").Value = "값";
        sheet4.Cell("A2").Value = "테스트";
        sheet4.Cell("B2").Value = "데이터";

        workbook.SaveAs(filePath);
    }

    public static void GenerateSmallSampleFiles(string testDataPath)
    {
        GenerateSmallOldFile(Path.Combine(testDataPath, "small_old.xlsx"));
        GenerateSmallNewFile(Path.Combine(testDataPath, "small_new.xlsx"));
    }

    private static void GenerateSmallOldFile(string filePath)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");

        sheet.Cell("A1").Value = "Name";
        sheet.Cell("B1").Value = "Age";
        sheet.Cell("A2").Value = "Alice";
        sheet.Cell("B2").Value = 30;
        sheet.Cell("A3").Value = "Bob";
        sheet.Cell("B3").Value = 25;

        workbook.SaveAs(filePath);
    }

    private static void GenerateSmallNewFile(string filePath)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");

        sheet.Cell("A1").Value = "Name";
        sheet.Cell("B1").Value = "Age";
        sheet.Cell("A2").Value = "Alice";
        sheet.Cell("B2").Value = 31;  // 변경: 30 → 31
        sheet.Cell("A3").Value = "Charlie";  // 변경: Bob → Charlie
        sheet.Cell("B3").Value = 28;  // 변경: 25 → 28

        workbook.SaveAs(filePath);
    }
}
