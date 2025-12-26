#!/usr/bin/env dotnet-script
#r "nuget: ClosedXML, 0.105.0"

using ClosedXML.Excel;

var testDataPath = Path.Combine(Directory.GetCurrentDirectory(), "src", "ExcelDiff.Tests", "TestData");
Directory.CreateDirectory(testDataPath);

Console.WriteLine($"샘플 Excel 파일 생성 중: {testDataPath}\n");

// sample_old.xlsx 생성
var oldFilePath = Path.Combine(testDataPath, "sample_old.xlsx");
using (var workbook = new XLWorkbook())
{
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

    var sheet2 = workbook.Worksheets.Add("매출");
    sheet2.Cell("A1").Value = "날짜";
    sheet2.Cell("B1").Value = "제품명";
    sheet2.Cell("C1").Value = "수량";
    sheet2.Cell("D1").Value = "금액";

    sheet2.Cell("A2").Value = "2025-01-01";
    sheet2.Cell("B2").Value = "노트북";
    sheet2.Cell("C2").Value = 10;
    sheet2.Cell("D2").Value = 15000000;

    workbook.SaveAs(oldFilePath);
}
Console.WriteLine($"✓ {oldFilePath} 생성 완료");

// sample_new.xlsx 생성
var newFilePath = Path.Combine(testDataPath, "sample_new.xlsx");
using (var workbook = new XLWorkbook())
{
    var sheet1 = workbook.Worksheets.Add("직원명단");
    sheet1.Cell("A1").Value = "사번";
    sheet1.Cell("B1").Value = "이름";
    sheet1.Cell("C1").Value = "부서";
    sheet1.Cell("D1").Value = "직급";
    sheet1.Cell("E1").Value = "연봉";

    sheet1.Cell("A2").Value = "E001";
    sheet1.Cell("B2").Value = "김철수";
    sheet1.Cell("C2").Value = "개발팀";
    sheet1.Cell("D2").Value = "과장"; // 변경!
    sheet1.Cell("E2").Value = 50000000; // 변경!

    sheet1.Cell("A3").Value = "E002";
    sheet1.Cell("B3").Value = "이영희";
    sheet1.Cell("C3").Value = "기획팀";
    sheet1.Cell("D3").Value = "과장";
    sheet1.Cell("E3").Value = 55000000;

    // E003 삭제됨

    // E005 추가됨
    sheet1.Cell("A4").Value = "E005";
    sheet1.Cell("B4").Value = "최동욱";
    sheet1.Cell("C4").Value = "영업팀";
    sheet1.Cell("D4").Value = "대리";
    sheet1.Cell("E4").Value = 48000000;

    var sheet2 = workbook.Worksheets.Add("매출");
    sheet2.Cell("A1").Value = "날짜";
    sheet2.Cell("B1").Value = "제품명";
    sheet2.Cell("C1").Value = "수량";
    sheet2.Cell("D1").Value = "금액";

    sheet2.Cell("A2").Value = "2025-01-01";
    sheet2.Cell("B2").Value = "노트북";
    sheet2.Cell("C2").Value = 15; // 변경!
    sheet2.Cell("D2").Value = 22500000; // 변경!

    // 새 시트 추가
    var sheet3 = workbook.Worksheets.Add("신규시트");
    sheet3.Cell("A1").Value = "카테고리";
    sheet3.Cell("B1").Value = "값";

    workbook.SaveAs(newFilePath);
}
Console.WriteLine($"✓ {newFilePath} 생성 완료");

Console.WriteLine("\n생성 완료! 📊");
Console.WriteLine("애플리케이션을 실행하고 위 파일들로 비교 테스트를 해보세요.");
