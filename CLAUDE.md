# ExcelDiff - Excel 파일 비교 도구

## 프로젝트 개요

ExcelDiff는 Microsoft Excel 파일(*.xlsx)을 시각적으로 비교하는 WPF 데스크톱 애플리케이션입니다. 두 개의 Excel 파일을 좌우로 나란히 표시하며, 셀별 차이를 색상으로 강조하여 변경 사항을 쉽게 파악할 수 있습니다.

### 주요 특징

- **사이드바이사이드 비교**: Old/New 파일을 좌우로 나란히 표시
- **색상 강조**: 추가(초록), 삭제(빨강), 수정(노랑)을 색상으로 구분
- **자동 시트 매칭**: 같은 이름의 시트를 자동으로 매칭하여 비교
- **스크롤 동기화**: 좌우 화면이 동시에 스크롤되어 비교 용이
- **상세 통계**: 전체/추가/삭제/수정 셀 개수 및 변경률 제공
- **성능 최적화**: 가상화를 통한 대용량 파일 처리 지원

## 기술 스택

### 프레임워크 및 언어
- **.NET 8.0**: 최신 .NET 플랫폼
- **C#**: 프로그래밍 언어
- **WPF (Windows Presentation Foundation)**: UI 프레임워크

### 주요 라이브러리
- **ClosedXML (0.105.0)**: Excel 파일 읽기/쓰기 (MIT 라이선스)
- **CommunityToolkit.Mvvm (8.4.0)**: MVVM 패턴 구현
- **Microsoft.Xaml.Behaviors.Wpf (1.1.135)**: WPF Behaviors
- **CsvHelper (33.1.0)**: CSV 내보내기 기능

### 테스트 도구
- **xUnit**: 단위 테스트 프레임워크
- **FluentAssertions**: 테스트 검증
- **Moq**: Mocking 프레임워크

## 프로젝트 구조

```
ExcelDiff/
├── SPEC.md                           # 구현 명세서
├── CLAUDE.md                         # 프로젝트 문서 (이 파일)
├── ExcelDiff.sln                     # Visual Studio 솔루션
└── src/
    ├── ExcelDiff.Core/               # 비즈니스 로직 라이브러리
    │   ├── Models/                   # 데이터 모델
    │   │   ├── Cell.cs               # 셀 정보
    │   │   ├── CellAddress.cs        # 셀 주소 (행/열)
    │   │   ├── Sheet.cs              # 시트 정보
    │   │   ├── ExcelFile.cs          # Excel 파일 정보
    │   │   ├── CellDiff.cs           # 셀 차이 정보
    │   │   ├── DiffResult.cs         # 비교 결과
    │   │   └── ComparisonStatistics.cs # 통계 정보
    │   ├── Services/                 # 비즈니스 로직
    │   │   ├── IExcelReader.cs       # Excel 읽기 인터페이스
    │   │   ├── ExcelReaderService.cs # ClosedXML 기반 구현
    │   │   ├── IDiffEngine.cs        # 비교 엔진 인터페이스
    │   │   └── DiffEngineService.cs  # 셀별 비교 알고리즘
    │   └── Enums/
    │       └── DiffType.cs           # 차이 타입 (Added/Deleted/Modified/Unchanged)
    │
    ├── ExcelDiff.App/                # WPF 애플리케이션
    │   ├── ViewModels/               # MVVM 뷰모델
    │   │   ├── MainViewModel.cs      # 메인 화면 로직
    │   │   └── SheetComparisonViewModel.cs # 시트 비교 로직
    │   ├── Converters/               # 데이터 변환기
    │   │   ├── DiffTypeToColorConverter.cs # 색상 변환
    │   │   └── BoolToVisibilityConverter.cs # 가시성 변환
    │   ├── Helpers/                  # 헬퍼 클래스
    │   │   ├── DataGridHelper.cs     # 동적 컬럼 생성
    │   │   └── ScrollSyncHelper.cs   # 스크롤 동기화
    │   ├── App.xaml                  # 애플리케이션 진입점
    │   └── MainWindow.xaml           # 메인 UI
    │
    └── ExcelDiff.Tests/              # 단위 테스트
        └── (테스트 파일들)
```

## 핵심 컴포넌트 설명

### 1. Core 라이브러리

#### ExcelReaderService
```csharp
// ClosedXML을 사용하여 Excel 파일 읽기
public ExcelFile ReadExcelFile(string filePath)
{
    using var workbook = new XLWorkbook(filePath);
    // 모든 시트를 읽어서 ExcelFile 객체로 반환
}
```

**특징:**
- 모든 시트 자동 로드
- 셀 값, 수식, 포맷 정보 추출
- 빈 셀 처리

#### DiffEngineService
```csharp
// 핵심 비교 알고리즘
public DiffResult CompareSheets(Sheet oldSheet, Sheet newSheet)
{
    // 1. 모든 셀 주소 수집 (Union)
    // 2. 각 셀별로 비교:
    //    - Old에만 존재 → Deleted
    //    - New에만 존재 → Added
    //    - 값이 다름 → Modified
    //    - 값이 같음 → Unchanged
}
```

**특징:**
- 시트 이름 기준 자동 매칭
- 셀 단위 정밀 비교
- 통계 자동 계산

### 2. WPF 애플리케이션

#### MainViewModel (MVVM 패턴)
```csharp
// CommunityToolkit.Mvvm 사용
public partial class MainViewModel : ObservableObject
{
    [ObservableProperty] private string _oldFilePath;
    [ObservableProperty] private string _newFilePath;

    [RelayCommand]
    private async Task CompareAsync()
    {
        // 비동기로 파일 읽기 및 비교
    }
}
```

**특징:**
- Source Generator 기반 속성 자동 생성
- RelayCommand로 간편한 커맨드 바인딩
- 비동기 처리로 UI 반응성 유지

#### DataGridHelper (동적 컬럼 생성)
```csharp
// Attached Property를 사용한 동적 컬럼 생성
helpers:DataGridHelper.AutoGenerateCellColumns="True"
```

**기능:**
- 런타임에 셀 개수만큼 컬럼 자동 생성
- 컬럼 헤더를 A, B, C... 형식으로 표시
- 각 셀에 색상 강조 적용

#### ScrollSyncHelper (스크롤 동기화)
```csharp
// Attached Property를 사용한 스크롤 동기화
helpers:ScrollSyncHelper.ScrollGroup="SheetComparison"
```

**기능:**
- 같은 ScrollGroup의 DataGrid 스크롤 동기화
- 가로/세로 스크롤 모두 지원
- 무한 루프 방지

### 3. UI/UX 기능

#### 색상 강조 시스템
```csharp
DiffType.Added    → LightGreen (#90EE90)
DiffType.Deleted  → LightCoral (#F08080)
DiffType.Modified → Yellow (#FFFF00)
DiffType.Unchanged → White
```

#### 통계 패널
- **전체 셀 개수**: 비교 대상 전체 셀
- **추가**: Old에 없고 New에만 있는 셀
- **삭제**: Old에만 있고 New에 없는 셀
- **수정**: 값이 변경된 셀
- **변경률**: (추가+삭제+수정) / 전체 × 100%

## 빌드 및 실행 방법

### 요구 사항
- **.NET 8.0 SDK** 이상
- **Windows 10/11** (WPF는 Windows 전용)
- **Visual Studio 2022** 또는 **VS Code** (선택사항)

### 빌드
```bash
# 솔루션 빌드
dotnet build

# 또는 Release 모드로 빌드
dotnet build -c Release
```

### 실행
```bash
# 개발 모드 실행
dotnet run --project src/ExcelDiff.App

# 또는 빌드된 실행 파일 직접 실행
cd src/ExcelDiff.App/bin/Debug/net8.0-windows
ExcelDiff.App.exe
```

### 테스트
```bash
# 단위 테스트 실행
dotnet test
```

## 사용 방법

1. **애플리케이션 실행**
   - ExcelDiff.App.exe 실행

2. **파일 선택**
   - "Old 파일" 찾아보기 버튼 클릭 → 이전 버전 Excel 파일 선택
   - "New 파일" 찾아보기 버튼 클릭 → 새 버전 Excel 파일 선택

3. **비교 실행**
   - "비교" 버튼 클릭
   - 로딩 인디케이터가 표시되며 파일 처리

4. **결과 확인**
   - 시트별 탭에서 원하는 시트 선택
   - 좌측(Old), 우측(New) 화면에서 차이 확인
   - 색상으로 변경 사항 확인:
     - 🟢 초록색: 추가된 셀
     - 🔴 빨간색: 삭제된 셀
     - 🟡 노란색: 수정된 셀
   - 상단 통계 패널에서 변경 통계 확인

5. **스크롤 및 탐색**
   - 마우스 스크롤 또는 스크롤바 사용
   - 좌우 화면이 자동으로 동기화되어 스크롤

## 구현 세부사항

### MVVM 패턴 적용
- **Model**: Core 라이브러리의 데이터 모델 (Cell, Sheet, DiffResult 등)
- **View**: XAML 파일 (MainWindow.xaml)
- **ViewModel**: MainViewModel, SheetComparisonViewModel

**장점:**
- UI와 비즈니스 로직 분리
- 테스트 용이성
- 유지보수성 향상

### 비동기 처리
```csharp
[RelayCommand]
private async Task CompareAsync()
{
    IsLoading = true;
    try
    {
        var oldFile = await Task.Run(() => _excelReader.ReadExcelFile(OldFilePath));
        var newFile = await Task.Run(() => _excelReader.ReadExcelFile(NewFilePath));
        var diffResults = await Task.Run(() => _diffEngine.CompareExcelFiles(oldFile, newFile));
        // UI 업데이트
    }
    finally
    {
        IsLoading = false;
    }
}
```

**장점:**
- UI 블로킹 방지
- 로딩 인디케이터 표시 가능
- 사용자 경험 향상

### 성능 최적화

#### DataGrid 가상화
```xml
<DataGrid EnableRowVirtualization="True"
          EnableColumnVirtualization="True"
          VirtualizingPanel.IsVirtualizing="True"
          VirtualizingPanel.VirtualizationMode="Recycling">
```

**효과:**
- 화면에 보이는 행만 렌더링
- 메모리 사용량 감소
- 대용량 파일 처리 가능

#### Dictionary 기반 셀 저장
```csharp
public class Sheet
{
    public Dictionary<CellAddress, Cell> Cells { get; set; } = new();
}
```

**효과:**
- 희소 행렬 효율적 저장
- O(1) 셀 접근 속도
- 빈 셀은 메모리에 저장하지 않음

## 개발 과정

### Phase 1: 프로젝트 기반 구조 (완료)
- ✅ 솔루션 및 3개 프로젝트 생성 (App, Core, Tests)
- ✅ NuGet 패키지 설치
- ✅ 기본 폴더 구조 생성

### Phase 2: Core 라이브러리 (완료)
- ✅ Models 클래스 구현 (Cell, Sheet, ExcelFile, DiffResult 등)
- ✅ ExcelReaderService 구현 (ClosedXML 기반)
- ✅ DiffEngineService 구현 (비교 알고리즘)
- ✅ 단위 테스트 준비

### Phase 3: WPF UI 기본 구조 (완료)
- ✅ MainWindow.xaml 레이아웃
- ✅ ViewModels 구현 (MVVM 패턴)
- ✅ Converters 구현 (색상 변환)

### Phase 4: 고급 UI 기능 (완료)
- ✅ DataGrid 동적 컬럼 생성
- ✅ 스크롤 동기화
- ✅ 셀별 색상 강조
- ✅ 통계 패널

### Phase 5: 성능 최적화 (완료)
- ✅ DataGrid 가상화
- ✅ 비동기 처리
- ✅ 로딩 인디케이터

## 향후 개선 사항

### 기능 추가
- [ ] **CSV Export**: 비교 결과를 CSV로 저장
- [ ] **필터링**: 변경된 셀만 표시
- [ ] **검색 기능**: 특정 값 검색
- [ ] **수식 비교**: 값 vs 수식 비교 옵션
- [ ] **포맷 비교**: 셀 색상, 폰트 등 비교
- [ ] **Excel 내보내기**: 비교 결과를 Excel 파일로 저장

### UI/UX 개선
- [ ] **다크 모드**: 테마 전환 기능
- [ ] **설정 저장**: 최근 파일 목록, 창 크기 등
- [ ] **드래그 앤 드롭**: 파일 드래그로 열기
- [ ] **키보드 단축키**: 빠른 탐색 지원
- [ ] **줌 기능**: 셀 크기 조절

### 고급 기능
- [ ] **3-way 비교**: Base, Old, New 3개 파일 비교
- [ ] **Git 통합**: Git diff 스타일 뷰
- [ ] **병합 기능**: 변경 사항 선택적 병합
- [ ] **대용량 파일 최적화**: 스트리밍 방식 처리
- [ ] **명령줄 인터페이스**: CLI 지원

## 알려진 제한사항

1. **Windows 전용**: WPF는 Windows에서만 실행 가능
2. **.xlsx만 지원**: 구 버전 .xls 파일 미지원 (ClosedXML 제약)
3. **시트 이름 매칭**: 시트 이름이 다르면 비교 불가
4. **매크로 미지원**: VBA 매크로 비교 불가
5. **차트/이미지**: 그래픽 요소는 비교하지 않음

## 라이선스 정보

### ClosedXML 라이선스
- **라이선스**: MIT License
- **상업적 사용**: 가능 (무료)
- **재배포**: 가능
- **출처**: https://github.com/ClosedXML/ClosedXML

### 기타 라이브러리
- **CommunityToolkit.Mvvm**: MIT License
- **Microsoft.Xaml.Behaviors.Wpf**: MIT License
- **CsvHelper**: MS-PL / Apache 2.0

## 기술 참고 자료

### ClosedXML
- 문서: https://github.com/ClosedXML/ClosedXML/wiki
- 예제: https://github.com/ClosedXML/ClosedXML/wiki/Examples

### WPF MVVM
- CommunityToolkit.Mvvm: https://learn.microsoft.com/en-us/dotnet/communitytoolkit/mvvm/
- WPF 가이드: https://learn.microsoft.com/en-us/dotnet/desktop/wpf/

### .NET 8.0
- 공식 문서: https://learn.microsoft.com/en-us/dotnet/core/whats-new/dotnet-8

## 문제 해결

### 빌드 오류
```bash
# NuGet 패키지 복원
dotnet restore

# 빌드 캐시 정리
dotnet clean
dotnet build
```

### 실행 오류
- **.NET 8.0 SDK 미설치**: https://dotnet.microsoft.com/download 에서 다운로드
- **파일 접근 오류**: Excel 파일이 다른 프로그램에서 열려있는지 확인
- **메모리 부족**: 매우 큰 파일의 경우 64비트 모드로 실행

## 기여 방법

이 프로젝트는 Claude Code로 생성되었으며, 다음과 같이 기여할 수 있습니다:

1. 버그 리포트 또는 기능 제안
2. 코드 개선 (리팩토링, 성능 최적화)
3. 문서 개선
4. 단위 테스트 추가

## 개발 환경

- **생성 도구**: Claude Code (Anthropic)
- **개발 일자**: 2025년 12월
- **.NET 버전**: .NET 8.0
- **IDE**: Visual Studio 2022 / VS Code 권장

## 연락처

프로젝트 관련 문의사항이나 버그 리포트는 GitHub Issues를 통해 제출해주세요.

---

**Generated with** [Claude Code](https://claude.com/claude-code) 🤖
