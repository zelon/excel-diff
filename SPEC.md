# Excel Diff GUI 도구 구현 계획

## 개요
Microsoft Excel 파일(*.xlsx)을 비교하는 WPF GUI 도구를 구현합니다.
- **기술 스택**: C# + WPF (.NET 8)
- **Excel 라이브러리**: ClosedXML (MIT 라이선스)
- **아키텍처**: MVVM 패턴

## 주요 기능
1. 좌우 분할 화면 (Old/New 동시 표시)
2. 모든 시트 자동 매칭 (같은 이름 기준)
3. 색상 강조 (추가=초록, 삭제=빨강, 수정=노랑)
4. 가로/세로 스크롤 동기화
5. 통계 요약 (변경된 셀 개수)

## 프로젝트 구조

```
ExcelDiff/
├── ExcelDiff.sln
└── src/
    ├── ExcelDiff.App/           # WPF 애플리케이션
    │   ├── ViewModels/
    │   ├── Views/
    │   ├── Controls/
    │   └── Converters/
    ├── ExcelDiff.Core/          # 비즈니스 로직
    │   ├── Models/
    │   └── Services/
    └── ExcelDiff.Tests/         # 단위 테스트
```

## 핵심 컴포넌트

### 1. Core Models
- **Cell**: 셀 주소, 값, 수식 정보
- **Sheet**: 시트 이름, 행/열 개수, 셀 컬렉션
- **ExcelFile**: 파일 경로, 시트 리스트
- **DiffType** (enum): Unchanged, Added, Deleted, Modified
- **CellDiff**: 셀별 차이 정보 (주소, 타입, 이전값, 새값)
- **DiffResult**: 시트별 비교 결과 + 통계
- **ComparisonStatistics**: 추가/삭제/수정 셀 개수, 변경률

### 2. Core Services
- **IExcelReader / ExcelReaderService**: ClosedXML을 사용한 Excel 파일 읽기
- **IDiffEngine / DiffEngineService**: 셀별 비교 알고리즘 구현

### 3. ViewModels (MVVM)
- **MainViewModel**: 파일 로드, 비교 실행 커맨드
- **SheetComparisonViewModel**: 시트별 데이터, 스크롤 오프셋
- **CellViewModel**: 셀 값, DiffType, 배경색
- **StatisticsViewModel**: 통계 데이터 표시

### 4. Custom Controls
- **SyncDataGrid**: 스크롤 동기화를 위한 커스텀 DataGrid

## 구현 단계

### Phase 1: 프로젝트 기반 구조
1. 솔루션 및 3개 프로젝트 생성 (App, Core, Tests)
2. NuGet 패키지 설치:
   - **ExcelDiff.Core**: ClosedXML, CsvHelper
   - **ExcelDiff.App**: CommunityToolkit.Mvvm, Microsoft.Xaml.Behaviors.Wpf
   - **ExcelDiff.Tests**: xUnit, FluentAssertions, Moq
3. 기본 폴더 구조 생성

### Phase 2: Core 라이브러리
1. Models 클래스 구현 (Cell, Sheet, ExcelFile, DiffResult 등)
2. ExcelReaderService 구현:
   - ClosedXML을 사용하여 .xlsx 파일 읽기
   - 모든 시트 로드
   - 셀 데이터 추출 (값, 수식, 포맷)
3. DiffEngineService 구현:
   - 시트 이름 기준 자동 매칭
   - 핵심 비교 알고리즘:
     ```
     모든 셀 주소 = Old 셀 주소 ∪ New 셀 주소
     각 셀 주소에 대해:
       - Old에만 존재 → Deleted (빨강)
       - New에만 존재 → Added (초록)
       - 둘 다 존재 + 값 다름 → Modified (노랑)
       - 둘 다 존재 + 값 같음 → Unchanged
     ```
   - 통계 계산 (추가/삭제/수정 개수)
4. 단위 테스트 작성

### Phase 3: WPF UI 기본 구조
1. MainWindow.xaml 레이아웃:
   - 상단: 파일 선택 영역 (Old/New 파일 경로, 선택 버튼, 비교 버튼)
   - 중간: TabControl (시트별 탭)
   - 하단: 통계 패널
2. ViewModels 구현:
   - MainViewModel (RelayCommand, 파일 다이얼로그)
   - SheetComparisonViewModel
3. 사이드바이사이드 레이아웃:
   - Grid 2열 분할 (GridSplitter 포함)
   - 왼쪽: Old DataGrid
   - 오른쪽: New DataGrid

### Phase 4: 스크롤 동기화
1. SyncDataGrid 커스텀 컨트롤 구현:
   - DependencyProperty로 스크롤 오프셋 바인딩
   - ScrollViewer의 ScrollChanged 이벤트 처리
   - 양방향 동기화 (왼쪽↔오른쪽)

### Phase 5: 차이 시각화
1. DiffTypeToColorConverter 구현:
   - Added → LightGreen (#90EE90)
   - Deleted → LightCoral (#F08080)
   - Modified → Yellow (#FFFF00)
   - Unchanged → White
2. DataGrid CellStyle 적용:
   - 배경색 바인딩
   - ToolTip으로 변경 전 값 표시
3. 시트 탭에 변경 개수 표시

### Phase 6: 통계 및 부가 기능
1. StatisticsPanel 구현:
   - 전체 변경 사항 요약
   - 시트별 통계
   - 변경률 표시
2. 시트 탭 UI 개선:
   - 변경이 있는 시트 강조
3. CSV Export 기능 (선택사항):
   - 비교 결과를 CSV로 저장

### Phase 7: 성능 최적화
1. DataGrid 가상화 설정:
   ```xml
   <DataGrid VirtualizingPanel.IsVirtualizing="True"
             VirtualizingPanel.VirtualizationMode="Recycling">
   ```
2. 비동기 처리:
   - async/await로 파일 로드 및 비교
   - Progress 표시 (로딩 인디케이터)
3. 대용량 파일 처리:
   - 시트는 선택될 때만 비교 (지연 로딩)
   - Dictionary<CellAddress, Cell>로 희소 행렬 구현

### Phase 8: 에러 처리 및 완성도
1. 예외 처리:
   - 파일 읽기 실패 (형식 오류, 권한 문제)
   - 손상된 Excel 파일
   - 메모리 부족
2. 사용자 피드백:
   - 로딩 상태 표시
   - 에러 메시지 다이얼로그
   - 빈 결과 처리 (차이 없음)

## 핵심 파일 경로

구현 시 가장 먼저 작업할 파일들:

1. **D:\git\ExcelDiff\src\ExcelDiff.Core\Models\Cell.cs** - 기본 데이터 구조
2. **D:\git\ExcelDiff\src\ExcelDiff.Core\Models\DiffResult.cs** - 비교 결과 모델
3. **D:\git\ExcelDiff\src\ExcelDiff.Core\Services\ExcelReaderService.cs** - ClosedXML 통합
4. **D:\git\ExcelDiff\src\ExcelDiff.Core\Services\DiffEngineService.cs** - 핵심 비교 로직
5. **D:\git\ExcelDiff\src\ExcelDiff.App\ViewModels\MainViewModel.cs** - MVVM 컨트롤러
6. **D:\git\ExcelDiff\src\ExcelDiff.App\MainWindow.xaml** - 메인 UI

## 프로젝트 생성 명령어

```bash
# 솔루션 생성
dotnet new sln -n ExcelDiff

# 프로젝트 생성
dotnet new wpf -n ExcelDiff.App -f net8.0-windows -o src/ExcelDiff.App
dotnet new classlib -n ExcelDiff.Core -f net8.0 -o src/ExcelDiff.Core
dotnet new xunit -n ExcelDiff.Tests -f net8.0 -o src/ExcelDiff.Tests

# 솔루션에 추가
dotnet sln add src/ExcelDiff.App/ExcelDiff.App.csproj
dotnet sln add src/ExcelDiff.Core/ExcelDiff.Core.csproj
dotnet sln add src/ExcelDiff.Tests/ExcelDiff.Tests.csproj

# 프로젝트 참조
dotnet add src/ExcelDiff.App reference src/ExcelDiff.Core
dotnet add src/ExcelDiff.Tests reference src/ExcelDiff.Core

# NuGet 패키지 설치
dotnet add src/ExcelDiff.Core package ClosedXML
dotnet add src/ExcelDiff.Core package CsvHelper
dotnet add src/ExcelDiff.App package CommunityToolkit.Mvvm
dotnet add src/ExcelDiff.App package Microsoft.Xaml.Behaviors.Wpf
dotnet add src/ExcelDiff.Tests package FluentAssertions
dotnet add src/ExcelDiff.Tests package Moq
```

## 기술적 고려사항

### ClosedXML vs EPPlus
- **ClosedXML 선택 이유**: MIT 라이선스 (무료, 상업용 가능)
- **단점**: EPPlus보다 약간 느릴 수 있음
- **해결**: 비동기 처리 + 가상화로 성능 보완

### Excel 처리 방식
- **직접 비교 방식 채택**: Excel → 메모리 → 비교
- **CSV 변환 방식 미채택**: 중간 파일 생성 불필요, 포맷 정보 유지
- **CSV Export**: 비교 결과 저장 기능으로만 제공

### UI 반응성
- 모든 I/O 작업은 비동기 처리
- 비교 중 UI 블로킹 방지
- 진행률 표시 (IProgress<int>)

## 예상 산출물

1. **실행 파일**: ExcelDiff.App.exe
2. **기능**:
   - Excel 파일 2개 선택
   - 모든 시트 자동 비교
   - 좌우 분할 뷰로 차이 시각화
   - 색상으로 변경 사항 강조
   - 통계 요약 표시
   - (선택) CSV로 결과 내보내기
