# ExcelDiff 테스트 프로젝트

## 개요

ExcelDiff.Core 라이브러리의 단위 테스트 프로젝트입니다.

## 테스트 구조

```
ExcelDiff.Tests/
├── Services/
│   ├── ExcelReaderServiceTests.cs    # Excel 파일 읽기 테스트 (6개)
│   └── DiffEngineServiceTests.cs     # 비교 알고리즘 테스트 (16개)
├── Models/
│   ├── CellAddressTests.cs           # CellAddress 모델 테스트 (6개)
│   └── ComparisonStatisticsTests.cs  # 통계 계산 테스트 (5개)
├── Helpers/
│   └── SampleDataGenerator.cs        # 샘플 Excel 파일 생성 유틸리티
└── TestData/
    └── (테스트 실행 시 임시 파일 생성됨)
```

## 테스트 실행

### 모든 테스트 실행
```bash
dotnet test
```

### 특정 테스트만 실행
```bash
# ExcelReaderService 테스트만
dotnet test --filter "FullyQualifiedName~ExcelReaderServiceTests"

# DiffEngineService 테스트만
dotnet test --filter "FullyQualifiedName~DiffEngineServiceTests"
```

### 코드 커버리지와 함께 실행
```bash
dotnet test /p:CollectCoverage=true /p:CoverletOutputFormat=opencover
```

## 테스트 통계

- **총 테스트 수**: 27개
- **Services 테스트**: 22개
  - ExcelReaderServiceTests: 6개
  - DiffEngineServiceTests: 16개
- **Models 테스트**: 11개
  - CellAddressTests: 6개
  - ComparisonStatisticsTests: 5개

## 샘플 데이터

### 자동 생성
테스트 실행 시 필요한 샘플 Excel 파일이 임시 디렉토리에 자동으로 생성됩니다.

### 수동 생성 (선택사항)
루트 디렉토리의 `GenerateSamples.csx` 스크립트를 실행하면 TestData 폴더에 샘플 파일이 생성됩니다:

```bash
# dotnet-script 설치 (한 번만 필요)
dotnet tool install -g dotnet-script

# 샘플 파일 생성
dotnet script GenerateSamples.csx
```

또는 SampleDataGenerator를 직접 호출:

```csharp
using ExcelDiff.Tests.Helpers;

var testDataPath = "원하는 경로";
SampleDataGenerator.GenerateSampleFiles(testDataPath);
```

## 테스트 커버리지

주요 테스트 시나리오:

### ExcelReaderService
- ✅ 파일 존재 여부 검증
- ✅ 모든 시트 읽기
- ✅ 셀 값 정확히 읽기
- ✅ 특정 시트만 읽기
- ✅ 존재하지 않는 시트 에러 처리
- ✅ 빈 시트 처리

### DiffEngineService
- ✅ 추가된 셀 감지
- ✅ 삭제된 셀 감지
- ✅ 수정된 셀 감지
- ✅ 변경 없는 셀 구분
- ✅ 통계 정확히 계산
- ✅ 시트 이름으로 매칭
- ✅ 삭제된 시트 처리
- ✅ 추가된 시트 처리
- ✅ 빈 시트 처리
- ✅ 셀 정렬 (행→열 순서)

### Models
- ✅ CellAddress 초기화
- ✅ CellAddress 동등성 비교
- ✅ CellAddress 해시코드
- ✅ CellAddress ToString
- ✅ Dictionary 키로 사용
- ✅ 통계 계산 (변경률, 변경 셀 수)
- ✅ 엣지 케이스 (0으로 나누기 등)

## 기술 스택

- **xUnit**: 테스트 프레임워크
- **FluentAssertions**: 읽기 쉬운 테스트 검증
- **Moq**: Mocking 프레임워크 (향후 사용 예정)
- **ClosedXML**: 테스트용 Excel 파일 생성

## 향후 추가 예정

- [ ] 통합 테스트 (전체 워크플로우)
- [ ] 성능 테스트 (대용량 파일)
- [ ] 벤치마크 테스트
- [ ] ViewModel 테스트
- [ ] UI 테스트 (WPF UI Automation)
