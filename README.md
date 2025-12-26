# ExcelDiff 📊

> Excel 파일을 시각적으로 비교하는 WPF 데스크톱 애플리케이션

[![.NET](https://img.shields.io/badge/.NET-8.0-512BD4?logo=dotnet)](https://dotnet.microsoft.com/)
[![C#](https://img.shields.io/badge/C%23-12.0-239120?logo=csharp)](https://docs.microsoft.com/en-us/dotnet/csharp/)
[![WPF](https://img.shields.io/badge/WPF-Windows-0078D4?logo=windows)](https://docs.microsoft.com/en-us/dotnet/desktop/wpf/)
[![License](https://img.shields.io/badge/License-ClosedXML_MIT-green)](https://github.com/ClosedXML/ClosedXML)

## ✨ 주요 기능

- 🔄 **사이드바이사이드 비교** - Old/New 파일을 좌우로 나란히 표시
- 🎨 **색상 강조** - 추가(🟢), 삭제(🔴), 수정(🟡) 셀을 색상으로 구분
- 📑 **자동 시트 매칭** - 같은 이름의 시트를 자동으로 매칭
- 🔗 **스크롤 동기화** - 좌우 화면이 동시에 스크롤
- 📊 **상세 통계** - 변경된 셀 개수 및 변경률 제공
- ⚡ **성능 최적화** - 가상화를 통한 대용량 파일 처리

## 🚀 빠른 시작

### 요구 사항

- Windows 10/11
- [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

### 설치 및 실행

```bash
# 리포지토리 클론
git clone https://github.com/yourusername/ExcelDiff.git
cd ExcelDiff

# 빌드
dotnet build

# 실행
dotnet run --project src/ExcelDiff.App
```

### 또는 릴리스 빌드

```bash
dotnet build -c Release
cd src/ExcelDiff.App/bin/Release/net8.0-windows
ExcelDiff.App.exe
```

## 📖 사용 방법

1. **애플리케이션 실행**
2. **Old 파일 선택** - "찾아보기..." 버튼으로 이전 버전 Excel 파일 선택
3. **New 파일 선택** - "찾아보기..." 버튼으로 새 버전 Excel 파일 선택
4. **비교 버튼 클릭** - 파일 비교 시작
5. **결과 확인** - 시트별 탭에서 차이점 확인

### 색상 의미

| 색상 | 의미 |
|------|------|
| 🟢 초록색 (`#90EE90`) | 추가된 셀 |
| 🔴 빨간색 (`#F08080`) | 삭제된 셀 |
| 🟡 노란색 (`#FFFF00`) | 수정된 셀 |
| ⚪ 흰색 | 변경 없음 |

## 🛠️ 기술 스택

### 프레임워크
- **.NET 8.0** - 최신 .NET 플랫폼
- **WPF** - Windows Presentation Foundation

### 주요 라이브러리
- **[ClosedXML](https://github.com/ClosedXML/ClosedXML)** `0.105.0` - Excel 파일 처리 (MIT)
- **[CommunityToolkit.Mvvm](https://github.com/CommunityToolkit/dotnet)** `8.4.0` - MVVM 패턴
- **[Microsoft.Xaml.Behaviors.Wpf](https://github.com/Microsoft/XamlBehaviorsWpf)** `1.1.135` - WPF Behaviors
- **[CsvHelper](https://github.com/JoshClose/CsvHelper)** `33.1.0` - CSV 처리

### 테스트
- **xUnit** - 단위 테스트
- **FluentAssertions** - 테스트 검증
- **Moq** - Mocking

## 📂 프로젝트 구조

```
ExcelDiff/
├── src/
│   ├── ExcelDiff.Core/          # 비즈니스 로직
│   │   ├── Models/              # 데이터 모델
│   │   ├── Services/            # Excel 읽기 및 비교
│   │   └── Enums/               # 열거형
│   ├── ExcelDiff.App/           # WPF 애플리케이션
│   │   ├── ViewModels/          # MVVM 뷰모델
│   │   ├── Converters/          # 데이터 변환
│   │   ├── Helpers/             # 헬퍼 클래스
│   │   └── MainWindow.xaml      # 메인 UI
│   └── ExcelDiff.Tests/         # 단위 테스트
├── SPEC.md                       # 구현 명세서
├── CLAUDE.md                     # 상세 문서
└── README.md                     # 이 파일
```

## 🏗️ 아키텍처

### MVVM 패턴

```
┌─────────────┐      ┌──────────────┐      ┌─────────────┐
│    View     │─────▶│  ViewModel   │─────▶│    Model    │
│  (XAML)     │◀─────│   (C#)       │◀─────│   (Core)    │
└─────────────┘      └──────────────┘      └─────────────┘
   MainWindow         MainViewModel         ExcelFile
                      SheetViewModel        DiffResult
```

### 핵심 알고리즘

```csharp
// 셀별 비교 알고리즘
모든 셀 주소 = Old 셀 ∪ New 셀

for each 셀 주소:
    if (Old에만 존재)     → Deleted (빨강)
    if (New에만 존재)     → Added (초록)
    if (값이 다름)        → Modified (노랑)
    if (값이 같음)        → Unchanged (흰색)
```

## 🎯 핵심 기능 상세

### 1. 동적 컬럼 생성

DataGrid의 컬럼을 런타임에 자동 생성하여 Excel의 A, B, C... 형식으로 표시합니다.

```csharp
// Attached Property 사용
helpers:DataGridHelper.AutoGenerateCellColumns="True"
```

### 2. 스크롤 동기화

좌우 DataGrid의 스크롤을 자동으로 동기화합니다.

```csharp
// ScrollGroup으로 동기화
helpers:ScrollSyncHelper.ScrollGroup="SheetComparison"
```

### 3. 성능 최적화

- **가상화**: 화면에 보이는 행만 렌더링
- **비동기 처리**: UI 블로킹 방지
- **Dictionary 기반 저장**: 빈 셀은 메모리에 저장하지 않음

```xml
<DataGrid VirtualizingPanel.IsVirtualizing="True"
          VirtualizingPanel.VirtualizationMode="Recycling">
```

## 📊 통계 정보

비교 결과는 다음 통계를 제공합니다:

- **전체 셀 개수** - 비교 대상 전체 셀
- **추가된 셀** - Old에 없고 New에만 있는 셀
- **삭제된 셀** - Old에만 있고 New에 없는 셀
- **수정된 셀** - 값이 변경된 셀
- **변경률** - (추가+삭제+수정) / 전체 × 100%

## 🔧 개발

### 빌드

```bash
# 개발 모드
dotnet build

# 릴리스 모드
dotnet build -c Release
```

### 테스트

```bash
# 모든 테스트 실행
dotnet test

# 커버리지와 함께 실행
dotnet test /p:CollectCoverage=true
```

### 디버깅

Visual Studio 2022 또는 VS Code에서 `ExcelDiff.sln` 파일을 열어 디버깅할 수 있습니다.

## 🚧 향후 개선 사항

- [ ] CSV Export 기능
- [ ] 변경된 셀만 필터링
- [ ] 수식 비교 모드
- [ ] 다크 모드 테마
- [ ] 드래그 앤 드롭 파일 열기
- [ ] 키보드 단축키
- [ ] Excel 파일로 결과 내보내기
- [ ] 3-way 비교 (Base, Old, New)

## ⚠️ 알려진 제한사항

1. **Windows 전용** - WPF는 Windows에서만 실행됩니다
2. **.xlsx만 지원** - .xls (구 버전) 파일은 지원하지 않습니다
3. **시트 이름 매칭** - 시트 이름이 같아야 비교됩니다
4. **매크로 미지원** - VBA 매크로는 비교하지 않습니다
5. **그래픽 요소** - 차트, 이미지는 비교하지 않습니다

## 🤝 기여하기

기여는 언제나 환영합니다! 다음 방법으로 기여할 수 있습니다:

1. 이 저장소를 Fork
2. Feature 브랜치 생성 (`git checkout -b feature/AmazingFeature`)
3. 변경사항 커밋 (`git commit -m 'Add some AmazingFeature'`)
4. 브랜치에 Push (`git push origin feature/AmazingFeature`)
5. Pull Request 생성

## 📝 라이선스

이 프로젝트는 다음 라이선스를 사용하는 오픈소스 라이브러리를 포함합니다:

- **ClosedXML** - [MIT License](https://github.com/ClosedXML/ClosedXML/blob/develop/LICENSE)
- **CommunityToolkit.Mvvm** - [MIT License](https://github.com/CommunityToolkit/dotnet/blob/main/License.md)
- **Microsoft.Xaml.Behaviors.Wpf** - [MIT License](https://github.com/Microsoft/XamlBehaviorsWpf/blob/master/LICENSE)

## 📚 문서

- [SPEC.md](SPEC.md) - 구현 명세서 (기술 사양)
- [CLAUDE.md](CLAUDE.md) - 상세 개발 문서 (아키텍처, 구현 세부사항)

## 🙏 감사의 말

이 프로젝트는 다음 훌륭한 오픈소스 프로젝트들을 기반으로 만들어졌습니다:

- [ClosedXML](https://github.com/ClosedXML/ClosedXML) - Excel 파일 처리
- [.NET Community Toolkit](https://github.com/CommunityToolkit/dotnet) - MVVM 도구
- [Microsoft XAML Behaviors](https://github.com/Microsoft/XamlBehaviorsWpf) - WPF Behaviors

## 📞 문의

문제가 발생하거나 질문이 있으시면 [GitHub Issues](https://github.com/yourusername/ExcelDiff/issues)를 통해 알려주세요.

---

<div align="center">

**Made with ❤️ using [Claude Code](https://claude.com/claude-code)**

🤖 *Generated by Claude Sonnet 4.5*

</div>
