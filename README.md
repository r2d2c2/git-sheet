# Git Sheet

JDK 25와 JavaFX로 만든 **Git 친화적인 데스크톱 스프레드시트**입니다. 셀 값·수식·기본 서식을 줄 단위 텍스트로 저장해서 일반 Git diff와 커밋 기록으로 문서 변경을 확인합니다.

현재 버전은 **0.1 초기 구현**입니다. Microsoft Excel의 대부분 기능을 완성한 제품이 아니며, 아래 지원 범위와 제한을 먼저 확인하세요.

![Git Sheet 실행 화면](docs/screenshot.png)

## 빠른 시작

필수: **JDK 25**, Git. Maven은 Wrapper가 내려받습니다. 첫 실행에는 인터넷이 필요합니다.

Windows PowerShell:

```powershell
.\mvnw.cmd clean verify
.\mvnw.cmd javafx:run
```

macOS / Linux:

```sh
./mvnw clean verify
./mvnw javafx:run
```

Windows에서 `run.cmd`를 실행해도 됩니다. `JAVA_HOME`은 JDK 25 설치 폴더를 가리켜야 합니다. Maven 3.9의 Guice가 JDK 25에서 출력하는 `sun.misc.Unsafe` 경고는 앱 코드에서 발생하는 경고가 아닙니다.

### IntelliJ IDEA

1. 이 폴더의 `pom.xml`을 **프로젝트로 열기**합니다.
2. Project SDK와 Maven Runner JRE를 **JDK 25**로 설정합니다.
3. Maven 프로젝트를 동기화합니다.
4. Maven 도구 창 → Plugins → javafx → `javafx:run`을 실행합니다. 공유 실행 설정 `Git Sheet`도 포함되어 있습니다.

## 구현된 기능

| 기능 | 지원 범위 |
| --- | --- |
| 셀 편집 | 텍스트·숫자·불리언·수식, 더블 클릭/F2 편집, 수식 입력줄 |
| 수식 | Apache POI가 지원하는 SUM, AVERAGE, IF 등, 셀·범위·다른 시트 참조, 변경 시 캐시 무효화 |
| 시트 | 추가·이름 변경·삭제, 하단 시트 탭 |
| 편집 이력 | 트랜잭션 단위 실행 취소/다시 실행, 오류 시 전체 작업 복구 |
| 클립보드 | 직사각형 TSV 복사·붙여넣기, 탭·따옴표·줄바꿈 처리 |
| 기본 서식 | 굵게·기울임·배경, 숫자·통화·백분율·날짜 표시 |
| 정렬·필터 | **현재 화면 1,000행**을 대상으로 하는 보기 정렬·텍스트 필터. 원본 셀 위치는 바꾸지 않음 |
| 찾기·이동 | 현재 시트의 값/수식 찾기, 주소 입력 이동, 1,000행 단위 이동 |
| 차트 | 선택 영역의 첫 열(이름)과 마지막 열(수치)로 막대 차트 미리보기, 최대 200행 |
| 파일 | `.gsheet` 저장/열기, `.xlsx` 가져오기/내보내기, UTF-8 CSV |
| Git | 저장된 문서의 diff·기록, 선택 문서만 커밋, 필요한 경우 문서 폴더에 저장소 초기화 |

앱은 예제 구매 목록으로 시작합니다. **파일 → 새 문서**로 빈 문서를 만들 수 있습니다.

## Git으로 문서 관리하기

1. 개인 문서용 폴더를 만들고 `.gsheet`로 저장합니다. 이 공개 소스 저장소와 개인 문서 저장소는 별도로 운영하세요.
2. Git 사용자 이름·이메일이 없다면 먼저 설정합니다.
3. **Git → 문서 저장 후 커밋**에서 메시지를 입력합니다.
4. 셀을 수정하고 저장한 뒤 **Git → 문서 변경 내용 / 기록**으로 비교합니다.

```sh
git config --global user.name "Your Name"
git config --global user.email "you@example.com"
```

Git 화면은 디스크에 저장된 파일을 비교합니다. 다른 파일의 staged 변경은 문서 커밋에 포함하지 않습니다. 브랜치·push·pull·merge 및 충돌 해결은 현재 Git CLI/IntelliJ에서 처리합니다. Git 인증 정보는 앱이 저장하지 않습니다.

`.gsheet`는 UTF-8 JSON Lines입니다. 한 셀의 변경은 해당 셀 레코드 한 줄에 나타나고 계산 결과 캐시는 저장되지 않습니다. 동일 셀을 두 브랜치에서 수정하면 일반 Git 충돌이 날 수 있으며, 충돌 표시를 해결해야 다시 열 수 있습니다.

## Excel 호환성 및 알려진 제한

- **VBA/매크로, 피벗 테이블 편집, Power Query, Power Pivot, 실시간 공동 편집, Solver, 전체 Excel 함수 집합, 동적 배열은 미지원**입니다.
- `.gsheet`는 값·수식·일부 기본 서식·열 너비·행 높이·병합/틀 고정 메타데이터만 저장합니다. 그림·고급 차트·이름 정의·테이블·조건부 서식·유효성 검사·댓글·하이퍼링크·인쇄 설정 등은 보존하지 않습니다. 테마 색상, 글꼴의 고급 속성 및 테두리 색도 완전 보존하지 않습니다.
- 이름 정의/외부 연결/구조화 참조에 의존한 수식은 `.gsheet` 변환 후 계산되지 않을 수 있습니다. **원본 Excel 파일을 보관하세요.** 가져오기는 원본을 덮어쓰지 않습니다.
- 병합·틀 고정 메타데이터는 왕복 저장하지만 화면에서는 병합/고정 렌더링을 하지 않습니다. 열 너비 드래그 변경은 화면에만 적용됩니다.
- 수식 미지원 오류는 `#UNSUPPORTED!`로 표시합니다. 일반 Excel 계산 오류(`#DIV/0!` 등)는 그대로 표시합니다.
- 붙여넣기는 수식 텍스트를 그대로 붙입니다. Excel처럼 상대 참조를 자동 이동하지 않습니다. 선행 0이 있는 값은 문자열이며, 앞에 `'`를 붙이면 강제로 텍스트 입력이 됩니다.
- CSV 가져오기는 수식 실행을 막기 위해 모든 필드를 텍스트로 읽습니다. CSV 내보내기는 표시값만 기록합니다. 다른 프로그램에서 CSV를 열 때 그 프로그램이 수식처럼 해석할 수 있는 텍스트가 포함될 수 있습니다.
- 그리드는 **1,000행 × 52열 창**을 가상 렌더링하고 주소 이동으로 Excel 좌표 범위에 접근합니다. 워크북 자체와 실행 취소 스냅샷은 메모리에 있으므로 대용량 파일의 성능 보장은 아닙니다.
- 실행 취소는 최대 30단계, 각 방향 합계 32MiB를 목표로 제한합니다. 가장 최근 스냅샷 하나는 한도를 넘더라도 보존합니다. 큰 워크북 저장/가져오기/수식 계산은 UI를 일시 정지시킬 수 있습니다.
- 차트는 일회성 미리보기이며 파일에 저장되지 않습니다. 날짜 입력 자동 인식, 행/열 삽입·삭제, 채우기 핸들, 인쇄/PDF, XLS 바이너리 형식은 아직 지원하지 않습니다.

## 기술 선택

- JDK 25: records, switch expressions, unnamed variables, virtual threads. Preview 기능 없음.
- JavaFX 25.0.4: JDK 25에서 동작하는 25 계열 안정 릴리스, 가상화 TableView.
- Apache POI 5.5.1: XLSX 모델·수식·서식.
- Gson 2.13.2 / Commons CSV 1.14.1: 텍스트 저장과 CSV 파싱.
- 문서 저장은 같은 디렉터리 임시 파일을 사용하고 가능한 파일시스템에서는 원자적으로 교체합니다.
- Git I/O는 가상 스레드에서 실행합니다. 셸 문자열을 조합하지 않고 인자 배열로 실행합니다.
- [OpenJFX Maven 가이드](https://openjfx.io/openjfx-docs/#maven), [JavaFX 25 요구사항](https://openjfx.io/highlights/25/), [Apache POI](https://poi.apache.org/).

## 검증

```powershell
.\mvnw.cmd clean verify
.\mvnw.cmd javafx:run '-Djavafx.args=--smoke-test'
```

Windows / JDK 25에서 단위·통합 테스트 14개와 실제 JavaFX 창 smoke test를 검증했습니다. Smoke test는 셀 편집 컨트롤, 수식 갱신, 실행 취소/다시 실행, 필터, 주소 이동을 확인한 후 종료합니다. Linux CI에서는 Xvfb로 실행합니다. Microsoft Excel 애플리케이션 자체와의 수동 상호운용 테스트는 수행하지 않았습니다.

[설계 및 확장 계획](docs/ARCHITECTURE.md)을 참고하세요.
