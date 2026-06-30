# CLAUDE.md — Excel Structure Analyzer

> 이 문서는 **다른 디바이스/새 세션에서 작업을 이어받기 위한 핸드오프 문서**다.
> Claude Code 의 대화 기록과 `~/.claude` 메모리는 이 기기 밖으로 따라오지 않으므로,
> 프로젝트의 맥락·결정·다음 단계를 모두 여기에 적어 둔다. **작업을 이어받으면 먼저 이 문서를 끝까지 읽을 것.**

---

## 1. 프로젝트 목표

회사에서 양식처럼 쓰이는 **다중 시트 .xlsx 파일**은 시트가 10~20개씩 되고 시트 간 수식으로 복잡하게
얽혀 있어, 처음 받은 사람이 구조를 파악하기 매우 어렵다(누가 만들었는지 추적도 안 됨).

이 도구는 그 엑셀 파일을 분석해서:
- 모든 시트 목록과 사용 범위·수식 통계를 파악하고
- 각 시트의 수식을 전부 추출하며
- **수식 안의 시트 참조를 따라 시트 간 데이터 흐름(입력→가공→산출) 관계를 시각화**한다.

결과는 단일 **HTML 리포트**로 보여준다.

## 2. ⚠️ 최우선 제약 — 데이터 프라이버시

분석 대상은 **회사 기밀 엑셀**이다. 따라서:
- **사용자 데이터가 기기 밖으로 절대 나가면 안 된다.** (Anthropic 서버, GitHub, 어떤 외부로도 X)
- 분석은 **100% 클라이언트 사이드(로컬/브라우저 내부)**로 동작해야 한다.
- 실제 업무 엑셀 파일, 그리고 그 데이터가 담긴 분석 결과 HTML(`*_analysis.html`)은
  **절대 git 에 커밋·푸시하지 말 것.** (`.gitignore` 에 이미 반영됨)
- 테스트는 합성 데이터 `sample.xlsx`(make_sample.py 로 생성, 회사 데이터 아님)로만 한다.

## 3. 현재 상태 (2026-06-30 기준)

- ✅ **분석 엔진** (`excel_analyzer/`) 완성 및 검증. **셀 단위 그래프까지 구축**(v2).
- ✅ **CLI** (`analyze.py`) 동작.
- ✅ **HTML 리포트** (`report.py`) — 인터랙티브 그래프 + 표 + 수식 상세, 디자인까지 다듬음.
  - 접철 토글 아이콘: 크기 키우고(시인성), 접힘 시 ◂(왼쪽) 방향으로 변경.
- ✅ **셀 단위 추적(v2)** 추가 — 시트 화살표 클릭 드릴다운 + 셀 추적기(흐름도). 아래 5번 7항·10절 참조.
- ✅ **웹페이지(Pyodide)** — `index.html` + `app.js`. 다중 파일 업로드 → [분석 시작] → 파일별
  리포트 탭 + 개별/ZIP 다운로드. **브라우저 실측 통과**. (셀 추적 v2 도 동일 엔진이라 자동 반영.)
- ✅ **GitHub Pages 배포 완료** — https://qriusquokka.github.io/Excel_Structure_Analyzer/ (정상 동작 확인).
  레포는 Public 전환됨(무료 플랜 Pages 사용 위해). 아래 6번 참조.

## 4. 아키텍처 & 파일 구조

```
excel_analyzer/
  __init__.py          공개 API (analyze_workbook, WorkbookAnalysis, SheetInfo, Dependency, CellNode, CellRef)
  formula_parser.py    수식에서 셀/시트 참조 추출 (openpyxl Tokenizer 기반)
  analyzer.py          openpyxl 로 시트/수식/의존관계/순환참조 + 셀 단위 그래프 분석 엔진
  report.py            분석 결과 → 단일 HTML 리포트 렌더링(시트 관계도 + 셀 추적기)
analyze.py             CLI 진입점:  python analyze.py 파일.xlsx
make_sample.py         검증용 합성 통합문서(sample.xlsx) 생성
requirements.txt       openpyxl
README.md              사용자용 설명서
```

**핵심 동작 원리**
- 수식은 openpyxl `data_only=False` 로 읽어 **수식 원문**을 얻는다.
- `formula_parser.py` 가 openpyxl 내장 **`Tokenizer`** 로 수식을 토큰화해 `OPERAND/RANGE`(셀·범위)만
  뽑는다(정규식보다 함수명·문자열·날짜 오탐이 적다). 시트명은 **실제 존재하는 시트 목록과 대조**,
  셀 부분은 셀/범위 형태인지 검증(정의된 이름 등 제외). 처리 케이스: 한글/공백 시트명 `'시트 이름'!`,
  따옴표 이스케이프 `''`, **동일시트 참조(`A1`)**, 3D `Sheet1:Sheet3!`(사이 시트 확장), 외부 `[1]Sheet!`.
  - `extract_cell_references()` 가 1차 함수, `extract_sheet_references()` 는 거기서 파생(단일 소스).
- **화살표 방향 = 데이터 흐름**: 셀이 `SheetB!A1` 을 참조하면 데이터는 SheetB → 현재시트로 흐른다.
  의존 엣지는 `from=참조된 셀/시트(출처)`, `to=참조하는 셀/시트(소비)` 로 저장.
- **셀 단위 그래프**(`analyzer.py`): `WorkbookAnalysis.cells` = `{ "시트!셀": CellNode(precedents/dependents) }`.
  범위(`A1:A100`)는 펼치지 않고 **범위 단위 유지**하되, 하류(dependents) 연결만 멤버 셀로 전개
  (상한 `MAX_RANGE_CELLS=4096`, 초과·전체열/행은 전개 생략 → 폭발 방지). `edge_cells` = 시트쌍 드릴다운용.
- `report.py` 의 시트 관계도·셀 추적기는 **외부 라이브러리 없는 순수 Canvas/SVG + JS**.
  → 다운로드한 HTML 을 인터넷 없이 열어도 드래그·화살표·추적이 동작한다.

## 5. 결정 로그 (왜 이렇게 했는가 — 번복 금지)

1. **출력 = HTML 리포트** (Mermaid 아님. Mermaid 는 폐기됨).
2. **분석 범위 = .xlsx/.xlsm 의 시트 단위 + 셀 단위 관계.**
   구버전 `.xls` 는 안내 메시지("xlsx 로 다시 저장 후 사용")와 함께 거부.
   셀 단위 추적은 v2(2026-06-30)에서 구현 완료 — 아래 7항·10절 참조.
3. **웹 배포 = GitHub Pages(정적) + Pyodide.**
   - JS 재작성(SheetJS) 대신 **Pyodide**(Python→WASM)로 기존 엔진을 그대로 브라우저에서 실행.
     이유: 단일 소스(CLI·웹 공용) 유지 + openpyxl 의 수식 읽기 충실도(SheetJS 는 공유수식 누락 위험).
   - Pyodide/CDN 로 런타임 로드(코드만 받음, 데이터 아님). 첫 로딩 ~10MB 감수 합의됨.
   - 파일은 브라우저 안에서만 처리 → 외부 업로드 0건(프라이버시 충족).
4. **그래프 시각화** = 드래그 가능한 **사각형 박스 노드** + **데이터 흐름 방향 화살표** + 드래그 시 엣지 빨강 강조.
   (초기 원형+무지개색 버전은 "유아스럽다"는 피드백으로 폐기.)
5. **색상 = 역할 기반 3+1색** (장식용 무지개 금지):
   - 입력단(input, 청록) / 가공(intermediate, 중립 회색) / 산출단(output, 그린) / 독립(isolated, 옅은 회색)
   - 전체 톤: 엑셀 그린 + 절제된 배경/테두리. "신뢰감·안정감" 우선, 캐주얼 이모지 자제.
6. **리포트 구성**:
   - 상단 통계 칩: 시트 수 / 총 수식 / 시트 간 참조 / **순환 참조 있음·없음**
   - 시트 목록 표: 시트명(역할 배지)·사용 범위·크기·수식 수·숫자 셀·텍스트 셀·수식 비율·참조하는 시트
   - **시트별 수식 상세**: 시트 기준 전체 수식, 접철식, **기본 접힘**, 펼치면 최대 ~12줄 후 스크롤
   - 큰 구획 4개(관계도/**셀 추적**/시트목록/수식상세): 각각 접철식, **기본 펼침**
   - 탭 UI 는 쓰지 않음(대신 접철식).
7. **셀 단위 추적(v2, 2026-06-30 확정·구현)** — 표현 규칙은 합성 데이터 시제품으로 검증 후 사용자 승인:
   - 시트 관계도 유지 + **화살표 클릭 → 두 시트 사이 셀 연결(소비셀 ← 출처셀) 드릴다운**.
   - **셀 추적기**: 데이터 왼쪽(출처)→오른쪽(소비) 흐름. 선택 셀 기준 전후 최대 **5단 창(WINDOW=5)**.
     입력값 셀이면 맨 왼쪽, 모두 받기만 하는 최종 셀이면 맨 오른쪽에 정렬.
   - 레벨은 **최장경로(longest-path)** 로 매김 — 단순 BFS 최단거리는 3D참조 등에서 의존셀이 같은 열에
     뭉개짐(폐기). 최장경로면 의존관계 셀이 반드시 다른 열에 놓여 **시트 내부 가공 흐름까지** 정확히 펼쳐짐.
   - **같은 시트의 같은 단계 셀은 시트 박스 하나로 묶어** 표시. 5단 초과 깊이는 양끝에 "N단 더 있음" 표기.
   - 박스 클릭 시 그 셀로 재중심이동(경로 breadcrumb). 범위/외부참조는 점선 말단 노드.
   - 동일시트 셀 참조(`A1`)는 셀 추적엔 포함, 시트 관계도엔 미포함(`self_refs` 로 집계).

## 6. Pyodide 웹페이지 — 구현 완료(브라우저 실측만 남음)

목표: GitHub Pages 에서 "누구나 접속 → 파일 첨부 → 브라우저 안 분석 → 결과 보기·다운로드".

구현됨(레포 루트):
```
index.html   드래그&드롭/선택 UI(여러 파일 가능), 프라이버시 안내, 엔진 상태 배너,
             "분석 대기" 목록, 파일 탭 전환, 결과 iframe 미리보기, 개별/ZIP 다운로드.
             색/톤은 리포트와 통일(엑셀 그린). JSZip(CDN)로 ZIP 생성.
app.js       Pyodide(v0.26.4, CDN) 로드 → micropip 로 openpyxl 설치 →
             ENGINE_FILES(excel_analyzer/*.py)를 fetch 해 가상 FS 에 기록 →
             업로드 파일 바이트를 /home/pyodide/_uploads 에 기록(업로드 없음) →
             _run_analysis(path)=analyze_workbook+render_html → {html,name,sheets,deps} JSON →
             파일별 탭 미리보기 + Blob 개별 다운로드 + JSZip 전체 다운로드. 분석 후 FS unlink.
```

UX 흐름(2026-06-26 개선): 파일 선택 → 바로 분석하지 않고 **"분석 대기" 목록**에 쌓음
(추가/제거/모두지우기) → **[분석 시작]** 클릭 → 여러 파일을 순차 분석(한 파일 실패해도
나머지 계속, 실패는 탭에 "실패" 표시) → 파일 탭으로 결과 전환 + 개별/ZIP 다운로드.
파일 간(A.xlsx↔B.xlsx) 관계는 분석 안 함 — 엔진은 한 통합문서 내부만 봄(A안: 파일별 개별 리포트).

### 로컬 테스트 방법(중요: file:// 는 fetch 가 막혀 안 됨 → 반드시 서버로)
```bash
python -m http.server 8765        # 레포 루트에서
# 브라우저로 http://localhost:8765 열기 → 준비 완료 뜨면 sample.xlsx 끌어다 놓기
```

### 완료됨
- ✅ **브라우저 실측 통과** (로컬 + 배포 사이트 양쪽).
- ✅ **GitHub Pages 배포 완료** — Settings→Pages, `main`/`(root)`, Source="Deploy from a branch".
  레포는 Public 으로 전환함(무료 플랜은 Private 레포 Pages 불가). 공개되는 건 코드뿐이고
  사용자 엑셀 데이터는 브라우저 밖으로 안 나감. 공개 전 비밀정보 스캔 완료(매칭 0).

### ⚠️ Pages 함정 — 반드시 알아둘 것
- GitHub Pages 는 기본 **Jekyll** 빌드라 **밑줄(`_`)로 시작하는 파일을 무시**한다.
  그래서 `excel_analyzer/__init__.py` 가 404 로 빠져 "엔진 준비 실패"가 났었다.
  → **레포 루트의 빈 `.nojekyll` 파일**로 Jekyll 을 꺼서 해결(이미 커밋됨). 이 파일 지우지 말 것.

주의/팁:
- `report.py` 의 그래프는 외부 CDN 불필요(순수 canvas) → 다운로드 HTML 은 오프라인 동작.
- Pyodide 만 CDN 으로 받음(코드, 데이터 아님). 첫 로딩 ~10MB.
- 폐쇄망 요구가 생기면 Pyodide/자원을 레포에 내장(현재는 CDN 으로 가기로 함).
- openpyxl 은 순수 파이썬이라 micropip 로 어떤 Pyodide 버전에서도 설치됨.
- `main` 에 푸시하면 Pages 가 자동 재배포(1~2분). 배포 검증은 curl 로
  `https://qriusquokka.github.io/Excel_Structure_Analyzer/excel_analyzer/__init__.py` 가 200 인지 보면 됨.

## 7. 실행 / 테스트 방법

```bash
pip install -r requirements.txt
python make_sample.py            # 합성 테스트 파일 sample.xlsx 생성
python analyze.py sample.xlsx    # 분석 → sample_analysis.html 생성 후 브라우저 자동 열기
python analyze.py sample.xlsx --no-open   # 브라우저 안 열기
```
빠른 검증용으로 콘솔 한글이 깨지면 `PYTHONIOENCODING=utf-8` 를 앞에 붙인다.

## 8. 코딩 규약 / 환경

- 환경: Windows, Python 3.12, openpyxl 3.1.x. 주석·UI 문자열은 한국어.
- 엔진을 키우면 CLI·웹에 동시에 반영된다(단일 소스). 기능 추가는 `excel_analyzer/` 에서.
- `report.py` 의 Python `ROLE_COLORS` 와 JS `ROLE_COLORS` 는 **항상 동일하게** 유지(색 일치).

## 9. 사용자/협업 메모

- 사용자는 개발을 직접 깊게 다루기보다 **방향성을 충분히 논의하고 계획을 확정한 뒤 개발 착수**하는 것을 선호.
  큰 변경 전에는 트레이드오프를 솔직히 설명하고 확인받을 것.
- 데이터 프라이버시에 민감 → 외부 전송이 생길 수 있는 모든 행위는 사전에 반드시 짚어줄 것.

## 10. 셀 단위 추적(v2) 구현 메모

- **엔진 데이터 모델** (`analyzer.py`):
  - `CellRef(sheet, ref, is_range, is_external, raw)` — 한 선행참조.
  - `CellNode(sheet, coord, formula, value, precedents:list[CellRef], dependents:list[str])`.
  - `WorkbookAnalysis.cells: dict["시트!셀", CellNode]`, `.edge_cells: dict["from->to", [{dst,src,formula}]]`,
    `.range_capped`(상한 초과로 하류 전개 생략한 범위 수).
  - 2-패스: ① 모든 셀의 수식/값·시트통계 기록 → ② 수식 참조를 풀어 precedents/dependents/시트엣지/드릴다운 구성.
- **리포트 JS** (`report.py` `_SCRIPT`): 시트 그래프(canvas, 엣지 hover/click 히트테스트)+`showEdge`(드릴다운),
  셀 추적기(`subgraph`→`longestLevels`→`renderFlow`+`drawFlowEdges` SVG). 데이터는 `__CELLS__`/`__EDGECELLS__`/
  `__SHEETS__` 로 주입. 대형 파일은 셀 JSON 임베드로 HTML 이 커질 수 있음(필요 시 경량화 여지).
- **검증 방식**: `sample.xlsx` 의 다단계 체인(입력→가공 단계→표C/표D→요약)으로 확인.
  레벨링 로직은 Node 시뮬레이션으로 단위 검증 가능(입력!A2=맨왼쪽 5단, 요약!A4=맨오른쪽, 3D참조 분리 등).
- **시제품**(임시, 레포 밖 scratchpad: `prototype_gen.py`/`prototype.html`)으로 표현을 먼저 합의한 뒤 본 엔진 이식.
  본 구현 완료 후 시제품은 폐기 대상(레포에 없음).
