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

## 3. 현재 상태 (2026-06-26 기준)

- ✅ **분석 엔진** (`excel_analyzer/`) 완성 및 검증.
- ✅ **CLI** (`analyze.py`) 동작.
- ✅ **HTML 리포트** (`report.py`) — 인터랙티브 그래프 + 표 + 수식 상세, 디자인까지 다듬음.
  - 접철 토글 아이콘: 크기 키우고(시인성), 접힘 시 ◂(왼쪽) 방향으로 변경.
- 🟡 **웹페이지(Pyodide)** — `index.html` + `app.js` **구현 완료**. 로컬 서버에서 정적 서빙·경로는
  확인했으나, **브라우저에서 실제 분석 동작은 아직 실측 검증 못 함**(이 기기에 Claude 브라우저 확장 미연결).
  → 사용자가 브라우저로 한 번 돌려보고 확인 필요. 아래 6번 참조.

## 4. 아키텍처 & 파일 구조

```
excel_analyzer/
  __init__.py          공개 API (analyze_workbook, WorkbookAnalysis, SheetInfo, Dependency)
  formula_parser.py    수식 문자열에서 시트 참조 추출
  analyzer.py          openpyxl 로 시트/수식/의존관계/순환참조 분석 엔진
  report.py            분석 결과 → 단일 HTML 리포트 렌더링
analyze.py             CLI 진입점:  python analyze.py 파일.xlsx
make_sample.py         검증용 합성 통합문서(sample.xlsx) 생성
requirements.txt       openpyxl
README.md              사용자용 설명서
```

**핵심 동작 원리**
- 수식은 openpyxl `data_only=False` 로 읽어 **수식 원문**을 얻는다.
- `formula_parser.py` 가 정규식으로 `시트명!셀` 참조를 뽑되, **실제 존재하는 시트명 목록과 대조**해
  오탐(함수명·텍스트)을 거른다. 처리하는 케이스: 한글 시트명, 공백 포함 `'시트 이름'!`,
  따옴표 이스케이프 `''`, 3D 참조 `Sheet1:Sheet3!`(사이 시트로 확장), 외부 참조 `[1]Sheet!`(별도 집계).
- **화살표 방향 = 데이터 흐름**: 셀이 `SheetB!A1` 을 참조하면 데이터는 SheetB → 현재시트로 흐른다.
  그래서 의존 엣지는 `from=참조된 시트(출처)`, `to=참조하는 시트(소비)` 로 저장한다.
- `report.py` 의 시트 간 관계도는 **외부 라이브러리 없는 순수 Canvas + JS**로 그린다.
  → 다운로드한 HTML 을 인터넷 없이 열어도 드래그·화살표가 동작한다.

## 5. 결정 로그 (왜 이렇게 했는가 — 번복 금지)

1. **출력 = HTML 리포트** (Mermaid 아님. Mermaid 는 폐기됨).
2. **분석 범위 = .xlsx/.xlsm 의 시트 단위 관계.**
   구버전 `.xls` 는 안내 메시지("xlsx 로 다시 저장 후 사용")와 함께 거부.
   셀 단위 추적은 향후 확장으로 보류(현재 `Dependency.links` 에 (셀, 수식) 이미 저장 — 확장 발판 있음).
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
   - 큰 구획 3개(관계도/시트목록/수식상세): 각각 접철식, **기본 펼침**
   - 탭 UI 는 쓰지 않음(대신 접철식).

## 6. Pyodide 웹페이지 — 구현 완료(브라우저 실측만 남음)

목표: GitHub Pages 에서 "누구나 접속 → 파일 첨부 → 브라우저 안 분석 → 결과 보기·다운로드".

구현됨(레포 루트):
```
index.html   드래그&드롭/선택 UI, 프라이버시 안내, 엔진 상태 배너, 결과 iframe 미리보기,
             HTML 다운로드 버튼, "다른 파일 분석". 색/톤은 리포트와 통일(엑셀 그린).
app.js       Pyodide(v0.26.4, CDN) 로드 → micropip 로 openpyxl 설치 →
             ENGINE_FILES(excel_analyzer/*.py)를 fetch 해 가상 FS 에 기록 →
             업로드 파일 바이트를 /home/pyodide/_uploads 에 기록(업로드 없음) →
             _run_analysis(path) = analyze_workbook + render_html → HTML 문자열 →
             iframe.srcdoc 미리보기 + Blob 다운로드. 분석 후 업로드 파일은 FS 에서 unlink.
```

### 로컬 테스트 방법(중요: file:// 는 fetch 가 막혀 안 됨 → 반드시 서버로)
```bash
python -m http.server 8765        # 레포 루트에서
# 브라우저로 http://localhost:8765 열기 → 준비 완료 뜨면 sample.xlsx 끌어다 놓기
```

### 남은 일 / 결정 필요
- ⬜ **브라우저 실측**: 위 방법으로 sample.xlsx 분석 → 그래프·표·수식·다운로드 동작 확인.
  (이 기기엔 Claude 브라우저 확장이 미연결이라 코드로는 못 돌려봄.)
- ⬜ **GitHub Pages 배포 결정**: Settings→Pages 에서 켜야 함(사용자 직접).
  ⚠️ **레포가 Private 인데 무료 플랜이면 Pages 사이트는 Public 으로 노출됨** — 단, 노출되는 건
  코드(index/app/엔진 .py)일 뿐 **사용자 엑셀 데이터는 어차피 브라우저 밖으로 안 나감**.
  코드 공개가 싫으면: 배포 보류 / Pages 비공개(Pro) / 로컬 서버로만 사용 중 택1. 사용자와 상의.

주의/팁:
- `report.py` 의 그래프는 외부 CDN 불필요(순수 canvas) → 다운로드 HTML 은 오프라인 동작.
- Pyodide 만 CDN 으로 받음(코드, 데이터 아님). 첫 로딩 ~10MB.
- 폐쇄망 요구가 생기면 Pyodide/자원을 레포에 내장(현재는 CDN 으로 가기로 함).
- openpyxl 은 순수 파이썬이라 micropip 로 어떤 Pyodide 버전에서도 설치됨.

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
