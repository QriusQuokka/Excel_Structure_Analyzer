"""엑셀 통합문서 구조 분석 엔진.

openpyxl 로 .xlsx 를 열어 시트별 수식을 추출하고, 수식 안의 시트 참조를
모아 시트 간 의존 관계를 만든다.

화살표 방향(데이터 흐름):
    셀이 SheetB!A1 을 참조하면 데이터는 SheetB -> (현재 시트) 로 흐른다.
    따라서 의존 엣지는 (참조된 시트) -> (참조하는 시트) 방향으로 저장한다.
    "입력 시트 -> 가공 시트 -> 산출 시트" 흐름이 자연스럽게 보인다.
"""

from __future__ import annotations

import datetime as _dt
import os
from collections import defaultdict
from dataclasses import dataclass, field

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

from .formula_parser import extract_sheet_references


@dataclass
class SheetInfo:
    """한 시트의 요약 정보."""

    name: str
    index: int
    state: str                    # visible / hidden / veryHidden
    dimensions: str = ""          # 사용 범위 (예: "A1:AM81")
    n_rows: int = 0               # 사용 범위의 행 수 (span)
    n_cols: int = 0               # 사용 범위의 열 수 (span)
    max_row: int = 0              # 절대 최대 행 (A1 기준)
    max_col: int = 0              # 절대 최대 열 (A1 기준)
    formula_count: int = 0        # 수식 셀 수
    number_count: int = 0         # 숫자/날짜 셀 수
    text_count: int = 0           # 텍스트 셀 수
    # 이 시트의 모든 수식: (셀주소, 수식원문)
    formulas: list[tuple[str, str]] = field(default_factory=list)
    # 이 시트가 참조하는 다른 시트 이름들
    references: set[str] = field(default_factory=set)

    @property
    def grid_area(self) -> int:
        return self.max_row * self.max_col

    @property
    def formula_ratio(self) -> float:
        """수식 비율 = 수식 셀 수 / (절대 최대행 × 절대 최대열)."""
        area = self.grid_area
        return (self.formula_count / area) if area else 0.0


@dataclass
class Dependency:
    """from_sheet -> to_sheet 데이터 흐름 (to_sheet 가 from_sheet 를 참조)."""

    from_sheet: str               # 데이터 출처 (참조 대상 시트)
    to_sheet: str                 # 데이터 소비 (수식이 있는 시트)
    links: list[tuple[str, str]] = field(default_factory=list)

    @property
    def weight(self) -> int:
        return len(self.links)


@dataclass
class WorkbookAnalysis:
    """통합문서 전체 분석 결과."""

    file_path: str
    sheets: list[SheetInfo] = field(default_factory=list)
    dependencies: list[Dependency] = field(default_factory=list)
    external_refs: int = 0        # 외부 통합문서 참조 개수
    self_refs: int = 0            # 자기 시트 내 참조 개수
    warnings: list[str] = field(default_factory=list)

    @property
    def file_name(self) -> str:
        return os.path.basename(self.file_path)

    @property
    def total_formulas(self) -> int:
        return sum(s.formula_count for s in self.sheets)

    def input_sheets(self) -> list[str]:
        """다른 시트를 참조하지 않는(=의존 대상이 없는) 시트 = 입력단 후보."""
        consumers = {d.to_sheet for d in self.dependencies}
        return [s.name for s in self.sheets if s.name not in consumers]

    def output_sheets(self) -> list[str]:
        """다른 시트가 참조하지 않는 시트 = 최종 산출 후보."""
        sources = {d.from_sheet for d in self.dependencies}
        return [s.name for s in self.sheets if s.name not in sources]

    def find_cycle(self) -> list[str]:
        """시트 간 순환 참조를 하나 찾아 경로로 반환한다. 없으면 빈 리스트."""
        adj: dict[str, list[str]] = defaultdict(list)
        for d in self.dependencies:
            adj[d.from_sheet].append(d.to_sheet)

        WHITE, GRAY, BLACK = 0, 1, 2
        color = {s.name: WHITE for s in self.sheets}
        stack: list[str] = []

        def dfs(u: str) -> list[str]:
            color[u] = GRAY
            stack.append(u)
            for v in adj.get(u, []):
                if color.get(v) == GRAY:
                    # v 부터 현재까지가 순환 경로
                    return stack[stack.index(v):] + [v]
                if color.get(v) == WHITE:
                    found = dfs(v)
                    if found:
                        return found
            color[u] = BLACK
            stack.pop()
            return []

        for s in self.sheets:
            if color[s.name] == WHITE:
                found = dfs(s.name)
                if found:
                    return found
        return []

    def has_cycle(self) -> bool:
        return bool(self.find_cycle())


def _formula_text(value) -> str | None:
    """셀 값에서 수식 원문을 꺼낸다. 수식이 아니면 None."""
    if isinstance(value, ArrayFormula):
        return value.text
    if isinstance(value, str) and value.startswith("="):
        return value
    return None


def analyze_workbook(file_path: str) -> WorkbookAnalysis:
    """.xlsx 파일을 분석해 WorkbookAnalysis 를 반환한다.

    .xls(구버전)는 openpyxl 로 열 수 없으므로 명확한 안내 메시지와 함께
    예외를 던진다.
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".xls":
        raise ValueError(
            "구버전 .xls 파일은 직접 분석할 수 없습니다.\n"
            "엑셀에서 파일을 열고 [다른 이름으로 저장] → 형식을 "
            "'Excel 통합 문서(*.xlsx)'로 선택해 저장한 뒤, 그 .xlsx 파일을 "
            "다시 분석해 주세요."
        )
    if ext not in (".xlsx", ".xlsm"):
        raise ValueError(
            f"지원하지 않는 형식입니다: {ext or '(확장자 없음)'}\n"
            ".xlsx 또는 .xlsm 파일을 사용해 주세요."
        )

    analysis = WorkbookAnalysis(file_path=file_path)

    # data_only=False: 계산값이 아니라 수식 원문을 읽는다.
    wb = load_workbook(filename=file_path, data_only=False, read_only=False)
    known_sheets = wb.sheetnames

    edge_links: dict[tuple[str, str], list[tuple[str, str]]] = defaultdict(list)

    for idx, ws in enumerate(wb.worksheets):
        min_row, max_row = ws.min_row or 1, ws.max_row or 1
        min_col, max_col = ws.min_column or 1, ws.max_column or 1
        info = SheetInfo(
            name=ws.title,
            index=idx,
            state=ws.sheet_state,
            dimensions=ws.dimensions,
            n_rows=max_row - min_row + 1,
            n_cols=max_col - min_col + 1,
            max_row=max_row,
            max_col=max_col,
        )

        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if value is None:
                    continue

                formula = _formula_text(value)
                if formula is not None:
                    info.formula_count += 1
                    info.formulas.append((cell.coordinate, formula))

                    refs = extract_sheet_references(formula, known_sheets)
                    for ref in refs:
                        if ref.is_external:
                            analysis.external_refs += 1
                            continue
                        for target in ref.sheet_names:
                            if target == ws.title:
                                analysis.self_refs += 1
                                continue
                            info.references.add(target)
                            edge_links[(target, ws.title)].append(
                                (cell.coordinate, formula)
                            )
                elif isinstance(value, str):
                    info.text_count += 1
                else:
                    # 숫자, 날짜/시간, bool 등은 숫자 셀로 묶는다.
                    info.number_count += 1

        analysis.sheets.append(info)

    wb.close()

    for (frm, to), links in edge_links.items():
        analysis.dependencies.append(
            Dependency(from_sheet=frm, to_sheet=to, links=links)
        )

    analysis.dependencies.sort(key=lambda d: (-d.weight, d.from_sheet, d.to_sheet))

    if not analysis.dependencies:
        analysis.warnings.append(
            "시트 간 참조 수식을 찾지 못했습니다. 시트들이 서로 독립적이거나, "
            "참조가 외부 파일/이름 정의를 통한 것일 수 있습니다."
        )

    return analysis
