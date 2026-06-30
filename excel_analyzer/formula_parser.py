"""수식 문자열에서 셀/시트 참조를 추출하는 모듈.

엑셀 수식 안의 참조는 다음과 같은 형태로 나타난다.

    A1, $A$1, A1:C10        같은 시트(시트명 없음) 셀/범위
    SheetB!A1               다른 시트 참조
    SheetB!$A$1:$C$10       절대/범위 참조
    '시트 이름'!A1          공백/특수문자가 있으면 작은따옴표로 감쌈
    '오타''난시트'!A1        이름 안의 작은따옴표는 ''로 이스케이프
    Sheet1:Sheet3!A1        3D 참조(여러 시트에 걸친 참조)
    [1]Sheet!A1             외부 통합문서 참조
    'C:\\경로\\[file.xlsx]Sheet'!A1   외부 파일 참조

핵심 전략: openpyxl 내장 ``Tokenizer`` 로 수식을 토큰화해 **셀/범위 피연산자만**
정확히 뽑는다(정규식보다 함수명·문자열·날짜 오탐이 적다). 시트명은 실제 존재하는
시트 목록과 대조해 거르고, 셀 부분은 셀/범위 형태인지 검증한다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from openpyxl.formula.tokenizer import Tokenizer


# ── 셀/범위 형태 검증용 (오탐 방지: 정의된 이름 등 제외) ──────────────────────
_CELL = re.compile(r"^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$")   # A1, $A$1
_COL = re.compile(r"^\$?[A-Za-z]{1,3}$")                  # A, $A  (전체 열)
_ROW = re.compile(r"^\$?[0-9]{1,7}$")                     # 1      (전체 행)


def _is_cell_ref(s: str) -> bool:
    """셀/범위(A1, A1:C10, A:A, 1:1) 형태인지."""
    if ":" in s:
        a, b = s.split(":", 1)
        return bool(
            (_CELL.match(a) and _CELL.match(b))
            or (_COL.match(a) and _COL.match(b))
            or (_ROW.match(a) and _ROW.match(b))
        )
    return bool(_CELL.match(s))


def _norm(cell: str) -> str:
    """셀 참조 정규화: 절대참조 기호 제거 + 열 문자 대문자화."""
    return cell.replace("$", "").upper()


@dataclass
class CellReference:
    """수식에서 발견된 하나의 셀/범위 참조."""

    sheet: str                # 참조된 시트명(같은 시트면 current_sheet). 외부참조면 "".
    ref: str                  # 정규화된 셀/범위 (예: "A1", "A1:C10")
    is_range: bool = False
    is_external: bool = False  # 외부 통합문서([..]) 참조 여부
    raw: str = ""             # 토큰 원문 (예: "'시트 A'!A1")


@dataclass
class SheetReference:
    """수식에서 발견된 시트 단위 참조(하위호환용 — 셀 참조에서 파생)."""

    raw: str
    sheet_names: list[str]
    is_external: bool = False


def _expand_3d(left, right, lower_to_actual, index_of, known_sheets):
    a = lower_to_actual.get(left.lower())
    b = lower_to_actual.get(right.lower())
    if a is None or b is None:
        return []
    i, j = index_of[a], index_of[b]
    if i > j:
        i, j = j, i
    return known_sheets[i : j + 1]


def _parse_operand(val, current_sheet, lower_to_actual, index_of, known_sheets):
    """Tokenizer 가 뽑은 RANGE 피연산자 하나를 CellReference 들로 변환."""
    # 외부 통합문서 참조: 이름에 [..] 가 포함됨
    if "[" in val and "]" in val:
        cell = val.rsplit("!", 1)[1] if "!" in val else ""
        return [CellReference(
            sheet="", ref=_norm(cell), is_range=":" in cell,
            is_external=True, raw=val,
        )]

    if "!" in val:
        sheet_part, cell = val.rsplit("!", 1)
        sheet_part = sheet_part.strip()
        if sheet_part.startswith("'") and sheet_part.endswith("'"):
            sheet_part = sheet_part[1:-1].replace("''", "'")
        if ":" in sheet_part:                      # 3D 참조 Sheet1:Sheet3
            sheets = _expand_3d(sheet_part.split(":", 1)[0], sheet_part.split(":", 1)[1],
                                lower_to_actual, index_of, known_sheets)
        else:
            actual = lower_to_actual.get(sheet_part.lower())
            sheets = [actual] if actual is not None else []
    else:                                          # 시트명 없음 = 같은 시트
        sheets, cell = [current_sheet], val

    if not _is_cell_ref(cell):                      # 정의된 이름 등은 제외
        return []
    norm = _norm(cell)
    is_rng = ":" in norm
    return [CellReference(sheet=s, ref=norm, is_range=is_rng, raw=val) for s in sheets if s]


def extract_cell_references(
    formula: str,
    current_sheet: str,
    known_sheets: list[str],
) -> list[CellReference]:
    """수식에서 (실제 존재하는 시트의) 셀/범위 참조 목록을 추출한다.

    같은 시트 참조(시트명 없는 A1 등)는 sheet=current_sheet 로 채운다.
    3D 참조는 사이 시트들로 확장, 외부 통합문서 참조는 is_external=True 로 표시한다.
    """
    if not formula:
        return []
    text = formula if formula.startswith("=") else "=" + formula
    try:
        tokens = Tokenizer(text).items
    except Exception:
        return []

    lower_to_actual = {s.lower(): s for s in known_sheets}
    index_of = {s: i for i, s in enumerate(known_sheets)}

    out: list[CellReference] = []
    for t in tokens:
        if t.type != "OPERAND" or t.subtype != "RANGE":
            continue
        out.extend(_parse_operand(t.value, current_sheet, lower_to_actual, index_of, known_sheets))
    return out


def extract_sheet_references(
    formula: str,
    known_sheets: list[str],
    current_sheet: str = "",
) -> list[SheetReference]:
    """수식에서 시트 단위 참조 목록을 추출한다(셀 참조에서 파생, 하위호환용)."""
    out: list[SheetReference] = []
    for r in extract_cell_references(formula, current_sheet, known_sheets):
        if r.is_external:
            out.append(SheetReference(raw=r.raw, sheet_names=[], is_external=True))
        elif r.sheet and r.sheet != current_sheet:
            out.append(SheetReference(raw=r.raw, sheet_names=[r.sheet]))
    return out
