"""수식 문자열에서 시트 참조를 추출하는 모듈.

엑셀 수식 안의 시트 참조는 다음과 같은 형태로 나타난다.

    SheetB!A1                 일반 참조
    SheetB!$A$1:$C$10         절대/범위 참조
    '시트 이름'!A1            공백/특수문자가 있으면 작은따옴표로 감쌈
    '오타''난시트'!A1         이름 안의 작은따옴표는 ''로 이스케이프
    Sheet1:Sheet3!A1          3D 참조(여러 시트에 걸친 참조)
    [1]Sheet!A1              외부 통합문서 참조
    'C:\\경로\\[file.xlsx]Sheet'!A1   외부 파일 참조

핵심 전략: 정규식으로 "...!" 앞의 시트명 후보를 뽑은 뒤,
실제 존재하는 시트 이름 목록과 대조해 오탐(함수명, 텍스트 등)을 걸러낸다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

# "<시트명후보>!" 형태를 찾는다.
#  - 그룹 quoted: 작은따옴표로 감싼 이름 (내부 '' 이스케이프 포함)
#  - 그룹 bare  : 따옴표 없는 이름 (한글/영문/숫자/밑줄/마침표 등, 연산자/구분자는 제외)
# 3D 참조(Sheet1:Sheet3!)는 bare 쪽에서 콜론을 포함해 한 덩어리로 잡은 뒤 후처리한다.
_SHEET_REF = re.compile(
    r"""
    (?:
        '(?P<quoted>(?:[^']|'')*)'      # '작은따옴표로 감싼 이름'
      | (?P<bare>[^\s'"(){},;:!+\-*/^&%<>=~]+
                 (?::[^\s'"(){},;:!+\-*/^&%<>=~]+)?)  # 따옴표 없는 이름 (옵션: :3D 끝점)
                                                       # 대괄호는 허용: [1]Sheet 같은 외부 참조 포착용
    )
    !                                    # 시트와 셀을 가르는 느낌표
    """,
    re.VERBOSE,
)


@dataclass
class SheetReference:
    """수식에서 발견된 하나의 시트 참조."""

    raw: str                 # 매칭된 원본 텍스트 (예: "'시트 A'!")
    sheet_names: list[str]   # 참조된 시트 이름들 (3D 참조면 여러 개)
    is_external: bool = False  # 외부 통합문서([..]) 참조 여부


def _unescape(name: str) -> str:
    """작은따옴표 이스케이프('')를 단일 따옴표로 되돌린다."""
    return name.replace("''", "'")


def extract_sheet_references(
    formula: str,
    known_sheets: list[str],
) -> list[SheetReference]:
    """수식에서 (실제로 존재하는) 시트 참조 목록을 추출한다.

    known_sheets 와 대소문자 무시 비교를 한다(엑셀은 시트명 대소문자 비구분).
    3D 참조 Sheet1:Sheet3 는 시트 순서상 사이에 있는 시트들로 확장한다.
    외부 통합문서 참조([..])는 is_external=True 로 표시하고 시트명은 비운다.
    """
    if not formula:
        return []

    # 대소문자 무시 매칭용 매핑: 소문자 이름 -> 실제 이름
    lower_to_actual = {s.lower(): s for s in known_sheets}
    index_of = {s: i for i, s in enumerate(known_sheets)}

    results: list[SheetReference] = []

    for m in _SHEET_REF.finditer(formula):
        quoted = m.group("quoted")
        bare = m.group("bare")

        if quoted is not None:
            candidate = _unescape(quoted)
        else:
            candidate = bare

        # 외부 통합문서 참조 판별: [숫자] 또는 [파일명] 이 이름에 포함됨
        if "[" in candidate and "]" in candidate:
            results.append(
                SheetReference(raw=m.group(0), sheet_names=[], is_external=True)
            )
            continue

        # 3D 참조 처리: "Sheet1:Sheet3"
        if ":" in candidate:
            left, right = candidate.split(":", 1)
            expanded = _expand_3d(left, right, lower_to_actual, index_of, known_sheets)
            if expanded:
                results.append(
                    SheetReference(raw=m.group(0), sheet_names=expanded)
                )
            continue

        actual = lower_to_actual.get(candidate.lower())
        if actual is not None:
            results.append(SheetReference(raw=m.group(0), sheet_names=[actual]))

    return results


def _expand_3d(left, right, lower_to_actual, index_of, known_sheets):
    a = lower_to_actual.get(left.lower())
    b = lower_to_actual.get(right.lower())
    if a is None or b is None:
        return []
    i, j = index_of[a], index_of[b]
    if i > j:
        i, j = j, i
    return known_sheets[i : j + 1]
