"""Excel Structure Analyzer.

여러 시트로 구성되고 시트 간 수식으로 얽힌 .xlsx 파일의 구조를 분석해
시트 목록, 수식, 시트 간 의존 관계를 파악하고 HTML 리포트로 출력한다.
"""

from .analyzer import (
    analyze_workbook,
    WorkbookAnalysis,
    SheetInfo,
    Dependency,
    CellNode,
    CellRef,
)

__all__ = [
    "analyze_workbook",
    "WorkbookAnalysis",
    "SheetInfo",
    "Dependency",
    "CellNode",
    "CellRef",
]

__version__ = "0.1.0"
