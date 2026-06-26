"""Excel 구조 분석기 CLI.

사용법:
    python analyze.py 파일.xlsx
    python analyze.py 파일.xlsx -o 리포트.html
    python analyze.py 파일.xlsx --no-open      # 분석 후 브라우저 자동 실행 안 함
"""

from __future__ import annotations

import argparse
import os
import sys
import webbrowser

from excel_analyzer import analyze_workbook
from excel_analyzer.report import write_report


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="엑셀(.xlsx) 파일의 시트 구조와 시트 간 수식 관계를 분석해 HTML 리포트를 만듭니다.",
    )
    parser.add_argument("file", help="분석할 .xlsx (또는 .xlsm) 파일 경로")
    parser.add_argument(
        "-o", "--output",
        help="출력 HTML 경로 (기본: 입력파일명_analysis.html)",
    )
    parser.add_argument(
        "--no-open", action="store_true",
        help="분석 후 브라우저를 자동으로 열지 않음",
    )
    args = parser.parse_args(argv)

    if not os.path.isfile(args.file):
        print(f"[오류] 파일을 찾을 수 없습니다: {args.file}", file=sys.stderr)
        return 1

    try:
        analysis = analyze_workbook(args.file)
    except ValueError as e:
        print(f"[안내] {e}", file=sys.stderr)
        return 2
    except Exception as e:  # noqa: BLE001
        print(f"[오류] 분석 중 문제가 발생했습니다: {e}", file=sys.stderr)
        return 3

    output = args.output or (os.path.splitext(args.file)[0] + "_analysis.html")
    write_report(analysis, output)

    # 콘솔 요약
    print(f"분석 완료: {analysis.file_name}")
    print(f"  시트 {len(analysis.sheets)}개 · 시트 간 관계 {len(analysis.dependencies)}개")
    if analysis.input_sheets():
        print(f"  입력단 후보: {', '.join(analysis.input_sheets())}")
    if analysis.output_sheets():
        print(f"  산출단 후보: {', '.join(analysis.output_sheets())}")
    print(f"  리포트: {os.path.abspath(output)}")

    if not args.no_open:
        webbrowser.open("file://" + os.path.abspath(output))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
