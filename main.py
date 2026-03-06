#!/usr/bin/env python3
"""COA 자동 생성기 - HPLC PDF 데이터에서 COA 엑셀/PDF 생성."""

import argparse
import sys
from pathlib import Path

from src.pdf_parser import parse_pdf
from src.hplc_analyzer import analyze_pdf_data
from src.coa_writer import fill_coa, convert_to_pdf


def main():
    parser = argparse.ArgumentParser(
        description="HPLC PDF 분석 데이터로 COA 자동 생성"
    )
    parser.add_argument(
        "--template", "-t",
        required=True,
        help="COA 엑셀 템플릿 경로 (예: samples/coa_templete.xlsx)"
    )
    parser.add_argument(
        "--pdfs", "-p",
        nargs="+",
        required=True,
        help="HPLC PDF 파일 경로 (여러 개 가능)"
    )
    parser.add_argument(
        "--output", "-o",
        default="output",
        help="출력 디렉토리 (기본: output/)"
    )
    parser.add_argument(
        "--manufactured-date",
        default="",
        help="제조일 (예: 2026.01.05)"
    )
    parser.add_argument(
        "--issue-no",
        default="",
        help="발행번호 (예: COAN4010001-00)"
    )
    parser.add_argument(
        "--issue-date",
        default="",
        help="발행일 (기본: 오늘 날짜, 형식: 2026.01.06)"
    )
    parser.add_argument(
        "--appearance",
        default="",
        help="외관 (예: 'Yellowish powder', 'White powder')"
    )
    parser.add_argument(
        "--qm-manager",
        default="",
        help="QM Manager (예: 'Chang Seok-Keon ( QM Manager )')"
    )
    parser.add_argument(
        "--no-pdf",
        action="store_true",
        help="PDF 변환 건너뛰기 (Excel만 생성)"
    )

    args = parser.parse_args()

    # 출력 디렉토리 생성
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 각 PDF 파일 처리
    all_results = []
    for pdf_path in args.pdfs:
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            print(f"[경고] 파일 없음: {pdf_path}", file=sys.stderr)
            continue

        print(f"\n{'='*60}")
        print(f"PDF 파싱 중: {pdf_path.name}")
        print(f"{'='*60}")

        pdf_data = parse_pdf(str(pdf_path))

        print(f"  품목명: {pdf_data.item_name}")
        print(f"  분석자: {pdf_data.analyst_name}")
        print(f"  Lot 목록: {pdf_data.lot_numbers}")
        print(f"  HPLC runs 수: {len(pdf_data.hplc_runs)}")

        # Purity 분석
        results = analyze_pdf_data(pdf_data)

        for r in results:
            print(f"\n  [Lot {r.lot_no}]")
            print(f"    Injection별 Purity: {r.injection_purities}")
            print(f"    최종 Purity (min): {r.purity:.4f}%")
            print(f"    Injection 수: {r.num_injections}")

        all_results.extend(results)

    if not all_results:
        print("\n[오류] 처리할 HPLC 데이터가 없습니다.", file=sys.stderr)
        sys.exit(1)

    # COA 생성
    print(f"\n{'='*60}")
    print(f"COA 생성 중...")
    print(f"{'='*60}")

    for result in all_results:
        xlsx_name = f"COA_{result.item_name}_{result.lot_no}.xlsx"
        xlsx_path = output_dir / xlsx_name

        fill_coa(
            template_path=args.template,
            result=result,
            output_path=str(xlsx_path),
            manufactured_date=args.manufactured_date,
            issue_no=args.issue_no,
            issue_date=args.issue_date,
            appearance=args.appearance,
            qm_manager=args.qm_manager,
        )
        print(f"  Excel 생성: {xlsx_path}")

        # PDF 변환
        if not args.no_pdf:
            try:
                pdf_out = convert_to_pdf(str(xlsx_path), str(output_dir))
                print(f"  PDF 생성: {pdf_out}")
            except Exception as e:
                print(f"  [경고] PDF 변환 실패: {e}", file=sys.stderr)

    print(f"\n완료! 총 {len(all_results)}개 COA 생성됨.")


if __name__ == "__main__":
    main()
