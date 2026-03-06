"""PDF에서 HPLC Area Percent Report 데이터를 추출하는 모듈."""

import re
import subprocess
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Peak:
    """HPLC 피크 데이터."""
    peak_num: int
    ret_time: float
    peak_type: str
    width: float
    area: float
    height: float
    area_pct: float


@dataclass
class HPLCRun:
    """하나의 HPLC injection 데이터."""
    seq_line: int
    inj_num: int
    sample_name: str
    sample_info: str  # Lot 번호 (ANA번호)
    peaks: list = field(default_factory=list)
    injection_date: str = ""


@dataclass
class PDFData:
    """PDF에서 추출한 전체 데이터."""
    filename: str
    item_name: str = ""
    analyst_name: str = ""
    hplc_runs: list = field(default_factory=list)
    lot_numbers: list = field(default_factory=list)


def extract_text_from_pdf(pdf_path: str) -> str:
    """pdftotext를 사용하여 PDF에서 텍스트 추출."""
    result = subprocess.run(
        ["pdftotext", "-layout", pdf_path, "-"],
        capture_output=True, text=True
    )
    return result.stdout


def extract_text_pages(pdf_path: str) -> list:
    """PDF에서 페이지별 텍스트 추출."""
    text = extract_text_from_pdf(pdf_path)
    pages = text.split("\x0c")
    return [p for p in pages if p.strip()]


def parse_analyst_from_filename(filename: str) -> str:
    """파일명에서 분석자 이름 추출.
    예: LTC-PH122_ANA1410_PG1999_HPLC_강병구.pdf → 강병구
    """
    stem = Path(filename).stem
    # 파일명 마지막 부분에서 한글 이름 추출
    match = re.search(r"[_]([가-힣]{2,4})$", stem)
    if match:
        return match.group(1)
    return ""


def parse_analyst_from_cover(text: str) -> str:
    """표지(출하품 분석 의뢰/결과서)에서 검사원 이름 추출."""
    match = re.search(r"검사원\s+([가-힣]{2,4})", text)
    if match:
        return match.group(1)
    return ""


def parse_cover_page(text: str) -> dict:
    """표지 페이지에서 메타데이터 추출."""
    info = {}

    # 품명
    match = re.search(r"품명\s+([A-Za-z0-9\-]+)", text)
    if match:
        info["item_name"] = match.group(1)

    # 검사원
    match = re.search(r"검사원\s+([가-힣]{2,4})", text)
    if match:
        info["analyst"] = match.group(1)

    # Batch No. 및 결과들
    # "1\s+ANA0101\s+99.991" 같은 패턴
    lots = []
    lot_pattern = re.compile(
        r"(\d+)\s+(ANA\d+)\s+([\d.]+)"
    )
    for m in lot_pattern.finditer(text):
        lots.append({
            "seq": int(m.group(1)),
            "lot_no": m.group(2),
            "hplc_purity": float(m.group(3)),
        })
    if lots:
        info["lots"] = lots

    return info


def parse_hplc_report(page_text: str) -> HPLCRun | None:
    """한 페이지에서 HPLC Area Percent Report 파싱.

    Returns None if no Area Percent Report found.
    """
    if "Area Percent Report" not in page_text:
        return None

    run = HPLCRun(seq_line=0, inj_num=0, sample_name="", sample_info="")

    # Seq. Line
    match = re.search(r"Seq\.\s*Line\s*:\s*(\d+)", page_text)
    if match:
        run.seq_line = int(match.group(1))

    # Inj 번호
    match = re.search(r"Inj\s*:\s*(\d+)", page_text)
    if match:
        run.inj_num = int(match.group(1))

    # Sample Name
    match = re.search(r"Sample Name[:\s]+([^\n]+)", page_text)
    if match:
        run.sample_name = match.group(1).strip()

    # Sample Info (= Lot 번호)
    match = re.search(r"Sample Info\s*:\s*([^\n]+)", page_text)
    if match:
        run.sample_info = match.group(1).strip()

    # Description (LCMS 리포트에서 Lot 번호)
    if not run.sample_info:
        match = re.search(r"Description\s*:\s*(ANA\d+)", page_text)
        if match:
            run.sample_info = match.group(1).strip()

    # Injection Date
    match = re.search(r"Injection [Dd]ate\s*:\s*([^\n]+)", page_text)
    if match:
        run.injection_date = match.group(1).strip()

    # Peak 테이블 파싱
    # Area Percent Report 이후의 피크 라인 추출
    peaks = parse_peak_table(page_text)
    run.peaks = peaks

    return run


def parse_peak_table(text: str) -> list:
    """Area Percent Report에서 피크 테이블 파싱."""
    peaks = []

    # "Signal 1:" 이후 부분 찾기
    signal_match = re.search(r"Signal\s+1\s*:.*?Ref\s*=\s*off", text)
    if not signal_match:
        return peaks

    after_signal = text[signal_match.end():]

    # 구분선(----) 이후의 데이터 행 찾기
    separator_match = re.search(r"-{4}\|.*\|", after_signal)
    if not separator_match:
        return peaks

    data_section = after_signal[separator_match.end():]

    # Totals 또는 === 이전까지의 데이터
    end_match = re.search(r"(Totals\s*:|={5,}|\*{3})", data_section)
    if end_match:
        data_section = data_section[:end_match.start()]

    # 피크 데이터 파싱
    # 형식: peak_num rettime type width area height area%
    # 여러 줄에 걸칠 수 있음 (pdftotext layout 특성)
    # 숫자들을 하나의 긴 문자열로 결합 후 파싱
    lines = data_section.strip().split("\n")
    combined = " ".join(l.strip() for l in lines if l.strip())

    # 과학 표기법 포함 숫자 패턴
    num = r"[\d.]+(?:e[+-]?\d+)?"
    peak_pattern = re.compile(
        rf"(\d+)\s+"           # peak_num
        rf"({num})\s+"         # ret_time
        rf"([A-Z]{{2}})\s+"   # type (BB, MM, MF, FM 등)
        rf"({num})\s+"         # width
        rf"({num})\s+"         # area
        rf"({num})\s+"         # height
        rf"({num})"            # area%
    )

    for m in peak_pattern.finditer(combined):
        peak = Peak(
            peak_num=int(m.group(1)),
            ret_time=float(m.group(2)),
            peak_type=m.group(3),
            width=float(m.group(4)),
            area=float(m.group(5)),
            height=float(m.group(6)),
            area_pct=float(m.group(7)),
        )
        peaks.append(peak)

    return peaks


def parse_pdf(pdf_path: str) -> PDFData:
    """PDF 파일 전체를 파싱하여 HPLC 데이터 추출."""
    pdf_path = str(pdf_path)
    filename = Path(pdf_path).name
    data = PDFData(filename=filename)

    # 분석자 이름: 파일명에서 먼저 시도
    data.analyst_name = parse_analyst_from_filename(filename)

    # 페이지별 텍스트 추출
    pages = extract_text_pages(pdf_path)

    for page_text in pages:
        # 표지 페이지 파싱
        if "출하품 분석 의뢰" in page_text or "분석 의뢰/결과서" in page_text:
            cover = parse_cover_page(page_text)
            if "item_name" in cover:
                data.item_name = cover["item_name"]
            if "analyst" in cover and not data.analyst_name:
                data.analyst_name = cover["analyst"]
            if "lots" in cover:
                data.lot_numbers = [l["lot_no"] for l in cover["lots"]]

        # HPLC 리포트 파싱
        run = parse_hplc_report(page_text)
        if run and run.peaks:
            # Item 이름 추출
            if run.sample_name and not data.item_name:
                # "LTC-PH122_0114" → "LTC-PH122"
                match = re.match(r"([A-Za-z0-9\-]+?)_\d+", run.sample_name)
                if match:
                    data.item_name = match.group(1)

            # Lot 번호 수집
            if run.sample_info and run.sample_info not in data.lot_numbers:
                data.lot_numbers.append(run.sample_info)

            data.hplc_runs.append(run)

    return data
