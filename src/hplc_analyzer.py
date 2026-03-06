"""HPLC 데이터 분석: 확대 데이터 제거 + Purity 계산."""

from dataclasses import dataclass
from .pdf_parser import HPLCRun, PDFData, Peak
from .config import PURITY_AREA_THRESHOLD


@dataclass
class PurityResult:
    """Purity 계산 결과."""
    lot_no: str
    item_name: str
    purity: float          # min(injection별 합산 purity)
    num_injections: int
    injection_purities: list  # 각 injection별 purity 값
    analyst_name: str = ""


def _peaks_fingerprint(peaks: list) -> tuple:
    """피크 리스트의 fingerprint (중복 판별용).

    동일 데이터의 원본/확대는 Area Percent Report 값이 완전히 동일하므로
    (peak_num, ret_time, area_pct) 튜플로 비교.
    """
    return tuple(
        (p.peak_num, round(p.ret_time, 3), round(p.area_pct, 6))
        for p in sorted(peaks, key=lambda x: x.peak_num)
    )


def deduplicate_runs(runs: list) -> list:
    """동일 injection의 확대(zoomed) 데이터 제거.

    규칙:
    1. 같은 seq_line + inj_num + 동일한 peak fingerprint → 중복 → 첫 번째만 유지
    2. Ratio 데이터(IMP 피크 제거됨): 피크 수가 적은 것은 ratio data일 수 있음
       → 같은 lot의 같은 injection에서 피크 수가 더 많은 것이 원본
    """
    seen = {}
    unique_runs = []

    for run in runs:
        fp = _peaks_fingerprint(run.peaks)
        key = (run.sample_info, fp)

        if key not in seen:
            seen[key] = run
            unique_runs.append(run)

    return unique_runs


def identify_ratio_data(runs: list) -> tuple:
    """원본 데이터와 ratio 데이터 분리.

    Ratio 데이터 = IMP 피크가 제거되어 피크 수가 적은 데이터.
    같은 lot + 같은 seq_line에서 피크 수가 다른 경우,
    피크 수가 많은 것 = 원본, 적은 것 = ratio.

    Returns: (original_runs, ratio_runs)
    """
    # lot + seq_line 기준으로 그룹핑
    groups = {}
    for run in runs:
        key = (run.sample_info, run.seq_line)
        if key not in groups:
            groups[key] = []
        groups[key].append(run)

    original = []
    ratio = []

    for key, group in groups.items():
        if len(group) == 1:
            original.append(group[0])
        else:
            # 피크 수가 가장 많은 것이 원본
            group.sort(key=lambda r: len(r.peaks), reverse=True)
            original.append(group[0])
            ratio.extend(group[1:])

    return original, ratio


def calculate_purity_for_run(run: HPLCRun) -> float:
    """단일 injection의 purity 계산.

    area% > PURITY_AREA_THRESHOLD (1%) 인 피크들의 area% 합산.
    """
    included = [p for p in run.peaks if p.area_pct > PURITY_AREA_THRESHOLD]

    if not included:
        # threshold 이상인 피크가 없으면 메인 피크(최대 area%) 사용
        if run.peaks:
            return max(p.area_pct for p in run.peaks)
        return 0.0

    return sum(p.area_pct for p in included)


def calculate_purity(runs: list, lot_no: str) -> PurityResult | None:
    """특정 Lot의 HPLC 데이터로 purity 계산.

    1. 해당 lot의 run들만 필터
    2. 확대 데이터 제거
    3. Ratio 데이터 분리
    4. 각 injection별 purity 계산
    5. min 값 반환
    """
    # 해당 lot의 run만 필터
    lot_runs = [r for r in runs if r.sample_info == lot_no]
    if not lot_runs:
        return None

    # 확대 제거
    deduped = deduplicate_runs(lot_runs)

    # Ratio 분리 → 원본만 사용
    originals, _ = identify_ratio_data(deduped)

    if not originals:
        return None

    # 각 injection별 purity
    injection_purities = []
    for run in originals:
        purity = calculate_purity_for_run(run)
        injection_purities.append(purity)

    # min 값
    min_purity = min(injection_purities)

    # Item 이름 추출
    item_name = ""
    for run in originals:
        if run.sample_name:
            import re
            match = re.match(r"([A-Za-z0-9\-]+?)_\d+", run.sample_name)
            if match:
                item_name = match.group(1)
                break

    return PurityResult(
        lot_no=lot_no,
        item_name=item_name,
        purity=min_purity,
        num_injections=len(originals),
        injection_purities=injection_purities,
    )


def analyze_pdf_data(pdf_data: PDFData) -> list:
    """PDF 데이터에서 Lot별 purity 결과 계산.

    Returns: list of PurityResult
    """
    results = []

    for lot_no in pdf_data.lot_numbers:
        result = calculate_purity(pdf_data.hplc_runs, lot_no)
        if result:
            result.item_name = result.item_name or pdf_data.item_name
            result.analyst_name = pdf_data.analyst_name
            results.append(result)

    return results
