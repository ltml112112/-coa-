# 분석자 한글 → 영문 이름 매핑 (추후 업데이트)
ANALYST_NAME_MAP = {
    "강병구": "강병구",  # 추후 영문명으로 교체
    "성재관": "성재관",
    "양원준": "양원준",
    "김수연": "김수연",
}

# QM Manager 목록 (추후 업데이트)
QM_MANAGERS = {
    "이진우": "Chang Seok-Keon ( QM Manager )",  # 기본값, 추후 수정
}

# 서명 이미지 경로 (추후 업데이트)
SIGNATURE_DIR = "signatures/"

# COA 템플릿 셀 위치
COA_CELLS = {
    "item": "I11",
    "lot_no": "I12",
    "manufactured_date": "I13",
    "issue_no": "I14",
    "issue_date": "I15",
    "appearance_spec": "I21",
    "appearance_result": "O21",
    "purity_spec": "I22",
    "purity_value": "X24",       # 수식이 O22에서 X24를 참조
    "analyst_name": "C39",
    "qm_manager": "M39",
}

# Purity 계산 시 포함할 최소 area%
PURITY_AREA_THRESHOLD = 1.0  # area% > 1% 인 피크만 합산
