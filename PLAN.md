# COA 자동 생성기 구현 계획

## 목표
PDF 분석 리포트에서 HPLC 데이터를 추출하여 COA(Certificate of Analysis) 엑셀 템플릿에 자동 기입하는 Python 프로그램

## 입력
1. `coa_templete.xlsx` - COA 엑셀 템플릿
2. HPLC PDF 파일들 (Agilent ChemStation 출력물)
   - 단일 Lot 파일: `LTC-PH122_ANA1410 (PG1999)_HPLC_강병구.pdf`
   - 복합 리포트 파일: `AEH0-G28_ANA_0101_0102_0103_LCMS_HPLC_DSC_TGA_0106_report.pdf`

## 출력
- Lot별 개별 COA 엑셀 파일 (예: `COA_LTC-PH122_ANA1410.xlsx`)
- COA 기재 항목: Purity (by HPLC) 결과값 + 분석자 이름

---

## 구현 단계

### Step 1: PDF 텍스트 파싱 모듈 (`pdf_parser.py`)
- `pdftotext` 또는 `PyMuPDF(fitz)`를 사용하여 PDF에서 텍스트 추출
- HPLC "Area Percent Report" 섹션 탐지 및 파싱
- 각 리포트에서 추출할 정보:
  - `Seq. Line` 번호 (Injection 구분)
  - `Inj` 번호
  - Peak 테이블: RetTime, Area%, 각 피크별 데이터
  - `Sample Info` (Lot 번호 = ANA번호)
  - `Sample Name` (품목명 포함)

### Step 2: 확대(Zoomed) 데이터 제거 로직
- **같은 Injection의 데이터가 2번 나오는 경우** (원본 + 확대):
  - 동일 Seq. Line + 동일 Inj 번호 → 동일한 Area Percent 데이터
  - 중복 제거: 같은 peak 데이터셋이면 첫 번째만 사용
- **IMP Peak 제거된 Ratio 데이터**:
  - 피크 수가 적고, 주요 피크만 남아있는 경우 식별
  - 원본 데이터(피크 수가 더 많은 것)만 사용

### Step 3: Purity 계산 로직
- **단일 성분** (예: LT-EB108, LTC-PH122):
  - 메인 피크 area% = Purity (보통 99.9%+)
  - 2회 injection 중 **min 값** 사용
- **다중 성분** (예: AEH0-G28 - P Type + N Type 혼합):
  - area% > 1% 인 모든 피크의 area%를 합산 = Purity
  - 2회 injection 중 합산값의 **min 값** 사용
- 소수점 처리: 기존 COA 포맷 따름 (0.000% 형식)

### Step 4: 메타데이터 추출
- **품목명 (Item)**: Sample Name에서 추출 (예: `LTC-PH122_0114` → `LTC-PH122`)
- **Lot 번호**: Sample Info 또는 Description 필드 (예: `ANA1410`, `ANA0101`)
- **분석자 이름**:
  - PDF 파일명에서 추출 (예: `_강병구.pdf` → 강병구)
  - 또는 복합 리포트의 표지 "검사원" 필드에서 추출
  - 한글 → 영문 이름 매핑 (사용자 제공 예정, 설정 파일로 관리)

### Step 5: COA 엑셀 생성 (`coa_writer.py`)
- `openpyxl`로 템플릿 로드
- 기입 위치:
  - `I11`: Item (품목명)
  - `I12`: Lot No.
  - `O22`: Purity 결과값 (수식 유지 또는 직접값 기입)
  - `C39`: 분석자 영문 이름 (좌측 하단)
- 서식/스타일/병합셀 보존
- Lot별 별도 파일 저장

### Step 6: 메인 실행 스크립트 (`main.py`)
- CLI 인터페이스:
  ```
  python main.py --template coa_templete.xlsx --pdfs *.pdf --output ./output/
  ```
- 처리 흐름:
  1. 템플릿 로드
  2. 각 PDF 파싱 → HPLC 데이터 추출
  3. Lot별 데이터 그룹핑
  4. Purity 계산
  5. COA 생성 및 저장

---

## 파일 구조
```
-coa-/
├── samples/              # 입력 파일들
├── src/
│   ├── pdf_parser.py     # PDF 텍스트 파싱 + HPLC 데이터 추출
│   ├── hplc_analyzer.py  # 확대 제거 + Purity 계산
│   ├── coa_writer.py     # COA 엑셀 생성
│   └── config.py         # 이름 매핑 등 설정
├── output/               # 생성된 COA 파일
├── main.py               # 메인 실행 스크립트
└── requirements.txt      # openpyxl, PyMuPDF
```

## 기술 스택
- Python 3
- `PyMuPDF (fitz)` 또는 `pdftotext`: PDF 텍스트 추출
- `openpyxl`: Excel 읽기/쓰기
- `re` (정규표현식): HPLC 데이터 파싱

## 미결정 사항
- [ ] 인원별 한글→영문 이름 매핑 (사용자 제공 예정)
- [ ] Appearance 값 처리 방법 (현재 수식 `=I21`로 Specification 값 복사)
- [ ] Issue No., Issue Date, Manufactured Date 등 자동생성 여부
