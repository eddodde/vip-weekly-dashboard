# VIP 주간 실적 대시보드

LF CRM/VIP 주간회의용 실적 대시보드. 기존 주간회의 엑셀의 **`Summary` 시트 2.실적**을
매주 수기 입력하던 작업을 대체합니다. 원본 BI export(전체관점/상품관점 · 일·주·월별)를
그대로 올리면 전년비까지 자동 계산합니다.

## 구성
- **핵심 KPI** — 최신 주차 거래액·고객수·DAU·유효회원수·유입률·CR·객단가 + 전년비
- **월별 / 주차별** — 지표별 올해 vs 전년 + 전년비
- **채널별** — 직접/광고/EP/PUSH/제휴 구성·추세·전년비
- **상품별** — e-영업1~4 × 카테고리 전년비 히트맵 + 급변 TOP

## 데이터
- 시드: `data/perf_long.csv` (tidy long: `grain, perspective, year, period, period_sort, seg1, seg2, metric, value, source`)
- 매주 갱신: 앱 사이드바 **원본 데이터 업로드**에 아래 6종 xlsx를 올리면 즉시 재계산
  1. 전체관점 - 월별 실적
  2. 전체관점 - 주별 실적 (TTL)
  3. 전체관점 - 주별 실적 (채널별)
  4. 전체관점 - 일자별 실적 (당해)
  5. 전체관점 - 일자별 실적 (전년)
  6. 상품관점 - 주별 실적

### 시드 CSV 재생성 (로컬, Excel 설치 필요)
```powershell
# Downloads의 원본 6종을 tidy CSV로 변환
$dl = "$env:USERPROFILE\Downloads"
$paths = @(
  "$dl\전체관점 - 월별 실적 (기본) (1).xlsx",
  "$dl\전체관점 - 주별 실적 (기본) (1).xlsx",
  "$dl\전체관점 - 주별 실적 (기본) (2).xlsx",
  "$dl\전체관점 - 2025년 일자별 실적 (기본) (6).xlsx",
  "$dl\전체관점 - 일자별 실적 (기본) (5).xlsx",
  "$dl\상품관점 - 주별 실적(기본).xlsx"
)
powershell -ExecutionPolicy Bypass -File .\convert.ps1 -Src $paths -Out data\perf_long.csv
```

## 배포
GitHub push → Streamlit Cloud (`app.py`). 파서는 `convert.ps1`(로컬 시드)과
`app.py`의 `parse_workbook`(업로드)이 동일 로직을 공유합니다.

## 파서 로직 (BI export 구조)
- **R1** = 연도 마커(값 열 시작 위치마다), 열 방향 forward-fill → 각 값 열의 연도 확정
- **R3** = 기간 라벨(월/주차/일자), **라벨 열 수**로 관점 판별(전체=6, 상품=7)
- 전년비 = 동일 기간 라벨을 올해 vs 전년으로 매칭

## 주간 갱신 절차 (파일만 바꾸면 화면이 따라 바뀌는 범위)
1. **실적 5종**(전체관점 일자별·주별 2종·월별, 상품관점 주별) → `convert.ps1`로 각각 CSV 변환 → `python tools/merge_seed.py <CSV 폴더> --write`
2. **표(문자)** → `sms_convert.ps1 -Src <파일>` → `data/sms_long.csv`(주차 집계) + `data/sms_daily.csv`(일자·캠페인 원장)
3. **팝업 실적(Overall.xlsx)** → `popup_convert.ps1 -Src <파일>` → `data/popup.csv`
4. `git add data && git commit && git push` → Streamlit Cloud Reboot

- 캐시는 데이터 파일 크기·수정시각으로 자동 무효화된다(버전 문자열 수정 불필요).
- 유입 액션 카드(장바구니 자동화·최근 조회 리마인드)는 문자 원장 최근 7일, 전환 카드(장바구니 팝업)는 팝업 누적에서 자동 산출된다.
- 구조 진단(다각도 전년비 표)은 직전 마감주 또는 진행주에 전관행사가 낀 경우(올해·전년 어느 쪽이든)에만 표시된다.
- 수기로 남는 것: 개선 목표(이번 주 신규 액션 대상 수), `PRED_LOG`(보고한 예상·목표), `data/comment.md`, `data/events.csv`.
