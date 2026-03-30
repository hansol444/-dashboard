# Topic: Macro Analysis 업데이트 (KOSIS 데이터)

## 개요
KOSIS에서 최신 고용 데이터를 다운로드해 Macro Analysis 엑셀 파일의 10개 시트를 업데이트하는 정기 업무.
`update_macro.py` 스크립트로 자동화되어 있음.

## 실행 방법
```bash
python update_macro.py
```

## 사전 준비
- KOSIS 폴더에 `산업_규모별_고용_*.xlsx` 파일 다운로드해서 넣기
- SharePoint 경로 자동 감지됨

## 업데이트 대상 시트 (10개)
빈일자리, 채용, 근로자, 입직자 × 상용/임시일용 조합

## 주기
- 월별 데이터: 매월 업데이트
- 분기별 데이터: 분기 종료 후 업데이트

## KOSIS 다운로드 경로
- KOSIS (kosis.kr) → 국가통계포털 → 산업별 고용 관련 통계
- 파일명: `산업_규모별_고용_YYYYMM.xlsx`

## 담당 연락처
- 지표 관련 문의: 창준님
