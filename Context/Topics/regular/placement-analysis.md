# Topic: Placement Survey 분석 (RMS)

## 개요
Placement Survey 원데이터를 받아 RMS(Recruitment Market Share) 분석을 수행하는 정기 업무.
Cubicle 방식의 분류표 매칭을 사용하며, 최종 결과물은 PPT로 생성됨.

## 실행 스크립트
```bash
python run_jk.py      # 분류표 매칭
python calc_rms.py    # RMS 계산
python gen_ppt.py     # PPT 생성
```

## 단계별 작업

### 1. Raw 데이터 로드
- 엠브레인으로부터 Raw 데이터 수령
- 파일 형식: xlsx

### 2. 분류표 매칭 (run_jk.py)
- Cubicle 방식: 직종/업종 분류표와 응답 매칭
- 미분류 항목 별도 리포트 생성
- 미분류 비율 확인 후 필요 시 수동 매핑 추가

### 3. RMS 계산 (calc_rms.py)
- 매칭된 데이터 기반 채용 플랫폼별 점유율 계산
- 직종별, 기업 규모별, 경력별 세분화

### 4. PPT 생성 (gen_ppt.py)
- 분석 결과를 Placement Survey 보고서 형식으로 자동 생성

## 주의사항
- 미분류 비율이 높으면 분류표 업데이트 필요 → 창준님과 확인
- Raw 데이터 파일명/컬럼 구조 변경 시 스크립트 수정 필요

## 담당 연락처
- 엠브레인 문주원님 — Raw 데이터 수령
- 창준님 — 분석 방향, 결과 검토
