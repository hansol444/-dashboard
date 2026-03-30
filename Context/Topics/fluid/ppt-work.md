# Topic: 장표 제작 및 번역

## 개요
전략추진실의 PPT 장표를 제작하거나, 기존 한국어 PPT를 영문으로 번역하는 업무.

---

## A. 장표 제작

### 도구
- **Claude** (스토리라인 구성, 텍스트 작성)
- **Genspark** (슬라이드 자동 생성)

### 프로세스
1. 창준님으로부터 목적/청중/분량 전달받기
2. Claude로 스토리라인 초안 작성
3. Genspark 또는 PowerPoint에서 슬라이드 구성
4. 데이터 삽입 및 디자인 적용
5. 창준님 검토 후 수정

### 폰트 규격
- 제목: Pretendard Bold 24pt
- 본문: Pretendard Regular 12-14pt
- 강조: Pretendard SemiBold

---

## B. 장표 번역 (한→영)

### 도구
```bash
cd ppt-translate
python translate.py --input input/파일.pptx --output output/파일_en.pptx
```

### 특징
- 텍스트박스 크기 제약 자동 계산 (영문이 길어져도 박스 넘침 방지)
- `terminology.json` 용어집 적용 (직종명, 서비스명 등 고유명사 통일)
- **Australian English** 스타일 (Jobkorea Australia 방향성)
- Claude API 기반 번역

### 번역 주의사항
- 용어집에 없는 신조어/고유명사는 번역 후 수동 검수 필요
- 그래프/이미지 내 텍스트는 수동 처리 필요

### 용어집 업데이트
- `ppt-translate/terminology.json` 직접 편집
- 새로운 고유명사 발견 시 추가

---

## 담당 연락처
- 창준님 — 제작 방향 및 최종 검토
