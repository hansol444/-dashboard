# PPT 장표 번역 자동화

## 프로젝트 목적
python-pptx와 Claude API를 사용해 PPT 슬라이드의 텍스트를 자동 번역한다.
기본 방향은 한국어 → 영어(Australian English), 역방향 옵션도 지원한다.

## 폴더 구조
```
장표_번역_자동화/
├── translate.py          # 메인 실행 스크립트
├── config.py             # 설정: API 키, 언어, 모델, 후처리 on/off
├── box_analyzer.py       # Q1: 텍스트박스 크기 분석 → 글자수 제약 계산
├── post_processor.py     # Q2: 번역 후처리 규칙
├── terminology.json      # Q3: 용어집 (ko↔en, preserve 목록)
├── CLAUDE.md             # 이 파일
├── SYSTEM_PROMPT.txt     # Claude API에 넘기는 번역용 시스템 프롬프트
├── input/                # 번역할 PPT 파일 위치
├── output/               # 번역된 PPT + 번역 리포트 (CSV/JSON)
└── requirements.txt      # 필요 패키지 목록
```

## 기술 스택
- Python 3.x
- python-pptx — PPT 읽기/쓰기
- anthropic — Claude API 호출
- 모델: `claude-sonnet-4-6`

---

## 핵심 기능 3가지

### Q1: 텍스트박스 제약 기반 번역 (`box_analyzer.py`)
- 각 텍스트박스의 너비(Emu)와 폰트 크기(pt)를 측정
- 계산 공식: `max_chars = floor(box_width_pt / (font_size_pt × 0.6))`
- API 요청 시 "이 텍스트는 영문 N자 이내로 번역하라" 제약 포함
- 번역 결과가 N자를 초과하면 output에 `⚠ 초과` 표시 + 로그 기록

### Q2: 번역 후처리 자동 수정 (`post_processor.py`)
각 규칙은 config.py에서 on/off 가능. 적용 순서:
1. `fix_billion` — 억/조 단위 오류 교정 (14B → 1.4B)
2. `fix_duplicates` — 연속 중복 제거 (JK JK → JK)
3. `fix_month_abbrev` — 월 약어 오류 교정 (1M → January)
4. `fix_currency_order` — 통화 코드 위치 교정 (100M KRW → KRW 100M)
5. `fix_australian_spelling` — 미국식 → 호주식 철자 교정

### Q3: 용어집 기반 번역 (`terminology.json`)
- `ko_to_en` / `en_to_ko` 딕셔너리로 번역 방향별 용어 관리
- `preserve` 리스트: 번역하지 않고 원문 그대로 유지 (APAC, KPI 등)
- API 호출 시 시스템 프롬프트에 관련 용어 자동 포함

---

## 번역 규칙 (SYSTEM_PROMPT.txt 참고)
- 영어 스타일: **Australian English** (colour, organisation, analyse 등)
- 전문 분야: 채용/HR (recruitment, talent acquisition 관련 용어 우선)
- 슬라이드 전체 텍스트를 문맥으로 함께 전송 (단편 번역 금지)
- 불릿 포인트·줄바꿈·특수문자 등 원문 서식 그대로 유지

---

## 실행 방법
```bash
pip install -r requirements.txt
python translate.py --input input/sample.pptx --output output/sample_en.pptx
```

## 출력물
- 번역된 PPT 파일 (`output/*.pptx`)
- 번역 리포트: 슬라이드별 원문/번역문/글자수 제약 초과 여부 기록
