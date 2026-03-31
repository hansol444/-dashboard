# PPT 장표 번역 자동화

python-pptx와 Claude API를 사용해 PPT 슬라이드의 텍스트를 자동 번역합니다.
기본 방향은 **한국어 → 영어(Australian English)**, 역방향 옵션도 지원합니다.

## 주요 기능

### 1. 텍스트박스 제약 기반 번역
- 각 텍스트박스의 너비와 폰트 크기를 측정하여 영문 최대 글자수를 계산
- 번역 결과가 제약을 초과하면 자동 재번역 시도 + 리포트 기록

### 2. 번역 후처리 자동 수정
| 규칙 | 설명 | 예시 |
|---|---|---|
| `fix_billion` | 억/조 단위 오류 교정 | 14B → 1.4B |
| `fix_duplicates` | 연속 중복 제거 | JK JK → JK |
| `fix_month_abbrev` | 월 약어 오류 교정 | 1M → January |
| `fix_currency_order` | 통화 코드 위치 교정 | 100M KRW → KRW 100M |
| `fix_australian_spelling` | 미국식 → 호주식 철자 | organization → organisation |

### 3. 용어집 기반 번역
- `terminology.json`으로 번역 방향별 용어 관리
- 번역하지 않고 보존할 약어 목록 (APAC, KPI 등)

## 실행 방법

### 사전 준비

```bash
pip install -r requirements.txt
```

환경변수 설정 (`.env.example` 참고):
```bash
export ANTHROPIC_API_KEY=sk-ant-api03-your-key-here
```

### CLI 실행

```bash
# 단일 파일 번역 (한→영)
python translate.py input/파일.pptx --to en

# 역방향 (영→한)
python translate.py input/파일.pptx --to ko

# input/ 폴더 전체 번역
python translate.py --batch --to en

# 정밀 모드 (느리지만 정확)
python translate.py input/파일.pptx --quality precise
```

### Web UI 실행

```bash
python app.py
# → http://localhost:5000
```

### 데스크탑 GUI 실행

```bash
python gui.py
```

### Windows 간편 실행

```
run.bat
```

## 폴더 구조

```
├── translate.py          # 메인 번역 스크립트
├── config.py             # 설정 (API 키, 언어, 모델, 후처리 on/off)
├── box_analyzer.py       # 텍스트박스 크기 분석 → 글자수 제약 계산
├── post_processor.py     # 번역 후처리 규칙
├── terminology.json      # 용어집 (ko↔en, preserve 목록)
├── SYSTEM_PROMPT.txt     # Claude API 번역용 시스템 프롬프트
├── app.py                # Web UI (Flask)
├── gui.py                # 데스크탑 GUI (tkinter)
├── integrations.py       # SharePoint 업로드 & Slack 알림
├── create_test_ppt.py    # 테스트용 PPT 생성 스크립트
├── templates/
│   └── index.html        # Web UI 템플릿
├── input/                # 번역할 PPT 파일
├── output/               # 번역된 PPT + 리포트
├── requirements.txt      # Python 패키지 목록
├── run.bat               # Windows 실행 배치
├── .env.example          # 환경변수 템플릿
└── .gitignore
```

## 기술 스택

- **Python 3.x**
- **python-pptx** — PPT 읽기/쓰기
- **anthropic** — Claude API 호출
- **Flask** — Web UI
- **모델**: `claude-sonnet-4-6`

## 연동 (선택)

- **SharePoint**: 번역 완료된 PPT를 자동 업로드
- **Slack**: 번역 완료 알림 전송

`.env.example`을 참고하여 환경변수를 설정하면 활성화됩니다.
