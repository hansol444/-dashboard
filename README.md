# Worxphere 전자결재 자동화

Playwright 기반 브라우저 자동화 스크립트.  
로그인 → 폼 자동입력 → 결재 상신까지 한 번에 실행됩니다.

---

## 설치 (최초 1회)

```bash
# 1. 이 폴더에서 의존성 설치
npm install

# 2. Playwright용 브라우저 설치
npm run install-browser

# 3. 인증 정보 설정
cp .env.example .env
# .env 파일을 열어 WORXPHERE_ID / WORXPHERE_PW 입력
```

---

## 사용법

```bash
node fill.js <문서유형> [옵션]
```

### 문서유형

| 커맨드 | 문서 |
|---|---|
| `ats-b` | ATS 인터뷰 기프티콘 — 예산품의 |
| `ats-p` | ATS 인터뷰 기프티콘 — 구매품의 |
| `survey-b` | Placement Survey — 예산품의 |
| `survey-p` | Placement Survey — 구매품의 |
| `free-b` | 산학협력 프리랜서 6주 — 예산품의 |
| `free-p` | 산학협력 프리랜서 6주 — 구매품의 |

### 예시

```bash
# ATS 기프티콘 예산품의 (네이버페이 15개, 스타벅스 8개)
node fill.js ats-b --nq 15 --sq 8

# Placement Survey 예산품의 (2026년)
node fill.js survey-b --yr 2026

# 산학협력 프리랜서 구매품의
node fill.js free-p --sup EGI --st 2026/4/1 --en 2026/5/16

# 제출 전 검토 (폼 입력 후 멈춤, 직접 확인 가능)
node fill.js survey-b --dry-run
```

### 전체 옵션

| 옵션 | 설명 | 기본값 |
|---|---|---|
| `--nq` | 네이버페이 수량 (ats-b/p) | 12 |
| `--sq` | 스타벅스 수량 (ats-b/p) | 10 |
| `--mon` | 신청월 YYYYMM (ats-b, free-b) | 오늘 기준 |
| `--si` | 시행일 YYYY.MM.DD (ats-b) | 빈값 |
| `--yr` | 연도 (survey-b/p) | 올해 |
| `--amt` | 금액 원 단위 (free-b) | 5000000 |
| `--sup` | 공급사명 (free-p) | EGI |
| `--st` | 계약시작일 (free-p) | 빈값 |
| `--en` | 계약종료일 (free-p) | 빈값 |
| `--dry-run` | 입력 후 제출하지 않음 | false |
| `--headless` | 브라우저 창 없이 실행 | false |

---

## 첫 실행 권장 순서

1. `--dry-run` 으로 먼저 실행해서 폼이 잘 채워지는지 육안 확인
2. 셀렉터가 맞지 않는 항목은 `fill.js` 내 해당 함수의 셀렉터 수정
3. 이상 없으면 `--dry-run` 제거하고 실제 제출

---

## 셀렉터 수정이 필요할 때

사이트의 실제 input `id` / `name` 속성은 개발자 도구(F12)로 확인합니다.

```
F12 → Elements → 원하는 필드 클릭 → id / name 확인
```

`fill.js` 에서 해당 함수의 셀렉터 문자열만 교체하면 됩니다.  
각 함수는 독립적으로 분리되어 있어 수정이 쉽습니다.

| 함수 | 담당 |
|---|---|
| `fillBasicInfo` | 문건제목, 읽기권한, 문건분류, 시행일 |
| `fillBudgetTable` | 예산 연동 테이블 (예산품의) |
| `fillPurchaseForm` | 구매품의 양식 (품목 테이블 포함) |
| `fillBody` | 본문 에디터 |
| `submitDoc` | 문건등록/상신 버튼 클릭 |

---

## 주의사항

- `.env` 파일은 절대 Git에 올리지 마세요 (`.gitignore`에 추가 권장)
- 제출 전 반드시 `--dry-run` 으로 검토하세요
- 오류 발생 시 `error_screenshot.png` 파일로 원인 파악 가능
