"""
config.py — 전역 설정
환경변수 ANTHROPIC_API_KEY가 설정되어 있으면 그것을 우선 사용합니다.
"""

import os

# ─── API ───────────────────────────────────────────────────────────────────────
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL = "claude-sonnet-4-6"

# ─── 번역 설정 ─────────────────────────────────────────────────────────────────
TARGET_LANGUAGE = "en"          # "en" 또는 "ko"
ENGLISH_STYLE = "Australian"    # Australian English
DOMAIN = "HR/Recruitment"

# ─── 품질 모드 ─────────────────────────────────────────────────────────────────
# "fast"    : 슬라이드 단위 번역 (빠름, 기본값)
# "precise" : 텍스트 블록별 개별 번역 (느리지만 정확)
QUALITY_MODE = "fast"

# ─── 글자수 초과 처리 ──────────────────────────────────────────────────────────
# 초과 시 재번역 시도 횟수 (축약 허용)
OVERFLOW_RETRY_COUNT = 1
# 재번역 후에도 초과 시 텍스트박스 전체를 빨간색으로 표시
OVERFLOW_HIGHLIGHT_COLOR = (255, 0, 0)  # RGB 빨간색

# ─── 후처리 규칙 (False로 바꾸면 해당 규칙 비활성화) ──────────────────────────
POST_PROCESS_RULES = {
    "fix_billion": True,
    "fix_duplicates": True,
    "fix_month_abbrev": True,
    "fix_currency_order": True,
    "fix_australian_spelling": True,
}

def get_enabled_rules() -> list[str]:
    """활성화된 후처리 규칙 이름 목록 반환."""
    return [rule for rule, enabled in POST_PROCESS_RULES.items() if enabled]

# ─── 병렬 처리 ─────────────────────────────────────────────────────────────────
# 동시에 번역할 최대 슬라이드 수 (높을수록 빠름, 너무 높으면 rate limit 발생)
# 권장: 4~6  /  rate limit 자주 걸리면 2~3으로 낮추세요
MAX_PARALLEL_SLIDES = 4

# ─── API 재시도 ────────────────────────────────────────────────────────────────
API_MAX_RETRIES = 3
API_RETRY_DELAY = 2  # 초

# ─── 경로 ─────────────────────────────────────────────────────────────────────
INPUT_DIR = "input"
OUTPUT_DIR = "output"
TERMINOLOGY_PATH = "terminology.json"
SYSTEM_PROMPT_PATH = "SYSTEM_PROMPT.txt"

# ─── SharePoint 연동 ───────────────────────────────────────────────────────────
# Azure Portal > App registrations 에서 발급
SHAREPOINT_TENANT_ID    = os.environ.get("SHAREPOINT_TENANT_ID", "")
SHAREPOINT_CLIENT_ID    = os.environ.get("SHAREPOINT_CLIENT_ID", "")
SHAREPOINT_CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET", "")
# 업로드 대상 SharePoint 폴더 공유 링크
SHAREPOINT_SHARE_URL = os.environ.get("SHAREPOINT_SHARE_URL", "")

# ─── Slack 연동 ────────────────────────────────────────────────────────────────
# Slack App > OAuth & Permissions 에서 발급 (xoxb-... 형식)
SLACK_BOT_TOKEN   = os.environ.get("SLACK_BOT_TOKEN", "")
SLACK_CHANNEL_ID  = os.environ.get("SLACK_CHANNEL_ID", "")
SLACK_MESSAGE_TEMPLATE = "번역 요청 PPT 초안입니다!"
