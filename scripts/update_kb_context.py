#!/usr/bin/env python3
"""
Daily KB Context Updater
Slack + Notion 새 내용 → Context/Topics/ KB 파일 자동 업데이트
"""

import os
import json
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path

import anthropic
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

# ── 설정 ────────────────────────────────────────────────
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
SLACK_TOKEN       = os.environ["SLACK_TOKEN"]

SLACK_CHANNEL_NAMES = [
    "전략추진실-창준님",
    "전략추진실-all",
]

KB_BASE    = Path("Context/Topics")
STATE_FILE = Path("scripts/.last_update")

# KB 파일 목록 및 담당 업무 설명 (Claude 라우팅용)
KB_FILES = {
    "budget/placement-concur.md":  "Placement Survey 예산품의·구매품의·Concur 처리",
    "budget/consulting-concur.md": "BCG 컨설팅 자문료 Concur (착수금·중도금·잔금)",
    "budget/enkoline-concur.md":   "엔코라인 구독 Concur 처리",
    "budget/law-firm.md":          "법무법인(태평양·LAB) 청구서 및 Concur",
    "budget/academia-contract.md": "산학협력(EGI·MCSA) 계약 및 Concur",
    "budget/gifticon.md":          "기프티콘 구매·발송 및 Concur (KT alpha)",
    "budget/ninehire.md":          "나인하이어 에스크로·주식매매대금·스톡옵션 지급",
    "budget/budget-101.md":        "예산품의·구매품의 기초 개념 및 프로세스",
    "fluid/ppt-work.md":           "장표 제작 및 번역 (한→영)",
    "fluid/academia-contract.md":  "산학협력 계약 전반",
    "fluid/meeting-notes.md":      "회의록 요약 및 Notion 등록",
    "regular/macro-update.md":     "KOSIS 매크로 데이터 업데이트 (update_macro.py)",
    "regular/macro-indicators.md": "매크로 신규 지표 발굴 및 추가",
    "regular/placement-update.md": "Placement Survey 쿼터·설문 변경",
    "regular/placement-analysis.md": "Placement Survey RMS 분석",
}


# ── 유틸 ────────────────────────────────────────────────
def load_last_update() -> datetime:
    """마지막 실행 시각 로드. 없으면 7일 전."""
    if STATE_FILE.exists():
        ts = STATE_FILE.read_text().strip()
        return datetime.fromisoformat(ts)
    return datetime.now(timezone.utc) - timedelta(days=7)


def save_last_update():
    STATE_FILE.write_text(datetime.now(timezone.utc).isoformat())


# ── Slack ────────────────────────────────────────────────
def get_channel_ids(client: WebClient, names: list[str]) -> dict[str, str]:
    """채널 이름 → ID 매핑."""
    result = {}
    for resp in client.conversations_list(types="public_channel,private_channel", limit=200):
        for ch in resp["channels"]:
            if ch["name"] in names:
                result[ch["name"]] = ch["id"]
        if not resp.get("response_metadata", {}).get("next_cursor"):
            break
    return result


def fetch_slack_messages(since: datetime) -> list[dict]:
    """since 이후 관련 채널 메시지 수집."""
    client = WebClient(token=SLACK_TOKEN)
    channel_ids = get_channel_ids(client, SLACK_CHANNEL_NAMES)
    oldest = str(since.timestamp())

    messages = []
    for ch_name, ch_id in channel_ids.items():
        try:
            resp = client.conversations_history(channel=ch_id, oldest=oldest, limit=200)
            for msg in resp.get("messages", []):
                if msg.get("text") and not msg.get("bot_id"):
                    ts = datetime.fromtimestamp(float(msg["ts"]), tz=timezone.utc)
                    # 발언자 이름 조회
                    user_id = msg.get("user", "")
                    try:
                        user_info = client.users_info(user=user_id)
                        display_name = user_info["user"]["profile"].get("display_name") or \
                                       user_info["user"]["profile"].get("real_name", user_id)
                    except Exception:
                        display_name = user_id

                    messages.append({
                        "source": "slack",
                        "channel": ch_name,
                        "date": ts.strftime("%Y-%m-%d"),
                        "author": display_name,
                        "text": msg["text"][:500],
                    })
        except SlackApiError as e:
            print(f"Slack error ({ch_name}): {e.response['error']}")

    return messages


# ── Claude 분석 ──────────────────────────────────────────
SYSTEM_PROMPT = """너는 전략추진실 인턴 워크스페이스의 KB(지식베이스) 자동 업데이터야.
새로운 Slack 메시지와 Notion 싱크 기록을 받아, 어떤 KB 파일의 어떤 섹션에 추가할지 분석해.

규칙:
1. 업무와 관련 없는 잡담, 인사, 단순 확인 메시지는 무시.
2. 업무 지시, 결정사항, 프로세스 변경, 방향성 발언만 수집.
3. 하나의 메시지가 여러 파일에 해당할 수 있음.
4. Slack 메시지 → "slack" 섹션 (## ── 맥락 (Slack) ──)
5. Notion 싱크 → "sync" 섹션 (## ── 맥락 (SYNC) ──)

JSON 배열로 응답. 업데이트할 내용이 없으면 빈 배열 [].
각 항목 형식:
{
  "file": "budget/placement-concur.md",
  "section": "slack",  // "slack" 또는 "sync"
  "row": "| 2026-04-01 | #전략추진실-창준님 | 창준님 | 내용 요약 |"
}

중요: row는 기존 테이블과 동일한 마크다운 파이프 형식으로 작성.
- slack 행: | 날짜 | 채널/출처 | 발언자 | 핵심 내용 |
- sync 행: | 날짜 | 미팅명 | 핵심 결정사항 |
"""


def analyze_with_claude(new_items: list[dict]) -> list[dict]:
    """새 내용을 Claude로 분석해 KB 업데이트 목록 반환."""
    if not new_items:
        return []

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    kb_list = "\n".join(f"- {k}: {v}" for k, v in KB_FILES.items())
    items_json = json.dumps(new_items, ensure_ascii=False, indent=2)

    user_msg = f"""KB 파일 목록:
{kb_list}

새로 수집된 내용:
{items_json}

위 내용 중 KB에 추가할 항목을 JSON 배열로 반환해."""

    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_msg}],
    )

    text = resp.content[0].text.strip()
    # JSON 추출 (코드블록 감싸진 경우 처리)
    match = re.search(r"\[.*\]", text, re.DOTALL)
    if match:
        return json.loads(match.group())
    return []


# ── KB 파일 업데이트 ─────────────────────────────────────
def apply_updates(updates: list[dict]):
    """각 KB 파일의 해당 섹션 테이블에 행 추가."""
    for item in updates:
        filepath = KB_BASE / item["file"]
        if not filepath.exists():
            print(f"파일 없음: {filepath}")
            continue

        content = filepath.read_text(encoding="utf-8")
        row = item["row"].strip()
        section_header = (
            "## ── 맥락 (Slack) ──" if item["section"] == "slack"
            else "## ── 맥락 (SYNC) ──"
        )
        summary_marker = "→ **핵심"

        # 해당 섹션 찾아서 요약 줄 바로 앞에 행 삽입
        if section_header not in content:
            print(f"섹션 없음 ({section_header}): {filepath}")
            continue

        # 중복 방지: 이미 동일 행이 있으면 스킵
        if row[:30] in content:
            print(f"이미 존재: {row[:40]}...")
            continue

        # 섹션 내 요약 줄 앞에 삽입
        section_start = content.index(section_header)
        summary_pos = content.find(summary_marker, section_start)
        if summary_pos == -1:
            print(f"요약 줄 없음: {filepath}")
            continue

        # 요약 줄 앞 빈 줄 위치 찾기 (테이블 마지막 행 다음)
        insert_pos = summary_pos
        # 이전 줄 끝 찾기
        line_start = content.rfind("\n", 0, summary_pos)
        insert_pos = line_start  # \n 앞에 삽입

        updated = content[:insert_pos] + "\n" + row + content[insert_pos:]
        filepath.write_text(updated, encoding="utf-8")
        print(f"✓ 추가: {item['file']} ({item['section']}) → {row[:60]}...")


# ── 메인 ────────────────────────────────────────────────
def main():
    since = load_last_update()
    print(f"Last update: {since.isoformat()}")
    print(f"Fetching content since: {since.strftime('%Y-%m-%d')}")

    # 1. 새 내용 수집
    slack_messages = fetch_slack_messages(since)
    print(f"Slack messages: {len(slack_messages)}")

    all_new = slack_messages
    if not all_new:
        print("새 내용 없음. 종료.")
        save_last_update()
        return

    # 2. Claude 분석
    updates = analyze_with_claude(all_new)
    print(f"KB 업데이트 항목: {len(updates)}")

    # 3. 파일 업데이트
    apply_updates(updates)

    # 4. 타임스탬프 저장
    save_last_update()
    print("완료.")


if __name__ == "__main__":
    main()
