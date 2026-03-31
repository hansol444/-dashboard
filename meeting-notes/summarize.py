"""
회의록 정리 에이전트
TXT 녹취록 → Claude 구조화 요약 → JSON 출력 (+ 선택적 Notion 등록)

사용법:
  python summarize.py input/회의록.txt
  python summarize.py input/회의록.txt --notion   # Notion에도 등록
  python summarize.py input/회의록.txt --dry-run   # 요약만 보기

출력:
  1. 콘솔에 요약 출력
  2. output/ 폴더에 JSON 저장
  3. --notion 시 Notion DB에 페이지 생성
"""

import sys
import io
import json
import os
from pathlib import Path
from datetime import datetime

# Windows 콘솔 UTF-8 출력
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from dotenv import load_dotenv
import anthropic

load_dotenv()

SYSTEM_PROMPT = """당신은 전략추진실 회의록 정리 에이전트입니다.
녹취록 TXT를 받아서 다음 3가지를 추출합니다:

1. **구조화 요약** — 회의 핵심 내용을 항목별로 정리
2. **업무 지시 추출** — 누가 누구에게 무엇을 언제까지 (→ Active Projects 연동)
3. **방향성/기대 변화** — 팀 리더가 인턴에게 바라는 역할, 일하는 방식 변화 (→ Team Context 연동)

반드시 아래 JSON 형식으로 출력하세요:

{
  "meeting_date": "YYYY-MM-DD",
  "participants": ["이름1", "이름2"],
  "duration_minutes": 30,
  "summary": "회의 핵심 요약 (2-3문장)",
  "key_topics": [
    {
      "topic": "주제명",
      "details": "상세 내용",
      "decisions": ["결정사항1", "결정사항2"]
    }
  ],
  "task_assignments": [
    {
      "assignee": "담당자",
      "task": "업무 내용",
      "deadline": "마감일 (있으면)",
      "priority": "high/medium/low"
    }
  ],
  "direction_changes": [
    {
      "from_who": "발언자",
      "content": "방향성/기대 변화 내용",
      "context": "어떤 맥락에서 나온 발언인지"
    }
  ],
  "action_items_for_dashboard": [
    {
      "from": "지시자",
      "to": "담당자",
      "message": "업무 내용 (슬랙 메시지처럼)",
      "deadline": "마감일",
      "category_hint": "라우팅 키워드 힌트 (예: 장표, 매크로, 번역 등)"
    }
  ]
}

주의사항:
- 업무 지시와 방향성 변화를 명확히 구분하세요
- 업무 지시: "~해줘", "~까지 마무리", "~진행해" 등 구체적 action
- 방향성 변화: "앞으로는 ~", "우리 팀은 ~해야", "인턴이 ~역할을" 등 기대/역할 변화
- 날짜가 불명확하면 "미정"으로 표기
- 참석자 이름이 불명확하면 "미확인"으로 표기
"""


def read_transcript(file_path: str) -> str:
    path = Path(file_path)
    if not path.exists():
        print(f"파일을 찾을 수 없습니다: {file_path}")
        sys.exit(1)

    encodings = ["utf-8", "cp949", "euc-kr", "utf-16"]
    for enc in encodings:
        try:
            return path.read_text(encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue

    print(f"파일 인코딩을 읽을 수 없습니다: {file_path}")
    sys.exit(1)


def summarize_with_claude(transcript: str) -> dict:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("ANTHROPIC_API_KEY가 설정되지 않았습니다. .env 파일을 확인하세요.")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": f"다음 회의 녹취록을 정리해주세요:\n\n{transcript}"
            }
        ]
    )

    response_text = message.content[0].text

    # JSON 블록 추출
    if "```json" in response_text:
        json_str = response_text.split("```json")[1].split("```")[0].strip()
    elif "```" in response_text:
        json_str = response_text.split("```")[1].split("```")[0].strip()
    else:
        json_str = response_text.strip()

    return json.loads(json_str)


def save_output(result: dict, input_path: str) -> tuple[str, str]:
    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)

    stem = Path(input_path).stem
    date_str = result.get("meeting_date", datetime.now().strftime("%Y-%m-%d"))
    output_name = f"{date_str}_{stem}.json"
    output_path = output_dir / output_name

    output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(output_path), output_name


def upload_to_github(result: dict, output_name: str) -> bool:
    """GitHub repo의 data/meeting-notes/에 요약 JSON 업로드"""
    import base64
    import urllib.request

    github_token = os.getenv("GITHUB_TOKEN")
    if not github_token:
        return False

    repo = "IMHY-dev/-dashboard"
    file_path = f"data/meeting-notes/{output_name}"
    content = json.dumps(result, ensure_ascii=False, indent=2)
    encoded = base64.b64encode(content.encode("utf-8")).decode("ascii")
    headers = {
        "Authorization": f"token {github_token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
    }

    # 기존 파일 SHA 확인
    sha = None
    try:
        req = urllib.request.Request(
            f"https://api.github.com/repos/{repo}/contents/{file_path}",
            headers=headers,
        )
        with urllib.request.urlopen(req) as resp:
            sha = json.loads(resp.read()).get("sha")
    except Exception:
        pass

    body: dict = {"message": f"sync meeting notes {output_name}", "content": encoded}
    if sha:
        body["sha"] = sha

    try:
        req = urllib.request.Request(
            f"https://api.github.com/repos/{repo}/contents/{file_path}",
            data=json.dumps(body).encode("utf-8"),
            headers=headers,
            method="PUT",
        )
        with urllib.request.urlopen(req) as resp:
            return resp.status in (200, 201)
    except Exception:
        return False


def upload_to_notion(result: dict) -> str:
    """Notion Meeting Notes DB에 페이지 생성"""
    notion_key = os.getenv("NOTION_API_KEY")
    db_id = os.getenv("NOTION_DATABASE_ID")

    if not notion_key or not db_id:
        print("NOTION_API_KEY 또는 NOTION_DATABASE_ID가 설정되지 않았습니다.")
        return ""

    import requests

    # 요약 본문 구성
    body_lines = []
    body_lines.append(f"## 요약\n{result['summary']}\n")

    if result.get("key_topics"):
        body_lines.append("## 주요 안건")
        for topic in result["key_topics"]:
            body_lines.append(f"### {topic['topic']}")
            body_lines.append(topic["details"])
            if topic.get("decisions"):
                for d in topic["decisions"]:
                    body_lines.append(f"- ✅ {d}")
            body_lines.append("")

    if result.get("task_assignments"):
        body_lines.append("## 업무 지시")
        for ta in result["task_assignments"]:
            deadline = f" (~{ta['deadline']})" if ta.get("deadline") and ta["deadline"] != "미정" else ""
            body_lines.append(f"- **{ta['assignee']}**: {ta['task']}{deadline}")
        body_lines.append("")

    if result.get("direction_changes"):
        body_lines.append("## 방향성 변화")
        for dc in result["direction_changes"]:
            body_lines.append(f"- **{dc['from_who']}**: {dc['content']}")
            body_lines.append(f"  - 맥락: {dc['context']}")
        body_lines.append("")

    body_text = "\n".join(body_lines)

    # Notion API 호출
    participants = ", ".join(result.get("participants", []))
    meeting_date = result.get("meeting_date", datetime.now().strftime("%Y-%m-%d"))
    title = f"[{meeting_date}] {result.get('summary', '회의록')[:50]}"

    # 본문을 paragraph 블록으로 변환
    children = []
    for line in body_text.split("\n"):
        if not line.strip():
            continue
        if line.startswith("## "):
            children.append({
                "object": "block",
                "type": "heading_2",
                "heading_2": {"rich_text": [{"type": "text", "text": {"content": line[3:]}}]}
            })
        elif line.startswith("### "):
            children.append({
                "object": "block",
                "type": "heading_3",
                "heading_3": {"rich_text": [{"type": "text", "text": {"content": line[4:]}}]}
            })
        elif line.startswith("- "):
            children.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": line[2:]}}]}
            })
        else:
            children.append({
                "object": "block",
                "type": "paragraph",
                "paragraph": {"rich_text": [{"type": "text", "text": {"content": line}}]}
            })

    payload = {
        "parent": {"database_id": db_id},
        "properties": {
            "Name": {"title": [{"text": {"content": title}}]},
        },
        "children": children[:100]  # Notion API 제한: 100블록
    }

    resp = requests.post(
        "https://api.notion.com/v1/pages",
        headers={
            "Authorization": f"Bearer {notion_key}",
            "Content-Type": "application/json",
            "Notion-Version": "2022-06-28",
        },
        json=payload,
    )

    if resp.status_code == 200:
        page_url = resp.json().get("url", "")
        return page_url
    else:
        print(f"Notion 업로드 실패: {resp.status_code} {resp.text}")
        return ""


def print_summary(result: dict):
    print("\n" + "=" * 60)
    print(f"📋 회의록 요약 — {result.get('meeting_date', '날짜 미상')}")
    print(f"   참석자: {', '.join(result.get('participants', []))}")
    print(f"   소요시간: {result.get('duration_minutes', '?')}분")
    print("=" * 60)

    print(f"\n📝 요약: {result['summary']}")

    if result.get("key_topics"):
        print("\n📌 주요 안건:")
        for t in result["key_topics"]:
            print(f"  • {t['topic']}: {t['details'][:80]}")
            for d in t.get("decisions", []):
                print(f"    ✅ {d}")

    if result.get("task_assignments"):
        print("\n📋 업무 지시:")
        for ta in result["task_assignments"]:
            deadline = f" (~{ta['deadline']})" if ta.get("deadline") and ta["deadline"] != "미정" else ""
            print(f"  → {ta['assignee']}: {ta['task']}{deadline} [{ta.get('priority', 'medium')}]")

    if result.get("direction_changes"):
        print("\n🧭 방향성 변화:")
        for dc in result["direction_changes"]:
            print(f"  → {dc['from_who']}: {dc['content']}")

    if result.get("action_items_for_dashboard"):
        print("\n🎯 대시보드 연동 항목:")
        for item in result["action_items_for_dashboard"]:
            print(f"  📨 {item['from']} → {item['to']}: \"{item['message']}\" (~{item.get('deadline', '미정')})")

    print("\n" + "=" * 60)


def main():
    if len(sys.argv) < 2:
        print("사용법: python summarize.py <녹취록.txt> [--notion] [--dry-run]")
        print("  --notion   Notion DB에도 등록")
        print("  --dry-run  Claude 호출 없이 파일 읽기만 테스트")
        sys.exit(1)

    input_file = sys.argv[1]
    use_notion = "--notion" in sys.argv
    dry_run = "--dry-run" in sys.argv

    print(f"📂 파일 읽는 중: {input_file}")
    transcript = read_transcript(input_file)
    print(f"   → {len(transcript)}자 로드 완료")

    if dry_run:
        print("\n[dry-run] 파일 읽기 성공. Claude 호출 없이 종료.")
        print(f"   첫 200자: {transcript[:200]}...")
        return

    print("🤖 Claude에게 요약 요청 중...")
    result = summarize_with_claude(transcript)

    print_summary(result)

    output_path, output_name = save_output(result, input_file)
    print(f"\n💾 JSON 저장: {output_path}")

    if upload_to_github(result, output_name):
        print(f"☁️  GitHub 업로드 완료: data/meeting-notes/{output_name}")
    else:
        print("⚠️  GitHub 업로드 실패 (GITHUB_TOKEN 확인)")

    if use_notion:
        print("📤 Notion 업로드 중...")
        url = upload_to_notion(result)
        if url:
            print(f"   → Notion 페이지: {url}")

    # 대시보드 연동용 action items를 별도 파일로 저장
    if result.get("action_items_for_dashboard"):
        actions_path = Path(__file__).parent / "output" / "pending_actions.json"
        existing = []
        if actions_path.exists():
            existing = json.loads(actions_path.read_text(encoding="utf-8"))
        existing.extend(result["action_items_for_dashboard"])
        actions_path.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"📊 대시보드 연동 항목 {len(result['action_items_for_dashboard'])}개 저장")


if __name__ == "__main__":
    main()
