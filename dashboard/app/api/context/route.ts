import { NextRequest, NextResponse } from "next/server";

/**
 * /api/context
 *
 * GitHub data/meeting-notes/ 에서 회의록 요약 JSON을 읽어
 * topicFile 키워드와 매칭되는 싱크를 반환.
 *
 * NOTION_TOKEN 불필요 — GITHUB_TOKEN만 사용 (기존 설정 그대로)
 *
 * GET ?topicFile=regular/macro-update
 * Response: { results: [{ id, title, date, decisions, url }] }
 */

const REPO = "IMHY-dev/-dashboard";
const NOTES_DIR = "data/meeting-notes";

const TOPIC_KEYWORDS: Record<string, string[]> = {
  "regular/macro-update":       ["매크로", "KOSIS", "macro"],
  "regular/macro-indicators":   ["매크로", "지표", "indicator"],
  "regular/placement-update":   ["플레이스먼트", "서베이", "설문", "placement", "survey"],
  "regular/placement-analysis": ["플레이스먼트", "RMS", "placement", "분석"],
  "fluid/ppt-work":             ["장표", "PPT", "번역", "translate"],
  "fluid/academia-contract":    ["산학협력", "기프티콘", "계약"],
  "budget/placement-concur":    ["플레이스먼트", "컨커", "concur"],
  "budget/enkoline-concur":     ["엔코라인", "통역"],
  "budget/consulting-concur":   ["컨설팅", "BCG"],
  "budget/law-firm":            ["법무법인"],
  "budget/ninehire":            ["나인하이어", "에스크로", "스톡옵션"],
  "budget/gifticon":            ["기프티콘"],
  "budget/budget-transfer":     ["예산", "품의", "이월"],
  "budget/vendor-registration": ["공급사", "벤더"],
};

type MeetingNote = {
  meeting_date?: string;
  summary?: string;
  key_topics?: { topic: string; details: string; decisions?: string[] }[];
  task_assignments?: { assignee: string; task: string; deadline?: string }[];
};

async function ghGet(path: string) {
  const token = process.env.GITHUB_TOKEN;
  const res = await fetch(
    `https://api.github.com/repos/${REPO}/contents/${path}`,
    {
      headers: {
        Authorization: `token ${token}`,
        Accept: "application/vnd.github.v3+json",
      },
      cache: "no-store",
    }
  );
  if (!res.ok) return null;
  return res.json();
}

export async function GET(request: NextRequest) {
  const topicFile = new URL(request.url).searchParams.get("topicFile") ?? "";
  const keywords = TOPIC_KEYWORDS[topicFile] ?? [];
  if (keywords.length === 0) return NextResponse.json({ results: [] });

  // data/meeting-notes/ 디렉토리 파일 목록
  const dirData = await ghGet(NOTES_DIR);
  if (!dirData || !Array.isArray(dirData)) return NextResponse.json({ results: [] });

  const jsonFiles = dirData
    .filter((f: { name: string }) => f.name.endsWith(".json"))
    .sort((a: { name: string }, b: { name: string }) => b.name.localeCompare(a.name)) // 최신순
    .slice(0, 20); // 최근 20개만

  const results = [];

  for (const file of jsonFiles) {
    const fileData = await ghGet(`${NOTES_DIR}/${file.name}`);
    if (!fileData?.content) continue;

    let note: MeetingNote;
    try {
      const content = Buffer.from(fileData.content.replace(/\s/g, ""), "base64").toString("utf-8");
      note = JSON.parse(content);
    } catch {
      continue;
    }

    // 키워드 매칭 (summary + key_topics)
    const searchText = [
      note.summary ?? "",
      ...(note.key_topics ?? []).map((t) => `${t.topic} ${t.details}`),
    ].join(" ").toLowerCase();

    const matched = keywords.some((kw) => searchText.includes(kw.toLowerCase()));
    if (!matched) continue;

    const decisions = (note.key_topics ?? [])
      .flatMap((t) => t.decisions ?? [])
      .join(", ");

    results.push({
      id: file.name,
      title: note.summary?.slice(0, 60) ?? file.name,
      date: note.meeting_date ?? "",
      decisions,
      url: `https://github.com/${REPO}/blob/main/${NOTES_DIR}/${file.name}`,
    });

    if (results.length >= 5) break;
  }

  return NextResponse.json({ results });
}
