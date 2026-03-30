/**
 * GitHub 파일 기반 DB 헬퍼
 * data/tasks.json을 IMHY-dev/-dashboard 레포에서 읽고 씁니다.
 * Vercel 환경변수: GITHUB_TOKEN (repo 스코프 PAT)
 */

const REPO = "IMHY-dev/-dashboard";
const TASKS_FILE = "data/tasks.json";

export interface GHTask {
  id: string;
  from: string;
  to: string;
  message: string;
  channel: string;
  timestamp: string;
  deadline: string;
  status: "pending" | "in_progress" | "done";
  slackTs?: string;
  notes?: string[]; // 스레드 댓글 키워드
}

// ─── 내부 헬퍼 ───

async function ghGet(filePath: string) {
  const token = process.env.GITHUB_TOKEN;
  const res = await fetch(
    `https://api.github.com/repos/${REPO}/contents/${filePath}`,
    {
      headers: {
        Authorization: `token ${token}`,
        Accept: "application/vnd.github.v3+json",
      },
      cache: "no-store",
    }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`GitHub GET failed: ${res.status}`);
  return res.json();
}

async function ghPut(filePath: string, content: string, sha: string | null, message: string) {
  const token = process.env.GITHUB_TOKEN;
  const body: Record<string, string> = {
    message,
    content: Buffer.from(content).toString("base64"),
  };
  if (sha) body.sha = sha;

  const res = await fetch(
    `https://api.github.com/repos/${REPO}/contents/${filePath}`,
    {
      method: "PUT",
      headers: {
        Authorization: `token ${token}`,
        Accept: "application/vnd.github.v3+json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    }
  );
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(`GitHub PUT failed: ${res.status} ${JSON.stringify(err)}`);
  }
}

// ─── Public API ───

/** 전체 태스크 읽기 */
export async function getAllTasks(): Promise<GHTask[]> {
  const file = await ghGet(TASKS_FILE);
  if (!file) return [];
  try {
    const content = Buffer.from(file.content.replace(/\s/g, ""), "base64").toString("utf-8");
    return JSON.parse(content) as GHTask[];
  } catch {
    return [];
  }
}

/** 태스크 추가 또는 업데이트 (id 기준 upsert) */
export async function upsertTask(task: GHTask): Promise<void> {
  const file = await ghGet(TASKS_FILE);
  const sha: string | null = file?.sha ?? null;
  let tasks: GHTask[] = [];

  if (file) {
    try {
      const content = Buffer.from(file.content.replace(/\s/g, ""), "base64").toString("utf-8");
      tasks = JSON.parse(content);
    } catch {
      tasks = [];
    }
  }

  const idx = tasks.findIndex((t) => t.id === task.id);
  if (idx >= 0) {
    tasks[idx] = { ...tasks[idx], ...task };
  } else {
    tasks.push(task);
  }

  await ghPut(TASKS_FILE, JSON.stringify(tasks, null, 2), sha, `upsert task ${task.id}`);
}

/** 태스크 삭제 */
export async function removeTask(taskId: string): Promise<void> {
  const file = await ghGet(TASKS_FILE);
  if (!file) return;

  const content = Buffer.from(file.content.replace(/\s/g, ""), "base64").toString("utf-8");
  const tasks: GHTask[] = JSON.parse(content);
  const filtered = tasks.filter((t) => t.id !== taskId);

  if (filtered.length === tasks.length) return; // 변경 없으면 스킵

  await ghPut(TASKS_FILE, JSON.stringify(filtered, null, 2), file.sha, `remove task ${taskId}`);
}

/** 특정 필드만 업데이트 */
export async function patchTask(taskId: string, patch: Partial<GHTask>): Promise<void> {
  const file = await ghGet(TASKS_FILE);
  if (!file) return;

  const content = Buffer.from(file.content.replace(/\s/g, ""), "base64").toString("utf-8");
  const tasks: GHTask[] = JSON.parse(content);
  const updated = tasks.map((t) => (t.id === taskId ? { ...t, ...patch } : t));

  await ghPut(TASKS_FILE, JSON.stringify(updated, null, 2), file.sha, `patch task ${taskId}`);
}
