require('dotenv').config();
const { execSync } = require('child_process');

async function parseCommand(input) {
  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': process.env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 500,
      messages: [{
        role: 'user',
        content: `전자결재 명령을 JSON으로만 파싱해. JSON 외 다른 텍스트 절대 금지.

문서유형:
- ats-b: ATS 기프티콘 예산품의
- ats-p: ATS 기프티콘 구매품의
- survey-b: Placement Survey 예산품의
- survey-p: Placement Survey 구매품의
- free-b: 산학협력 프리랜서 예산품의
- free-p: 산학협력 프리랜서 구매품의
- interview-b: 산학협력 인터뷰 비용 지급 예산품의

파싱 규칙:
- amt: 금액(원단위. 10만원→100000, 500만원→5000000, 1500만원→15000000)
- mon: 단일월 YYYYMM (연도 없으면 2026 기준. 12월→202612, 6월→202606, 3월→202603)
- months: 복수월 쉼표구분 (3월과 6월→202603,202606)
- yr: 연도(기본2026)
- nq: 네이버페이수량(숫자)
- sq: 스타벅스수량(숫자)
- nprice: 네이버페이 단가 원단위(5만원→50000, ats-b용, 기본30000)
- sprice: 스타벅스 단가 원단위(2만원→20000, ats-b용, 기본10000)
- n: 산학협력 인원수(숫자, free-b용, 기본9)
- weeks: 산학협력 진행주차(숫자, free-b용, 기본6)
- n30: 30분 인터뷰 인원(숫자, interview-b용)
- n60: 1시간 인터뷰 인원(숫자, interview-b용)
- rate30: 30분 단가 원단위(2만원→20000, interview-b용)
- rate60: 1시간 단가 원단위(4만원→40000, interview-b용)
- from_mon: FROM 신청월 YYYYMM (interview-b용, 차감할 월)

예시:
입력: "placement survey 10만원으로 12월 예산품의"
출력: {"type":"survey-b","amt":100000,"mon":"202612"}

입력: "6월 산학협력 프리랜서 예산품의"
출력: {"type":"free-b","mon":"202606"}

입력: "ATS 기프티콘 네이버페이 20개 스타벅스 5개 구매품의"
출력: {"type":"ats-p","nq":20,"sq":5}

명령: ${input}`
      }]
    })
  });

  const data = await response.json();

  if (!data.content || !data.content[0]) {
    console.error('API 응답 오류:', JSON.stringify(data));
    process.exit(1);
  }

  const text = data.content[0].text.trim();
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) {
    console.error('JSON 파싱 실패. API 응답:', text);
    process.exit(1);
  }

  return JSON.parse(match[0]);
}

async function main() {
  const input = process.argv.slice(2).join(' ');
  if (!input) {
    console.log('\n사용법: node auto.js "명령어"');
    console.log('\n예시:');
    console.log('  node auto.js "placement survey 10만원으로 12월 예산품의"');
    console.log('  node auto.js "6월 산학협력 프리랜서 예산품의"');
    console.log('  node auto.js "ATS 기프티콘 네이버페이 20개 스타벅스 5개 구매품의"');
    console.log('  node auto.js "survey 3월 6월 각 500만원 예산품의"');
    process.exit(1);
  }

  if (!process.env.ANTHROPIC_API_KEY) {
    console.error('\n.env 파일에 ANTHROPIC_API_KEY가 없어요!');
    process.exit(1);
  }

  console.log(`\n명령어 분석 중: "${input}"`);
  const parsed = await parseCommand(input);
  console.log('분석 결과:', JSON.stringify(parsed, null, 2));

  let cmd = `node fill.js ${parsed.type}`;
  if (parsed.amt    != null) cmd += ` --amt ${parsed.amt}`;
  if (parsed.mon    != null) cmd += ` --mon ${parsed.mon}`;
  if (parsed.months != null) cmd += ` --months ${parsed.months}`;
  if (parsed.yr     != null) cmd += ` --yr ${parsed.yr}`;
  if (parsed.nq     != null) cmd += ` --nq ${parsed.nq}`;
  if (parsed.sq      != null) cmd += ` --sq ${parsed.sq}`;
  if (parsed.n30     != null) cmd += ` --n30 ${parsed.n30}`;
  if (parsed.n60     != null) cmd += ` --n60 ${parsed.n60}`;
  if (parsed.from_mon != null) cmd += ` --from_mon ${parsed.from_mon}`;
  if (parsed.rate30   != null) cmd += ` --rate30 ${parsed.rate30}`;
  if (parsed.rate60   != null) cmd += ` --rate60 ${parsed.rate60}`;
  if (parsed.nprice   != null) cmd += ` --nprice ${parsed.nprice}`;
  if (parsed.sprice   != null) cmd += ` --sprice ${parsed.sprice}`;
  if (parsed.n        != null) cmd += ` --n ${parsed.n}`;
  if (parsed.weeks    != null) cmd += ` --weeks ${parsed.weeks}`;
  cmd += ` --dry-run`;

  console.log(`\n실행: ${cmd}\n`);
  execSync(cmd, { stdio: 'inherit' });
}

main().catch(e => {
  console.error('오류:', e.message);
  process.exit(1);
});
