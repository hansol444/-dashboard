require('dotenv').config();
const express = require('express');
const { spawn } = require('child_process');
const path = require('path');
const app = express();

app.use(express.json());
app.use(express.static(__dirname));

// 문건등록 실행 API
app.post('/run', (req, res) => {
  const { type, args } = req.body;
  if (!type) return res.status(400).json({ error: '문서유형이 없습니다.' });

  // 명령어 조합
  const cmdArgs = ['fill.js', type];
  const argMap = {
    amt: '--amt', mon: '--mon', months: '--months', yr: '--yr',
    nq: '--nq', sq: '--sq', nprice: '--nprice', sprice: '--sprice',
    n: '--n', weeks: '--weeks', hpw: '--hpw',
    n30: '--n30', n60: '--n60', rate30: '--rate30', rate60: '--rate60',
    from_mon: '--from_mon', from_cc: '--from_cc', to_cc: '--to_cc',
    sup: '--sup', st: '--st', en: '--en',
    parent: '--parent', files: '--files',
  };

  Object.entries(argMap).forEach(([key, flag]) => {
    if (args[key] != null && args[key] !== '') {
      cmdArgs.push(flag, String(args[key]));
    }
  });

  cmdArgs.push('--dry-run');

  console.log('\n실행:', 'node', cmdArgs.join(' '));

  // fill.js 실행
  const child = spawn('node', cmdArgs, {
    cwd: __dirname,
    stdio: 'pipe',
    detached: false,
  });

  let output = '';
  child.stdout.on('data', d => { output += d.toString(); process.stdout.write(d); });
  child.stderr.on('data', d => { output += d.toString(); process.stderr.write(d); });

  child.on('close', (code) => {
    res.json({ success: code === 0, output, code });
  });

  child.on('error', (err) => {
    res.status(500).json({ error: err.message });
  });
});

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`\n✅ 전자결재 자동화 서버 실행 중`);
  console.log(`   브라우저에서 열기: http://localhost:${PORT}`);
  console.log(`   종료: Ctrl+C\n`);
});
