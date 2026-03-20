require('dotenv').config();
const fs = require('fs');
const path = require('path');
const CACHE_FILE = path.join(__dirname, '.doc_cache.json');

function saveDocNo(type, docNo) {
  let cache = {};
  try { cache = JSON.parse(fs.readFileSync(CACHE_FILE, 'utf8')); } catch(e) {}
  cache[type] = { docNo, savedAt: new Date().toISOString() };
  fs.writeFileSync(CACHE_FILE, JSON.stringify(cache, null, 2));
  console.log(`\n[💾] 문건번호 저장: ${type} → ${docNo}`);
}

function loadDocNo(type) {
  try {
    const cache = JSON.parse(fs.readFileSync(CACHE_FILE, 'utf8'));
    const entry = cache[type];
    if (entry) return entry.docNo;
  } catch(e) {}
  return null;
}
const { chromium } = require('playwright');
const argv = require('minimist')(process.argv.slice(2));
const TEMPLATES = require('./config/templates');

const URLS = {
  login:  'https://login.jobkorea.co.kr/Login',
  docReg: 'https://eda.jobkorea.co.kr/Eda/DocReg',
};

function log(msg)  { console.log(`\n[✓] ${msg}`); }
function warn(msg) { console.log(`[!] ${msg}`); }
function err(msg)  { console.error(`\n[✗] ${msg}`); }

async function main() {
  const docType = argv._[0];
  if (!docType || !TEMPLATES[docType]) {
    err(`문서유형이 잘못되었습니다: "${docType ?? ''}"`);
    process.exit(1);
  }

  const isDryRun   = !!argv['dry-run'];
  const isHeadless = !!argv['headless'];
  const data = TEMPLATES[docType](argv);

  console.log('\n══════════════════════════════════════════════');
  console.log(` Worxphere 전자결재 자동화`);
  console.log(` 문서유형: ${docType}  |  ${isDryRun ? '검토 모드 (제출 안 함)' : '제출 모드'}`);
  console.log(` 문건분류: ${data.기본.문건분류선택}`);
  if (data.예산연동) console.log(` 예산연동 월: ${data.예산연동.map(r => r.from.신청월).join(', ')}`);
  console.log('══════════════════════════════════════════════\n');

  const browser = await chromium.launch({ headless: isHeadless, slowMo: 80, channel: 'msedge' });
  const page = await browser.newPage();
  page.setDefaultTimeout(20000);

  try {
    log('전자결재 접속 중...');
    await page.goto(URLS.docReg, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(2000);

    for (let attempt = 0; attempt < 3; attempt++) {
      const url = page.url();

      if (url.includes('login.jobkorea') || url.includes('portal.jobkorea')) {
        log('SSO 로그인 중...');
        await page.waitForSelector('input[name="userID"]', { timeout: 10000 });
        await page.waitForTimeout(1000);
        // Vue 프레임워크 대응 - triple click으로 선택 후 입력
        await page.click('input[name="userID"]', { clickCount: 3 });
        await page.waitForTimeout(300);
        await page.keyboard.type(process.env.WORXPHERE_ID, { delay: 80 });
        await page.waitForTimeout(300);
        await page.click('input[name="password"]', { clickCount: 3 });
        await page.waitForTimeout(300);
        await page.keyboard.type(process.env.WORXPHERE_PW, { delay: 80 });
        await page.waitForTimeout(500);
        await page.keyboard.press('Enter');
        await page.waitForTimeout(4000);
        if (page.url().includes('login.jobkorea') || page.url().includes('portal.jobkorea')) {
          await page.goto(URLS.docReg, { waitUntil: 'domcontentloaded', timeout: 30000 });
        }
        await page.waitForTimeout(2000);
        continue;
      }

      if (await page.locator('text=재인증').count() > 0 || (url.includes('eda') && url.includes('Login'))) {
        log('전자결재 재인증 중...');
        await page.waitForSelector('input[type="password"]', { timeout: 10000 });
        await page.click('input[type="password"]');
        await page.keyboard.type(process.env.WORXPHERE_PW, { delay: 50 });
        await page.click('button:has-text("Login")');
        await page.waitForTimeout(3000);
        continue;
      }

      if (await page.locator('a:has-text("문건등록"), button:has-text("문건등록")').count() > 0) {
        log('로그인 완료!');
        break;
      }

      await page.waitForTimeout(2000);
    }

    // 구매품의인 경우 상위문건 자동 로드
    let parentDocNo = argv['parent'];
    if (!parentDocNo && docType && docType.endsWith('-p')) {
      const budgetType = docType.replace('-p', '-b');
      parentDocNo = loadDocNo(budgetType);
      if (parentDocNo) {
        log(`저장된 예산품의 문건번호 자동 로드: ${parentDocNo}`);
      }
    }

    if (parentDocNo) {
      log(`상위문건 ${parentDocNo} 에서 하위문건 등록 진행 중...`);
      await page.goto(`https://eda.jobkorea.co.kr/Eda/DocView?docNo=${parentDocNo}`, { waitUntil: 'networkidle', timeout: 30000 });
      await page.waitForTimeout(2000);
      const subBtn = page.locator('button:has-text("하위문건 등록"), a:has-text("하위문건 등록")').first();
      if (await subBtn.count() > 0) {
        await subBtn.click();
      } else {
        await page.evaluate(() => {
          const btn = Array.from(document.querySelectorAll('button, a')).find(b => b.textContent.includes('하위문건 등록'));
          if (btn) btn.click();
        });
      }
    } else {
      await page.click('a:has-text("문건등록"), button:has-text("문건등록")');
    }
    await page.waitForTimeout(2000);
    await page.waitForSelector('input[placeholder="문건제목을 입력해주세요"]', { timeout: 30000 });

    log('기본 정보 입력 중...');
    await fillBasicInfo(page, data.기본);

    if (data.예산연동) {
      log('예산 연동 테이블 입력 중...');
      await fillBudgetTable(page, data.예산연동);
    }
    if (data.구매품의) {
      log('구매품의 양식 입력 중...');
      await fillPurchaseForm(page, data.구매품의);
    }
    if (data.본문) {
      log('본문 입력 중...');
      await fillBody(page, data.본문);
    }

    // 첨부파일 처리
    const files = argv['files'] ? String(argv['files']).split(',').map(f => f.trim()) : [];
    if (files.length > 0) {
      log('첨부파일 업로드 중...');
      await attachFiles(page, files);
    }

    if (isDryRun) {
      warn('--dry-run 모드: 입력 완료. 직접 확인 후 브라우저를 닫으세요.');
      await page.waitForEvent('close', { timeout: 300000 });
    } else {
      log('결재 상신(제출) 중...');
      await submitDoc(page, docType);
      log('🎉 제출 완료!');
    }

  } catch (e) {
    err(`오류 발생: ${e.message}`);
    await page.screenshot({ path: 'error_screenshot.png' }).catch(() => {});
    await browser.close();
    process.exit(1);
  }

  await browser.close();
}

async function fillBasicInfo(page, info) {
  await page.fill('input[placeholder="문건제목을 입력해주세요"]', info.문건제목);

  const 읽기권한라벨 = page.locator(`label:has-text("${info.읽기권한}")`).first();
  await 읽기권한라벨.click();

  // 문건분류 팝업 열기
  const searchBtn = page.locator('button:has-text("검색"), a:has-text("검색")').first();
  await searchBtn.click();
  await page.waitForTimeout(2000);

  // 팝업에서 오른쪽 패널(주무부서 경유) 항목 클릭
  const clicked = await page.evaluate((target) => {
    const cells = document.querySelectorAll('.el-dialog__wrapper div.cell');
    for (const cell of cells) {
      if (cell.textContent.trim() === target) {
        // 해당 행의 라디오버튼 찾아서 클릭
        const row = cell.closest('tr');
        if (row) {
          const radio = row.querySelector('input[type="radio"], .el-radio__inner');
          if (radio) radio.click();
        }
        cell.click();
        return true;
      }
    }
    return false;
  }, info.문건분류선택);

  await page.waitForTimeout(800);

  if (clicked) {
    // 선택하기 버튼 클릭
    await page.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      const btn = btns.find(b => b.textContent.includes('선택하기'));
      if (btn) btn.click();
    });
    await page.waitForTimeout(800);
    // 확인 팝업 처리
    await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '확인');
      if (btn) btn.click();
    });
    await page.waitForTimeout(500);
    log('문건분류 선택 완료!');
  } else {
    warn(`문건분류 항목을 찾지 못했습니다: ${info.문건분류선택}`);
    await page.evaluate(() => {
      const btn = document.querySelector('.el-dialog__headerbtn');
      if (btn) btn.click();
    });
    await page.waitForTimeout(500);
  }

  if (info.시행일) {
    const siField = page.locator('input[placeholder="시행일"]').first();
    if (await siField.count() > 0) {
      await siField.click();
      await siField.fill(info.시행일);
      await page.keyboard.press('Escape');
    }
  }
}

async function selectDropdown(page, placeholder, value) {
  await page.locator(`input[placeholder="${placeholder}"]`).first().click();
  await page.waitForTimeout(600);
  await page.evaluate((val) => {
    const items = document.querySelectorAll('.el-select-dropdown:not([style*="display: none"]) .el-select-dropdown__item');
    const item = Array.from(items).find(el => el.textContent.includes(val));
    if (item) item.click();
  }, value);
  await page.waitForTimeout(300);
  await page.keyboard.press('Escape');
  await page.waitForTimeout(300);
}

async function fillBudgetTable(page, rows) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.includes('예산 연동'));
      if (btn) btn.click();
    });
    await page.waitForTimeout(1000);
    await page.waitForSelector('input[placeholder="코스트센터 선택"]', { timeout: 10000 });
    await page.waitForTimeout(500);

    // 유형선택
    await page.evaluate((type) => {
      const btns = Array.from(document.querySelectorAll('.el-dialog button, .el-dialog .el-radio-button__inner'));
      const btn = btns.find(b => b.textContent.trim() === type);
      if (btn) btn.click();
    }, row.구분);
    await page.waitForTimeout(300);

    await selectDropdown(page, '코스트센터 선택', row.from.코스트센터);
    await selectDropdown(page, '계정코드 선택', row.from.GL계정);

    await page.locator('input[placeholder="YYYYMM"]').first().fill(String(row.from.신청월));
    await page.waitForTimeout(300);

    // 잔액조회 버튼 클릭
    await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('.el-dialog button')).find(b => b.textContent.includes('잔액조회'));
      if (btn) btn.click();
    });

    // 잔액이 실제로 로드될 때까지 대기 (최대 10초)
    let balance = 0;
    for (let w = 0; w < 10; w++) {
      await page.waitForTimeout(1000);
      balance = await page.evaluate(() => {
        const inputs = document.querySelectorAll('.el-dialog input[placeholder="0"], .el-dialog input[readonly]');
        for (const input of inputs) {
          const val = parseInt(input.value.replace(/,/g, ''));
          if (val > 0) return val;
        }
        return 0;
      });
      if (balance > 0) break;
    }
    log(`잔액 확인: ${balance.toLocaleString()}원`);

    // 신청예산이 잔액 초과인지 확인
    if (balance > 0 && row.from.금액 > balance) {
      warn(`신청예산(${row.from.금액.toLocaleString()}원)이 잔액(${balance.toLocaleString()}원)을 초과합니다. 건너뜁니다.`);
      await page.evaluate(() => {
        const btn = Array.from(document.querySelectorAll('.el-dialog button')).find(b => b.textContent.includes('닫기'));
        if (btn) btn.click();
      });
      await page.waitForTimeout(1000);
      continue;
    }

    // 금액 입력 후 즉시 추가하기 클릭 (잔액 리셋 방지)
    const addResult = await page.evaluate((amt) => {
      const amtInput = document.querySelector('input[placeholder="금액입력"]');
      if (!amtInput) return false;

      // Vue 네이티브 setter로 값 설정
      const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      nativeSetter.call(amtInput, String(amt));
      amtInput.dispatchEvent(new Event('input', { bubbles: true }));
      amtInput.dispatchEvent(new Event('change', { bubbles: true }));

      // 즉시 추가하기 클릭
      setTimeout(() => {
        const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.includes('추가하기'));
        if (btn) btn.click();
      }, 200);

      return true;
    }, row.from.금액);

    await page.waitForTimeout(3000);
    if (!addResult) warn('금액 입력 또는 추가하기 실패');

    log(`예산 연동 ${i + 1}/${rows.length} 완료 (${row.from.신청월}, ${row.from.금액.toLocaleString()}원)`);
  }
}

async function fillPurchaseForm(page, form) {
  const reqLabel = page.locator(`label:has-text("${form.요청유형}")`).first();
  if (await reqLabel.count() > 0) await reqLabel.click();

  const contractLabel = page.locator(`label:has-text("${form.계약유무}")`).first();
  if (await contractLabel.count() > 0) await contractLabel.click();

  const supplierField = page.locator('input[placeholder*="공급사"], input[placeholder*="거래처"]').first();
  if (await supplierField.count() > 0) await supplierField.fill(form.공급사);

  if (form.계약기간) {
    const periodField = page.locator('input[placeholder*="계약기간"]').first();
    if (await periodField.count() > 0) await periodField.fill(form.계약기간);
  }

  for (let i = 0; i < form.품목.length; i++) {
    const item = form.품목[i];
    if (i > 0) {
      const addBtn = page.locator('button:has-text("행 추가"), button:has-text("+행")').first();
      if (await addBtn.count() > 0) {
        await addBtn.click();
        await page.waitForTimeout(300);
      }
    }
    const row   = page.locator('table tbody tr').nth(i);
    const cells = row.locator('input');
    if (item.품목명         && await cells.count() > 0) await cells.nth(0).fill(item.품목명);
    if (item.수량 != null   && await cells.count() > 1) await cells.nth(1).fill(String(item.수량));
    if (item.단위           && await cells.count() > 2) await cells.nth(2).fill(item.단위);
    if (item.단가 != null   && await cells.count() > 3) await cells.nth(3).fill(String(item.단가));
    if (                       await cells.count() > 4) await cells.nth(4).fill(String(item.비용));
  }
}

async function fillBody(page, bodyText) {
  // 페이지 맨 아래로 스크롤
  await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
  await page.waitForTimeout(1000);

  // Summernote 에디터 클릭 후 입력
  const success = await page.evaluate((text) => {
    const editor = document.querySelector('.note-editable');
    if (editor) {
      editor.focus();
      editor.click();
      editor.innerHTML = text.replace(/\n/g, '<br>');
      editor.dispatchEvent(new Event('input', { bubbles: true }));
      editor.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    }
    return false;
  }, bodyText);

  if (success) {
    await page.waitForTimeout(500);
    log('본문 입력 완료!');
  } else {
    warn('본문 입력 필드를 찾지 못했습니다. 수동으로 입력해주세요.');
  }
}

async function attachFiles(page, files) {
  if (!files || files.length === 0) return;

  for (const filePath of files) {
    log(`첨부파일 업로드 중: ${filePath}`);
    const fileInput = page.locator('input[type="file"]').first();
    if (await fileInput.count() > 0) {
      await fileInput.setInputFiles(filePath);
      await page.waitForTimeout(1500);
      log(`첨부파일 완료: ${filePath}`);
    } else {
      // 파일 등록 버튼 클릭
      const fileBtn = page.locator('button:has-text("파일등록"), button:has-text("파일 등록"), a:has-text("파일등록")').first();
      if (await fileBtn.count() > 0) {
        await fileBtn.click();
        await page.waitForTimeout(500);
        const input = page.locator('input[type="file"]').first();
        if (await input.count() > 0) {
          await input.setInputFiles(filePath);
          await page.waitForTimeout(1500);
        }
      } else {
        warn(`파일 입력 필드를 찾지 못했습니다. 수동으로 첨부해주세요: ${filePath}`);
      }
    }
  }
}

async function submitDoc(page, docType) {
  page.on('dialog', async (dialog) => {
    log(`팝업: "${dialog.message()}" → 확인`);
    await dialog.accept();
  });
  const submitBtn = page.locator('button:has-text("문건등록"), button:has-text("상신")').first();
  if (await submitBtn.count() === 0) {
    warn('제출 버튼을 찾지 못했습니다. 수동으로 제출해주세요.');
    await page.waitForTimeout(10000);
    return;
  }
  await submitBtn.click();
  await page.waitForTimeout(3000);

  // 제출 후 문건번호 저장 (예산품의인 경우)
  if (docType && docType.endsWith('-b')) {
    try {
      const url = page.url();
      const match = url.match(/docNo=([0-9]+)/);
      if (match) {
        saveDocNo(docType, match[1]);
      } else {
        // URL에서 못 찾으면 페이지에서 찾기
        const docNo = await page.evaluate(() => {
          const el = document.querySelector('[class*="docNo"], [id*="docNo"]');
          return el ? el.textContent.trim() : null;
        });
        if (docNo) saveDocNo(docType, docNo);
      }
    } catch(e) {}
  }
}

main();
