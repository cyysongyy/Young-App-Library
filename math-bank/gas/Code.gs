/**
 * 辰光國小：數學英雄傳（v2）
 * - 題庫：南一版 1上～6下，每單元 10 題（5 選擇、3 填充、2 應用）
 * - 判分移到後端：全形/半形、空格、單位、等值分數都能判對
 * - 每題作答明細寫入「Log_作答明細」，供分析後台使用
 * - 教師後台：網址後加 ?page=admin
 */
const SPREADSHEET_ID = '1FopsShn-X1xyGq-3PcimhUVGNHWfbLfWWm0lwhk2G24';
const ADMIN_PASSWORD = '0000';            // 🔑 請改成自己的後台密碼
const BANK_URL = 'https://raw.githubusercontent.com/cyysongyy/young-app-library/main/math-bank/questions.json'; // 匯入題庫用
const GRADE_SHEETS = ['1上','1下','2上','2下','3上','3下','4上','4下','5上','5下','6上','6下'];
const LOG_SHEET = 'Log_登入紀錄';
const DETAIL_SHEET = 'Log_作答明細';
const HEADER = ['單元名稱','題號','題型','題目內容','選項/提示','正確答案','解析/插圖建議'];

// ───────── 網頁入口 ─────────
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) === 'admin' ? 'Admin' : 'Index';
  return HtmlService.createTemplateFromFile(page)
    .evaluate()
    .setTitle(page === 'Admin' ? '數學英雄傳｜教師分析後台' : '辰光國小：數學英雄傳')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 數學英雄傳')
    .addItem('📥 從 GitHub 匯入南一題庫（舊題自動備份）', 'importBank')
    .addItem('🧹 建立／整理作答明細表', 'ensureDetailSheet')
    .addToUi();
}

const ss_ = () => SpreadsheetApp.openById(SPREADSHEET_ID);

// ───────── 題庫讀取 ─────────
function getUnitList(grade) {
  const sheet = ss_().getSheetByName(grade);
  if (!sheet) return [];
  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
  return [...new Set(data.map(r => String(r[0]).trim()).filter(Boolean))];
}

function readUnit_(grade, unitName) {
  const sheet = ss_().getSheetByName(grade);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const out = [];
  let tracking = false;
  for (let i = 1; i < data.length; i++) {
    const a = String(data[i][0]).trim();
    if (a === String(unitName).trim()) tracking = true;
    else if (a !== '' && tracking) break;
    if (tracking && String(data[i][3]).trim() !== '') {
      out.push({
        id: data[i][1], type: String(data[i][2]).trim(),
        content: clean_(data[i][3]), options: clean_(data[i][4]),
        answer: String(data[i][5]).trim(), desc: clean_(data[i][6])
      });
    }
  }
  return out;
}

const clean_ = t => String(t)
  .replace(/\\div/g, ' ÷ ').replace(/\\times/g, ' × ').replace(/\\circ/g, '°')
  .replace(/\$/g, '').replace(/\^/g, '').replace(/\\/g, '');

/** 給學生的題目：不含答案與解析 */
function getQuestions(grade, unitName) {
  return readUnit_(grade, unitName).map(q => {
    const choices = parseChoices_(q.options);
    return { id: q.id, type: q.type, content: q.content,
             choices: choices,                       // 選擇題：[{key:'A',text:'..'}]
             hint: choices.length ? '' : q.options }; // 填充/應用：作答格式提示
  });
}

function parseChoices_(s) {
  const out = [];
  const re = /\(([A-D])\)\s*([\s\S]*?)(?=\s*\([A-D]\)|$)/g;
  let m;
  while ((m = re.exec(String(s))) !== null) out.push({ key: m[1], text: m[2].trim() });
  return out;
}

// ───────── 判分 ─────────
function toHalf_(s) {
  return String(s == null ? '' : s)
    .replace(/[！-～]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/　/g, ' ').replace(/[，、]/g, ',').replace(/[。]/g, '');
}
function norm_(s) { return toHalf_(s).toLowerCase().replace(/\s+/g, '').replace(/^\((.*)\)$/, '$1'); }
function toNum_(s) {
  s = norm_(s);
  let m;
  if ((m = s.match(/^(\d+)又(\d+)\/(\d+)$/))) return +m[1] + m[2] / m[3];
  if ((m = s.match(/^(\d+)\/(\d+)$/))) return m[2] == 0 ? NaN : m[1] / m[2];
  if (/^-?\d+(\.\d+)?$/.test(s)) return parseFloat(s);
  return NaN;
}
function splitUnit_(s) {
  const m = norm_(s).match(/^(\d+又\d+\/\d+|\d+\/\d+|\d+(?:\.\d+)?)(.*)$/);
  return m && !/\d/.test(m[2]) ? { num: m[1], unit: m[2] } : null;
}
function numEq_(a, b) { const x = toNum_(a), y = toNum_(b); return !isNaN(x) && !isNaN(y) && Math.abs(x - y) < 1e-9; }

function isCorrect_(user, correct, type) {
  const u = norm_(user), c = norm_(correct);
  if (!u) return false;
  if (type === '選擇') return u.replace(/[^a-d]/g, '').charAt(0) === c.replace(/[^a-d]/g, '').charAt(0);
  if (u === c) return true;
  if (numEq_(u, c)) return true;
  const cu = splitUnit_(c), uu = splitUnit_(u);
  if (cu && uu && (uu.unit === '' || uu.unit === cu.unit)) return numEq_(uu.num, cu.num);
  if (cu && !uu && numEq_(u, cu.num)) return true;
  return false;
}

/** 學生交卷：後端判分、寫紀錄、回傳結果 */
function submitAnswers(p) {
  const qs = readUnit_(p.grade, p.unit);
  const n = qs.length || 1;
  const results = qs.map((q, i) => {
    const ua = (p.answers || [])[i] || '';
    const ok = isCorrect_(ua, q.answer, q.type);
    let shown = q.answer;
    if (q.type === '選擇') {
      const ch = parseChoices_(q.options).find(c => c.key === q.answer.toUpperCase());
      if (ch) shown = `(${ch.key}) ${ch.text}`;
    }
    return { no: i + 1, type: q.type, user: ua, ok: ok, answer: shown, desc: q.desc };
  });
  const right = results.filter(r => r.ok).length;
  const score = Math.round(right / n * 100);
  if (!p.practice) {
    const now = new Date();
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      ensureLogSheet_().appendRow([now, p.name, p.gradeClass, p.seatNum, '😊', p.unit, score, p.duration, p.grade, p.seconds]);
      const det = ensureDetailSheet();
      const rows = results.map(r => [now, p.name, p.gradeClass, p.seatNum, p.grade, p.unit, r.no, r.type, r.user, qs[r.no - 1].answer, r.ok ? 1 : 0, p.seconds]);
      if (rows.length) {
        const r0 = det.getLastRow() + 1;
        det.getRange(r0, 9, rows.length, 2).setNumberFormat('@');   // 答案欄當文字，避免 3/4 變日期
        det.getRange(r0, 1, rows.length, rows[0].length).setValues(rows.map(r => r.map((v, k) => k === 8 || k === 9 ? String(v) : v)));
      }
    } finally { lock.releaseLock(); }
  }
  return { score: score, right: right, total: n, results: results };
}

/** 舊版相容：保留原本的 saveLog */
function saveLog(d) {
  ensureLogSheet_().appendRow([new Date(), d.name, d.gradeClass, d.seatNum, '😊', d.unit, d.score, d.duration]);
}

function ensureLogSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) { sh = ss.insertSheet(LOG_SHEET); sh.appendRow(['登入時間','學生姓名','班級','座號','狀態','測驗單元','分數','考試時間','年級分頁','用時秒數']); }
  if (sh.getRange(1, 9).getValue() === '') sh.getRange(1, 9, 1, 2).setValues([['年級分頁','用時秒數']]);
  return sh;
}
function ensureDetailSheet() {
  const ss = ss_();
  let sh = ss.getSheetByName(DETAIL_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DETAIL_SHEET);
    sh.appendRow(['作答時間','學生姓名','班級','座號','年級分頁','測驗單元','題號','題型','學生答案','正確答案','是否答對','用時秒數']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ───────── 教師後台 ─────────
function getDashboard(pw) {
  if (String(pw) !== String(ADMIN_PASSWORD)) throw new Error('密碼錯誤');
  const ss = ss_();
  const fmt = d => d instanceof Date ? Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : String(d);
  const logSh = ss.getSheetByName(LOG_SHEET);
  const logs = logSh && logSh.getLastRow() > 1
    ? logSh.getRange(2, 1, logSh.getLastRow() - 1, 10).getValues()
        .filter(r => r[0] !== '' && r[1] !== '')
        .map(r => [fmt(r[0]), String(r[1]), String(r[2]), String(r[3]), String(r[5]), Number(r[6]) || 0, String(r[8] || ''), Number(r[9]) || 0])
    : [];
  const detSh = ss.getSheetByName(DETAIL_SHEET);
  const details = detSh && detSh.getLastRow() > 1
    ? detSh.getRange(2, 1, detSh.getLastRow() - 1, 11).getValues()
        .map(r => [fmt(r[0]), String(r[1]), String(r[2]), String(r[3]), String(r[4]), String(r[5]), Number(r[6]), String(r[7]), String(r[8]), String(r[9]), Number(r[10])])
    : [];
  // 題目文字（給「最常錯的題目」顯示用）
  const qText = {};
  GRADE_SHEETS.forEach(g => {
    const sh = ss.getSheetByName(g); if (!sh) return;
    let unit = '';
    sh.getDataRange().getValues().slice(1).forEach(r => {
      if (String(r[0]).trim()) unit = String(r[0]).trim();
      if (unit && r[1] !== '') qText[g + '|' + unit + '|' + r[1]] = clean_(r[3]).slice(0, 80);
    });
  });
  return { logs: logs, details: details, qText: qText, updated: fmt(new Date()) };
}

// ───────── 匯入題庫（在試算表選單執行） ─────────
function importBank() {
  const res = UrlFetchApp.fetch(BANK_URL, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('下載題庫失敗：' + res.getResponseCode() + '，請檢查 BANK_URL');
  const bank = JSON.parse(res.getContentText());
  const ss = ss_();
  const stamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'MMdd_HHmm');
  GRADE_SHEETS.forEach(g => {
    const rows = bank.sheets[g]; if (!rows) return;
    let sh = ss.getSheetByName(g);
    if (sh) { sh.copyTo(ss).setName(g + '_備份' + stamp).hideSheet(); sh.clear(); }
    else sh = ss.insertSheet(g);
    const all = [HEADER].concat(rows);
    // 先設成純文字格式，再寫入，避免 3/4、1/2 被試算表自動轉成日期
    sh.getRange(1, 1, all.length, HEADER.length).setNumberFormat('@')
      .setValues(all.map(r => r.map(v => v === null ? '' : String(v))));
    sh.getRange(2, 2, rows.length, 1).setNumberFormat('0').setValues(rows.map(r => [Number(r[1])]));
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold').setBackground('#28a745').setFontColor('#ffffff');
  });
  try { SpreadsheetApp.getUi().alert('✅ 匯入完成：' + bank.version + '\n舊題目已備份為隱藏分頁（名稱含「_備份」）。'); } catch (e) {}
}
