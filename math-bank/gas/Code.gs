/**
 * 辰光國小：數學英雄傳 v3
 * - 學生：選班級＋座號，姓名從「學生名單」自動帶出；每題作答後立即回饋
 * - 題庫：南一版 1上～6下，每題標註「概念」與「迷思概念」
 * - 紀錄：Log_登入時間（每次登入）、Log_登入紀錄（每次測驗分數）、Log_作答明細（每題＋迷思）
 * - 教師後台：網址後加 ?page=admin
 */
const SPREADSHEET_ID = '1FopsShn-X1xyGq-3PcimhUVGNHWfbLfWWm0lwhk2G24';
const ADMIN_PASSWORD = '0000';            // 🔑 請改成自己的後台密碼
const BANK_URL = 'https://raw.githubusercontent.com/cyysongyy/Young-App-Library/main/math-bank/questions.json';
const GRADE_SHEETS = ['1上','1下','2上','2下','3上','3下','4上','4下','5上','5下','6上','6下'];
const ROSTER_SHEET = '學生名單';
const LOGIN_SHEET = 'Log_登入時間';
const LOG_SHEET = 'Log_登入紀錄';      // 沿用舊名：每次測驗一列
const DETAIL_SHEET = 'Log_作答明細';
const HEADER = ['單元名稱','題號','題型','題目內容','選項/提示','正確答案','解析/插圖建議','概念','迷思概念'];
const TZ = 'Asia/Taipei';

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
    .addItem('👩‍🎓 建立學生名單分頁', 'ensureRoster')
    .addToUi();
}

const ss_ = () => SpreadsheetApp.openById(SPREADSHEET_ID);
const now_ = () => new Date();

// ───────── 學生名單與登入 ─────────
function ensureRoster() {
  const ss = ss_();
  let sh = ss.getSheetByName(ROSTER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ROSTER_SHEET, 0);
    sh.appendRow(['班級', '座號', '姓名']);
    sh.getRange('A1:C1').setFontWeight('bold').setBackground('#28a745').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('0');
  }
  return sh;
}

/** 回傳 { '三甲': { '1': '王小明', ... }, ... } */
function getRoster() {
  const sh = ss_().getSheetByName(ROSTER_SHEET);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
    const c = String(r[0]).trim(), s = String(r[1]).trim(), n = String(r[2]).trim();
    if (c && s) (out[c] = out[c] || {})[s] = n;
  });
  return out;
}

function logLogin(p) {
  const ss = ss_();
  let sh = ss.getSheetByName(LOGIN_SHEET);
  if (!sh) { sh = ss.insertSheet(LOGIN_SHEET); sh.appendRow(['登入時間', '班級', '座號', '姓名', '裝置']); sh.setFrozenRows(1); }
  sh.appendRow([now_(), p.gradeClass, p.seatNum, p.name, String(p.device || '').slice(0, 60)]);
  return true;
}

// ───────── 題庫讀取 ─────────
function getUnitList(grade) {
  const sheet = ss_().getSheetByName(grade);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
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
        answer: String(data[i][5]).trim(), desc: clean_(data[i][6]),
        concept: String(data[i][7] || '').trim(), mis: String(data[i][8] || '').trim(),
        strict: /最簡/.test(String(data[i][3]))
      });
    }
  }
  return out;
}

const clean_ = t => String(t == null ? '' : t)
  .replace(/\\div/g, ' ÷ ').replace(/\\times/g, ' × ').replace(/\\circ/g, '°')
  .replace(/\$/g, '').replace(/\^/g, '').replace(/\\/g, '');

/** 題目＋答案＋解析＋迷思對照（前端即時回饋用） */
function getQuestions(grade, unitName) {
  return readUnit_(grade, unitName).map(q => {
    const choices = parseChoices_(q.options);
    return { id: q.id, type: q.type, content: q.content, choices: choices,
             hint: choices.length ? '' : q.options,
             answer: q.answer, shown: shownAnswer_(q, choices), desc: q.desc,
             concept: q.concept, mis: q.mis, strict: q.strict };
  });
}

function parseChoices_(s) {
  const out = [];
  const re = /\(([A-D])\)\s*([\s\S]*?)(?=\s*\([A-D]\)|$)/g;
  let m;
  while ((m = re.exec(String(s))) !== null) out.push({ key: m[1], text: m[2].trim() });
  return out;
}
function shownAnswer_(q, choices) {
  if (q.type !== '選擇') return q.answer;
  const ch = (choices || parseChoices_(q.options)).find(c => c.key === q.answer.toUpperCase());
  return ch ? `(${ch.key}) ${ch.text}` : q.answer;
}

// ───────── 判分與迷思診斷（前端 Index.html 有同一份） ─────────
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
function isCorrect_(user, correct, type, strict) {
  const u = norm_(user), c = norm_(correct);
  if (!u) return false;
  if (type === '選擇') return u.replace(/[^a-d]/g, '').charAt(0) === c.replace(/[^a-d]/g, '').charAt(0);
  if (u === c) return true;
  if (strict) {   // 題目要求「最簡分數」：不接受未約分的等值分數，只放寬單位
    const cu = splitUnit_(c), uu = splitUnit_(u);
    return !!(cu && ((uu && uu.unit === '' && uu.num === cu.num) || u === cu.num));
  }
  if (numEq_(u, c)) return true;
  const cu = splitUnit_(c), uu = splitUnit_(u);
  if (cu && uu && (uu.unit === '' || uu.unit === cu.unit)) return numEq_(uu.num, cu.num);
  if (cu && !uu && numEq_(u, cu.num)) return true;
  return false;
}
/** 迷思對照字串 "B=…;C,D=…;*=…" → 學生答錯時的迷思 */
function diagnose_(user, q) {
  if (!String(user).trim()) return '未作答';
  let fallback = '';
  for (const part of String(q.mis || '').split(';')) {
    const i = part.indexOf('='); if (i < 0) continue;
    const keys = part.slice(0, i).split(','), label = part.slice(i + 1).trim();
    for (let k of keys) {
      k = k.trim();
      if (k === '*') { fallback = label; continue; }
      if (q.type === '選擇' ? isCorrect_(user, k, '選擇') : isCorrect_(user, k, '填充', q.strict)) return label;
    }
  }
  return fallback || (q.concept ? q.concept + '概念不清' : '其他錯誤');
}

/** 學生完成單元：後端重新判分、寫紀錄 */
function submitAnswers(p) {
  const qs = readUnit_(p.grade, p.unit);
  const n = qs.length || 1;
  const results = qs.map((q, i) => {
    const ua = (p.answers || [])[i] || '';
    const ok = isCorrect_(ua, q.answer, q.type, q.strict);
    return { no: i + 1, type: q.type, user: ua, ok: ok, answer: shownAnswer_(q), desc: q.desc,
             concept: q.concept, mis: ok ? '' : diagnose_(ua, q) };
  });
  const right = results.filter(r => r.ok).length;
  const score = Math.round(right / n * 100);
  if (!p.practice) {
    const t = now_();
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      ensureLogSheet_().appendRow([t, p.name, p.gradeClass, p.seatNum, '😊', p.unit, score, p.duration, p.grade, p.seconds]);
      const det = ensureDetailSheet_();
      const rows = results.map(r => [t, p.name, p.gradeClass, p.seatNum, p.grade, p.unit, r.no, r.type,
        String(r.user), String(qs[r.no - 1].answer), r.ok ? 1 : 0, p.seconds, r.concept, r.mis]);
      if (rows.length) {
        const r0 = det.getLastRow() + 1;
        det.getRange(r0, 9, rows.length, 2).setNumberFormat('@');   // 答案欄當文字，避免 3/4 變日期
        det.getRange(r0, 1, rows.length, rows[0].length).setValues(rows);
      }
    } finally { lock.releaseLock(); }
  }
  return { score: score, right: right, total: n, results: results };
}

/** 舊版相容 */
function saveLog(d) { ensureLogSheet_().appendRow([now_(), d.name, d.gradeClass, d.seatNum, '😊', d.unit, d.score, d.duration]); }

function ensureLogSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) { sh = ss.insertSheet(LOG_SHEET); sh.appendRow(['測驗時間','學生姓名','班級','座號','狀態','測驗單元','分數','考試時間','年級分頁','用時秒數']); }
  if (sh.getRange(1, 9).getValue() === '') sh.getRange(1, 9, 1, 2).setValues([['年級分頁','用時秒數']]);
  return sh;
}
function ensureDetailSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(DETAIL_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DETAIL_SHEET);
    sh.appendRow(['作答時間','學生姓名','班級','座號','年級分頁','測驗單元','題號','題型','學生答案','正確答案','是否答對','用時秒數','概念','迷思概念']);
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, 13).getValue() === '') {
    sh.getRange(1, 13, 1, 2).setValues([['概念', '迷思概念']]);
  }
  return sh;
}

// ───────── 教師後台 ─────────
function getDashboard(pw) {
  if (String(pw) !== String(ADMIN_PASSWORD)) throw new Error('密碼錯誤');
  const ss = ss_();
  const fmt = d => d instanceof Date ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm') : String(d);
  const read = (name, cols) => {
    const sh = ss.getSheetByName(name);
    return sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues() : [];
  };
  // logs: [time,name,class,seat,unit,score,grade,secs,duration]
  const logs = read(LOG_SHEET, 10).filter(r => r[0] !== '' && r[1] !== '')
    .map(r => [fmt(r[0]), String(r[1]), String(r[2]), String(r[3]), String(r[5]), Number(r[6]) || 0, String(r[8] || ''), Number(r[9]) || 0, String(r[7] || '')]);
  // details: [time,name,class,seat,grade,unit,no,type,user,correct,ok,concept,mis]
  const details = read(DETAIL_SHEET, 14).filter(r => r[0] !== '')
    .map(r => [fmt(r[0]), String(r[1]), String(r[2]), String(r[3]), String(r[4]), String(r[5]), Number(r[6]), String(r[7]),
               String(r[8]), String(r[9]), Number(r[10]), String(r[12] || ''), String(r[13] || '')]);
  // logins: [time,class,seat,name]
  const logins = read(LOGIN_SHEET, 4).filter(r => r[0] !== '').map(r => [fmt(r[0]), String(r[1]), String(r[2]), String(r[3])]);
  const qText = {};
  GRADE_SHEETS.forEach(g => {
    const sh = ss.getSheetByName(g); if (!sh) return;
    let unit = '';
    sh.getDataRange().getValues().slice(1).forEach(r => {
      if (String(r[0]).trim()) unit = String(r[0]).trim();
      if (unit && r[1] !== '') qText[g + '|' + unit + '|' + r[1]] = clean_(r[3]).slice(0, 80);
    });
  });
  return { logs, details, logins, qText, roster: getRoster(), updated: fmt(now_()), appUrl: ScriptApp.getService().getUrl() };
}

// ───────── 匯入題庫（在試算表選單或編輯器執行） ─────────
function importBank() {
  const res = UrlFetchApp.fetch(BANK_URL, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('下載題庫失敗：' + res.getResponseCode() + '，請檢查 BANK_URL');
  const bank = JSON.parse(res.getContentText());
  const ss = ss_();
  const stamp = Utilities.formatDate(now_(), TZ, 'MMdd_HHmm');
  GRADE_SHEETS.forEach(g => {
    const rows = bank.sheets[g]; if (!rows) return;
    let sh = ss.getSheetByName(g);
    if (sh) { sh.copyTo(ss).setName(g + '_備份' + stamp).hideSheet(); sh.clear(); }
    else sh = ss.insertSheet(g);
    const all = [HEADER].concat(rows);
    // 先設成純文字再寫入，避免 3/4、1/2 被自動轉成日期
    sh.getRange(1, 1, all.length, HEADER.length).setNumberFormat('@')
      .setValues(all.map(r => HEADER.map((_, k) => r[k] == null ? '' : String(r[k]))));
    sh.getRange(2, 2, rows.length, 1).setNumberFormat('0').setValues(rows.map(r => [Number(r[1])]));
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold').setBackground('#28a745').setFontColor('#ffffff');
  });
  ensureRoster(); ensureDetailSheet_();
  try { SpreadsheetApp.getUi().alert('✅ 匯入完成：' + bank.version + '\n舊題目已備份為隱藏分頁（名稱含「_備份」）。\n請到「學生名單」分頁填入班級、座號、姓名。'); } catch (e) {}
}
