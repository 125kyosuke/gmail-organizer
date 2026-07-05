/**
 * ✉️ Gmail整理シート (Google Apps Script / スプレッドシート連動版)
 * ------------------------------------------------------------------
 * コード編集は不要です。設定はぜんぶスプレッドシートのセルで行い、
 * 実行はメニュー「✉️ Gmail整理」から選ぶだけ。
 *
 * 初回のみ: メニュー「✉️ Gmail整理 > ① 初期セットアップ」を実行すると
 * 「設定」「プレビュー」「実行ログ」シートが自動で作られます。
 * ================================================================== */

const SHEET_SETTINGS = '設定';
const SHEET_PREVIEW = 'プレビュー';
const SHEET_LOG = '実行ログ';
const MAX_THREADS = 200; // 1カテゴリあたりの処理上限(実行時間対策)

/* ────────────────────────────────────────────────
 *  メニュー
 * ──────────────────────────────────────────────── */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('✉️ Gmail整理')
    .addItem('① 初期セットアップ(最初に1回)', 'setup')
    .addSeparator()
    .addItem('② お試し実行(見るだけ・変更なし)', 'previewRun')
    .addItem('③ 本番実行(実際に整理する)', 'realRun')
    .addSeparator()
    .addItem('④ 毎日の自動整理を ON', 'enableDaily')
    .addItem('　 毎日の自動整理を OFF', 'disableDaily')
    .addToUi();
}

/* ────────────────────────────────────────────────
 *  ① 初期セットアップ — 設定シートを作る
 * ──────────────────────────────────────────────── */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 設定シート
  let s = ss.getSheetByName(SHEET_SETTINGS);
  if (!s) {
    s = ss.insertSheet(SHEET_SETTINGS, 0);
    const rows = [
      ['✉️ Gmail整理の設定', '', ''],
      ['', '', ''],
      ['■ 重要メール(★を付けて受信トレイに残す)', '', ''],
      ['見逃したくない送信元(1行に1つ。@ドメインでもOK)', '@aiu.ac.jp', ''],
      ['', '', ''],
      ['', '', ''],
      ['', '', ''],
      ['重要あつかいにする件名キーワード(1行に1つ)', '締切', ''],
      ['', '請求', ''],
      ['', '至急', ''],
      ['', '面接', ''],
      ['', '', ''],
      ['■ 自動仕分け(ラベルを付けて受信トレイから外す)', '', 'ON/OFF'],
      ['メルマガ・宣伝を仕分ける', '届いてから2日後', 'ON'],
      ['SNS通知を仕分ける', '届いてから2日後', 'ON'],
      ['各種通知・お知らせを仕分ける', '届いてから2日後', 'ON'],
      ['フォーラム・メーリングリストを仕分ける', '届いてから2日後', 'ON'],
      ['', '', ''],
      ['■ おそうじ', '', ''],
      ['読み終わって30日たったメールを受信トレイから外す', '', 'ON'],
      ['', '', ''],
      ['※ 削除は一切しません。「外す」=アーカイブで、検索すれば残っています。', '', ''],
      ['※ ★付きメールと「★重要」ラベルのメールは絶対に動かしません。', '', ''],
    ];
    s.getRange(1, 1, rows.length, 3).setValues(rows);
    // 見た目
    s.setColumnWidth(1, 380).setColumnWidth(2, 220).setColumnWidth(3, 80);
    s.getRange('A1').setFontSize(14).setFontWeight('bold');
    [3, 13, 19].forEach((r) =>
      s.getRange(r, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe')
    );
    // ON/OFF はプルダウンに
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['ON', 'OFF'], true).build();
    s.getRange('C14:C17').setDataValidation(rule);
    s.getRange('C20').setDataValidation(rule);
    s.getRange('B4:B11').setBackground('#fffde7'); // 入力欄を色付け
  }

  getOrCreatePreviewSheet_();
  getOrCreateLogSheet_();
  ss.setActiveSheet(ss.getSheetByName(SHEET_SETTINGS));

  SpreadsheetApp.getUi().alert(
    'セットアップ完了!\n\n' +
    '「設定」シートの黄色いセルに、見逃したくない送信元やキーワードを書いてください。\n\n' +
    'そのあとメニューから「② お試し実行」をどうぞ。メールは一切変更されず、' +
    '何が整理されるかが「プレビュー」シートに表示されます。'
  );
}

/* ────────────────────────────────────────────────
 *  設定シートを読む
 * ──────────────────────────────────────────────── */
function readSettings_() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!s) throw new Error('先に「① 初期セットアップ」を実行してください。');

  const colB = s.getRange('B1:B24').getValues().map((r) => String(r[0]).trim());
  const colC = s.getRange('C1:C24').getValues().map((r) => String(r[0]).trim());

  const senders = colB.slice(3, 7).filter(Boolean);   // B4:B7
  const keywords = colB.slice(7, 11).filter(Boolean); // B8:B11

  return {
    senders,
    keywords,
    rules: [
      { on: colC[13] !== 'OFF', name: 'メルマガ・宣伝', cat: 'promotions', label: '整理/メルマガ' },
      { on: colC[14] !== 'OFF', name: 'SNS通知',        cat: 'social',     label: '整理/SNS' },
      { on: colC[15] !== 'OFF', name: '通知・お知らせ', cat: 'updates',    label: '整理/通知' },
      { on: colC[16] !== 'OFF', name: 'フォーラム',     cat: 'forums',     label: '整理/フォーラム' },
    ],
    cleanupOn: colC[19] !== 'OFF',
  };
}

/* ────────────────────────────────────────────────
 *  ② お試し実行 / ③ 本番実行
 * ──────────────────────────────────────────────── */
function previewRun() { organize_(true); }

function realRun() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    '本番実行',
    '実際にメールを整理します(★付け・ラベル付け・受信トレイから外す)。\n' +
    '削除はしません。よろしいですか?',
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;
  organize_(false);
}

function organize_(dryRun) {
  const cfg = readSettings_();
  const out = []; // [種別, 差出人, 件名, 日付, アクション]

  // ── 1. 重要メール ──
  const importantLabel = getOrCreateLabel_('★重要');
  const clauses = [];
  if (cfg.senders.length) clauses.push('(' + cfg.senders.map((x) => `from:"${x}"`).join(' OR ') + ')');
  if (cfg.keywords.length) clauses.push('(' + cfg.keywords.map((x) => `subject:"${x}"`).join(' OR ') + ')');
  if (clauses.length) {
    const q = `in:inbox -label:★重要 (${clauses.join(' OR ')})`;
    for (const t of GmailApp.search(q, 0, MAX_THREADS)) {
      out.push(['重要', senderOf_(t), subjectOf_(t), dateOf_(t), '★を付けてトレイに残す']);
      if (!dryRun) {
        t.addLabel(importantLabel);
        const m = t.getMessages();
        if (m.length) m[0].star();
      }
    }
  }

  // ── 2. カテゴリ仕分け ──
  for (const r of cfg.rules) {
    if (!r.on) continue;
    const q = `in:inbox category:${r.cat} older_than:2d -is:starred -label:★重要`;
    const label = getOrCreateLabel_(r.label);
    for (const t of GmailApp.search(q, 0, MAX_THREADS)) {
      out.push([r.name, senderOf_(t), subjectOf_(t), dateOf_(t), `「${r.label}」を付けてトレイから外す`]);
      if (!dryRun) {
        t.addLabel(label);
        t.moveToArchive();
      }
    }
  }

  // ── 3. おそうじ ──
  if (cfg.cleanupOn) {
    const q = 'in:inbox is:read older_than:30d -is:starred -label:★重要';
    const label = getOrCreateLabel_('整理済み');
    for (const t of GmailApp.search(q, 0, MAX_THREADS)) {
      out.push(['おそうじ', senderOf_(t), subjectOf_(t), dateOf_(t), '30日過ぎた既読 → トレイから外す']);
      if (!dryRun) {
        t.addLabel(label);
        t.moveToArchive();
      }
    }
  }

  // ── 結果を書き出し ──
  if (dryRun) {
    writePreview_(out);
    SpreadsheetApp.getUi().alert(
      out.length
        ? `お試し完了! ${out.length} 件が対象です。\n「プレビュー」シートで内容を確認してください。\nメールはまだ何も変更されていません。`
        : '対象のメールはありませんでした。設定を見直すか、このままでOKです。'
    );
  } else {
    appendLog_(out);
    if (isUiAvailable_()) {
      SpreadsheetApp.getUi().alert(`整理完了! ${out.length} 件を処理しました。詳細は「実行ログ」シートへ。`);
    }
  }
}

/* ────────────────────────────────────────────────
 *  ④ 自動実行 ON / OFF
 * ──────────────────────────────────────────────── */
function enableDaily() {
  disableDaily_();
  ScriptApp.newTrigger('autoRun').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('毎朝7時ごろの自動整理を ON にしました。\nOFF にしたいときはメニューからいつでもどうぞ。');
}

function disableDaily() {
  disableDaily_();
  SpreadsheetApp.getUi().alert('自動整理を OFF にしました。');
}

function disableDaily_() {
  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t));
}

/** トリガーから呼ばれる本番実行(確認ダイアログなし) */
function autoRun() { organize_(false); }

/* ────────────────────────────────────────────────
 *  シート出力ヘルパー
 * ──────────────────────────────────────────────── */
function getOrCreatePreviewSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_PREVIEW);
  if (!s) s = ss.insertSheet(SHEET_PREVIEW);
  return s;
}

function getOrCreateLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_LOG);
  if (!s) s = ss.insertSheet(SHEET_LOG);
  return s;
}

function writePreview_(rows) {
  const s = getOrCreatePreviewSheet_();
  s.clear();
  const header = [['種別', '差出人', '件名', '受信日', 'こうなります']];
  s.getRange(1, 1, 1, 5).setValues(header).setFontWeight('bold').setBackground('#e8f0fe');
  if (rows.length) s.getRange(2, 1, rows.length, 5).setValues(rows);
  s.setColumnWidth(1, 110).setColumnWidth(2, 220).setColumnWidth(3, 320)
    .setColumnWidth(4, 100).setColumnWidth(5, 260);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(s);
}

function appendLog_(rows) {
  const s = getOrCreateLogSheet_();
  if (s.getLastRow() === 0) {
    s.getRange(1, 1, 1, 6)
      .setValues([['実行日時', '種別', '差出人', '件名', '受信日', '実行した内容']])
      .setFontWeight('bold').setBackground('#e8f0fe');
  }
  if (!rows.length) return;
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  const withTime = rows.map((r) => [now, ...r]);
  s.getRange(s.getLastRow() + 1, 1, withTime.length, 6).setValues(withTime);
}

/* ────────────────────────────────────────────────
 *  Gmailヘルパー
 * ──────────────────────────────────────────────── */
function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function subjectOf_(t) { return t.getFirstMessageSubject() || '(件名なし)'; }

function senderOf_(t) {
  const m = t.getMessages();
  return m.length ? m[0].getFrom().replace(/<.*>/, '').trim() || m[0].getFrom() : '';
}

function dateOf_(t) {
  return Utilities.formatDate(t.getLastMessageDate(), Session.getScriptTimeZone(), 'MM/dd');
}

function isUiAvailable_() {
  try { SpreadsheetApp.getUi(); return true; } catch (e) { return false; }
}
