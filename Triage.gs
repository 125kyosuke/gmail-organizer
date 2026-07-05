/**
 * ✉️ メール仕分けカード (Gmail Triage)
 * ------------------------------------------------------------------
 * 受信トレイのメールを1枚ずつカードで表示し、キー1発で仕分けるウェブアプリ。
 * このファイルと Triage.html を同じ Apps Script プロジェクトに入れて、
 * 「デプロイ > 新しいデプロイ > ウェブアプリ(自分のみ)」で公開して使います。
 * ================================================================== */

const TRIAGE_PAGE_SIZE = 25;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Triage')
    .setTitle('メール仕分け')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/* ── 受信トレイからカードの束を取ってくる ── */
function triageGetBatch() {
  const threads = GmailApp.search('in:inbox', 0, TRIAGE_PAGE_SIZE);
  const cards = threads.map((t) => {
    let from = '', fromEmail = '', snippet = '';
    try {
      const msgs = t.getMessages();
      const last = msgs[msgs.length - 1];
      const raw = last.getFrom(); // 例: "山田太郎 <taro@example.com>"
      const m = raw.match(/^\s*"?([^"<]*)"?\s*<(.+)>\s*$/);
      from = m ? m[1].trim() || m[2] : raw;
      fromEmail = m ? m[2].trim() : raw;
      snippet = (last.getPlainBody() || '').replace(/\s+/g, ' ').trim().slice(0, 600);
    } catch (e) { /* 取れなくてもカードは出す */ }
    return {
      id: t.getId(),
      from: from,
      fromEmail: fromEmail.toLowerCase(),
      subject: t.getFirstMessageSubject() || '(件名なし)',
      snippet: snippet,
      date: Utilities.formatDate(t.getLastMessageDate(), Session.getScriptTimeZone(), 'M/d'),
      unread: t.isUnread(),
      count: t.getMessageCount(),
    };
  });
  return { cards: cards, inboxLeft: countInbox_() };
}

function countInbox_() {
  // ざっくりの残数(500件で頭打ち)
  return GmailApp.search('in:inbox', 0, 500).length;
}

/* ── 仕分けアクション ──
 * ids: スレッドIDの配列 / action: archive|label|star|trash
 * labelName: labelのとき。文字列1つ or 配列(複数ラベル)どちらでも可 */
function triageAct(ids, action, labelName) {
  const threads = ids.map((id) => GmailApp.getThreadById(id)).filter(Boolean);
  if (action === 'archive') {
    threads.forEach((t) => t.moveToArchive());
  } else if (action === 'label') {
    const ls = normalizeLabels_(labelName).map(triageLabel_);
    threads.forEach((t) => { ls.forEach((l) => t.addLabel(l)); t.moveToArchive(); });
  } else if (action === 'star') {
    const l = triageLabel_('★重要');
    threads.forEach((t) => {
      t.addLabel(l);
      const m = t.getMessages();
      if (m.length) m[0].star();
    });
  } else if (action === 'trash') {
    threads.forEach((t) => t.moveToTrash());
  }
  return threads.length;
}

/* ── 元に戻す ── */
function triageUndo(ids, action, labelName) {
  const threads = ids.map((id) => GmailApp.getThreadById(id)).filter(Boolean);
  if (action === 'archive') {
    threads.forEach((t) => t.moveToInbox());
  } else if (action === 'label') {
    const ls = normalizeLabels_(labelName).map(triageLabel_);
    threads.forEach((t) => { ls.forEach((l) => t.removeLabel(l)); t.moveToInbox(); });
  } else if (action === 'star') {
    const l = triageLabel_('★重要');
    threads.forEach((t) => {
      t.removeLabel(l);
      const m = t.getMessages();
      if (m.length) m[0].unstar();
    });
  } else if (action === 'trash') {
    threads.forEach((t) => t.moveToInbox()); // ゴミ箱から受信トレイへ戻す
  }
  return threads.length;
}

/* ── ラベルボタンの設定(ユーザーごとに保存) ── */
function triageGetLabels() {
  const p = PropertiesService.getUserProperties().getProperty('TRIAGE_LABELS');
  return p ? JSON.parse(p) : ['大学', 'バイト', 'お金', 'メルマガ', 'あとで読む'];
}

function triageSaveLabels(arr) {
  const clean = (arr || []).map((s) => String(s).trim()).filter(Boolean).slice(0, 8);
  PropertiesService.getUserProperties().setProperty('TRIAGE_LABELS', JSON.stringify(clean));
  return clean;
}

function triageLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function normalizeLabels_(v) {
  if (!v) return [];
  return (Array.isArray(v) ? v : [v]).map((s) => String(s).trim()).filter(Boolean);
}
