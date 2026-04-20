/**
 * ============================================================
 * 現場タスク自動登録システム  -  Google Apps Script (Web App版 v3)
 * ============================================================
 *
 * v3変更点:
 *   doGet に以下を追加
 *   - mode=getProperties  → 担当物件一覧DBから物件と日付を取得してJSONで返す
 *   - mode=updateDates    → 指定ページIDの日付3点をPATCHで更新してJSONで返す
 *   - notionPatch() ヘルパー追加
 *
 * ============================================================
 * 【Web アプリとしてデプロイする手順】
 * ============================================================
 *   1. Apps Script エディタを開く
 *   2. このコードを貼り付けて保存（Ctrl+S）
 *   3. 右上「デプロイ」→「既存のデプロイを管理」
 *   4. 鉛筆アイコン（編集）→ バージョン「新バージョン」
 *   5. 「デプロイ」→ URLは変わらずそのまま使えます
 *
 * ============================================================
 */

// ============================================================
// 定数（変更不要）
// ============================================================
const PROPERTY_DB_ID     = '2f56ad84622180a9891bef7e5514fa78'; // 担当物件一覧
const NOTION_API_BASE    = 'https://api.notion.com/v1';
const NOTION_API_VERSION = '2022-06-28';

// ============================================================
// ★ v7追加: GCal マスター化（4工程日付のSSOT = GCal）
// ------------------------------------------------------------
// 設計:
//   - 物件作成 / 日付更新時に GCal 側にも iCalUID 付きの allDay
//     イベントを作る → ダッシュボードはこれを SSOT として参照
//   - Notion の日付プロパティは ミラー（人が見るための控え）
//   - GCal にイベントが無ければ Notion 日付で fallback（移行期間）
//   - iCalUID は (物件, 工程) を一意に識別:
//       mshub-{notionPageIdHex}-{key}@mshub.jp
//       key: chakou / tatemae / shunko / hikiwatashi
//
// 注意:
//   - 工程作成ツール (gas-schedule-creation.js) は別途
//     「🚜○○本体着工」「⚒️○○建て方」「🔑○○引渡し」を作成する
//   - 両方の登録フローを走らせると同じ日に似たイベントが
//     2つ並ぶ可能性がある（SSOTとしての整合性には影響しない）
// ============================================================
const MSHUB_CAL_ID     = 'jognge12@gmail.com'; // 工程マスターカレンダー
const MSHUB_UID_DOMAIN = '@mshub.jp';

// Notion「dates」キー → iCalUID key + 工程ラベル + colorId(1-11) + Notion日付プロパティ名
const MSHUB_PROC = {
  '着工': { key: 'chakou',      label: '着工',   color: '6', notionProp: '本体着工' }, // ORANGE
  '建方': { key: 'tatemae',     label: '建て方', color: '7', notionProp: '建て方'   }, // CYAN
  '竣工': { key: 'shunko',      label: '竣工',   color: '3', notionProp: '竣工'     }, // GRAPE
  '引渡': { key: 'hikiwatashi', label: '引渡し', color: '3', notionProp: '引渡し'   }  // GRAPE
};
const MSHUB_KEY_LIST = ['chakou', 'tatemae', 'shunko', 'hikiwatashi'];

// notionId(hex) + key → iCalUID
function makeMshubUid(notionId, key) {
  return 'mshub-' + String(notionId).replace(/-/g, '') + '-' + key + MSHUB_UID_DOMAIN;
}

// iCalUID で既存イベントを検索（Advanced Calendar Service）
function findMshubEvent(notionId, key) {
  var uid = makeMshubUid(notionId, key);
  try {
    var result = Calendar.Events.list(MSHUB_CAL_ID, { iCalUID: uid, showDeleted: false, maxResults: 5 });
    if (result && result.items && result.items.length > 0) return result.items[0];
  } catch(e) {
    Logger.log('⚠ findMshubEvent(' + key + ') 失敗: ' + e.message);
  }
  return null;
}

// iCalUID を指定して allDay イベントを upsert
// dateStr が null の場合は既存イベントを削除
function upsertMshubEvent(notionId, key, dateStr, propertyName) {
  var proc = null;
  for (var label in MSHUB_PROC) {
    if (MSHUB_PROC[label].key === key) { proc = MSHUB_PROC[label]; break; }
  }
  if (!proc) throw new Error('未知の工程キー: ' + key);

  var existing = findMshubEvent(notionId, key);

  // 日付未指定 → 既存削除
  if (!dateStr) {
    if (existing) {
      try { Calendar.Events.remove(MSHUB_CAL_ID, existing.id); }
      catch(e) { Logger.log('⚠ イベント削除失敗 (' + key + '): ' + e.message); }
    }
    return null;
  }

  // allDay は end = 翌日
  var startD = new Date(dateStr + 'T00:00:00+09:00');
  var endD   = new Date(startD.getTime() + 24 * 60 * 60 * 1000);
  var endStr = Utilities.formatDate(endD, 'Asia/Tokyo', 'yyyy-MM-dd');
  var title  = (propertyName || '物件') + ' ' + proc.label;

  var body = {
    summary: title,
    start:   { date: dateStr },
    end:     { date: endStr },
    colorId: proc.color
  };

  if (existing) {
    try {
      Calendar.Events.patch(body, MSHUB_CAL_ID, existing.id);
      return existing.id;
    } catch(e) {
      Logger.log('⚠ イベント更新失敗 (' + key + '): ' + e.message);
      return null;
    }
  }

  // 新規: iCalUID を付与して insert
  body.iCalUID = makeMshubUid(notionId, key);
  try {
    var created = Calendar.Events.insert(body, MSHUB_CAL_ID);
    return created.id;
  } catch(e) {
    Logger.log('⚠ イベント作成失敗 (' + key + '): ' + e.message);
    // 409 などの場合は再検索 → update
    var retry = findMshubEvent(notionId, key);
    if (retry) {
      try {
        delete body.iCalUID;
        Calendar.Events.patch(body, MSHUB_CAL_ID, retry.id);
        return retry.id;
      } catch(e2) { Logger.log('⚠ リトライ失敗 (' + key + '): ' + e2.message); }
    }
    return null;
  }
}

// 物件の4工程イベントを全て削除（rollback/archive時）
function deleteAllMshubEvents(notionId) {
  MSHUB_KEY_LIST.forEach(function(k) {
    var ev = findMshubEvent(notionId, k);
    if (ev) {
      try { Calendar.Events.remove(MSHUB_CAL_ID, ev.id); }
      catch(e) { Logger.log('⚠ 削除失敗 (' + k + '): ' + e.message); }
    }
  });
}

// 複数工程を一括 upsert（dates: { chakou: 'YYYY-MM-DD' | null, tatemae: ..., ... }）
function bulkUpsertMshubEvents(notionId, datesByKey, propertyName) {
  MSHUB_KEY_LIST.forEach(function(k) {
    if (datesByKey.hasOwnProperty(k)) {
      upsertMshubEvent(notionId, k, datesByKey[k], propertyName);
    }
  });
}


// ============================================================
// Web アプリ: フォーム画面を返す (GET)
// ★ v2: mode=submit でポータルからの登録に対応
// ★ v3: mode=getProperties / mode=updateDates を追加
// ============================================================
function doGet(e) {
  // アイコン配信
  if (e && e.parameter && e.parameter.mode === 'icon') {
    var svg = '<svg xmlns="http://www.w3.org/2000/svg" width="180" height="180" viewBox="0 0 180 180"><rect width="180" height="180" rx="40" fill="%231D6B40"/><path d="M90 35L30 90h15v55h35v-35h20v35h35V90h15L90 35z" fill="white"/></svg>';
    return ContentService.createTextOutput(svg).setMimeType(ContentService.MimeType.XML);
  }

  // ★ v3追加: 物件一覧取得
  if (e && e.parameter && e.parameter.mode === 'getProperties') {
    return handleGetProperties();
  }

  // ★ v3追加: 日付更新
  if (e && e.parameter && e.parameter.mode === 'updateDates') {
    return handleUpdateDates(e.parameter);
  }

  // ★ ポータルサイトからのフォーム送信
  if (e && e.parameter && e.parameter.mode === 'submit') {
    return handlePortalSubmit(e.parameter);
  }

  // ★ v6追加: 物件ページロールバック（カレンダー登録失敗時用）
  if (e && e.parameter && e.parameter.mode === 'rollbackProperty') {
    return handleRollbackProperty(e.parameter);
  }

  // ★ v4追加: ダッシュボードデータ一括取得（優先タスク + カレンダー）
  if (e && e.parameter && e.parameter.mode === 'getDashboardData') {
    return handleGetDashboardData();
  }

  // ★ v5追加: Googleタスクを完了にする
  if (e && e.parameter && e.parameter.mode === 'completeGoogleTask') {
    return handleCompleteGoogleTask(e.parameter);
  }

  // 通常: フォーム画面を返す（iPhoneから直接アクセスした場合）
  return HtmlService.createHtmlOutput(getFormHtml())
    .setTitle('新規物件登録')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
// ★ v3追加: 担当物件一覧DB から物件と日付を取得してJSONで返す
// ============================================================
function handleGetProperties() {
  try {
    var token = getToken();

    // 工程作成対象の物件のみ取得:
    //   - 「進捗」に「引渡し」を含む物件を除外（引渡し済み）
    //   - 「本体着工」が今日以前の物件を除外（着工済み）
    //   → 着工前の物件のみ、本体着工日の昇順で表示
    var now = new Date();
    var todayStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');

    var result = notionPost('/databases/' + PROPERTY_DB_ID + '/query', {
      filter: {
        and: [
          { property: '進捗', select: { does_not_equal: '引渡し' } },
          {
            or: [
              { property: '本体着工', date: { is_empty: true } },
              { property: '本体着工', date: { after: todayStr } }
            ]
          }
        ]
      },
      sorts: [{ property: '本体着工', direction: 'ascending' }],
      page_size: 100
    }, token);

    if (result.object === 'error') throw new Error(result.message);

    var properties = (result.results || []).map(function(page) {
      var props = page.properties || {};
      var titleArr = props['物件名'] && props['物件名'].title ? props['物件名'].title : [];
      var name = titleArr.length > 0 ? titleArr[0].plain_text : '';
      var chakou      = props['本体着工'] && props['本体着工'].date ? props['本体着工'].date.start : null;
      var tatemae     = props['建て方']   && props['建て方'].date   ? props['建て方'].date.start   : null;
      var hikiwatashi = props['引渡し']   && props['引渡し'].date   ? props['引渡し'].date.start   : null;
      var city        = (props['市町村'] && props['市町村'].select) ? props['市町村'].select.name : '';
      return { id: page.id, name: name, chakou: chakou, tatemae: tatemae, hikiwatashi: hikiwatashi, city: city };
    }).filter(function(p) { return p.name !== ''; });

    return ContentService.createTextOutput(JSON.stringify({
      success: true, properties: properties
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('❌ getProperties エラー: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false, message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
// ★ v3追加: 指定ページIDの日付プロパティをPATCHで更新
// パラメータ: pageId, 本体着工, 建て方, 引渡し（日付文字列 YYYY-MM-DD）
// ============================================================
function handleUpdateDates(p) {
  try {
    var token  = getToken();
    var pageId = p['pageId'];
    if (!pageId) throw new Error('pageId が指定されていません');

    // ★ v7: 4工程すべてを受け付け（竣工も追加）
    var chakou      = p['本体着工'] || '';
    var tatemae     = p['建て方']   || '';
    var shunko      = p['竣工']     || '';
    var hikiwatashi = p['引渡し']   || '';

    // 何も指定されていない場合は no-op
    if (!chakou && !tatemae && !shunko && !hikiwatashi) {
      return ContentService.createTextOutput(JSON.stringify({
        success: true, message: '更新対象なし'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ── Notion 側更新（ミラー） ──
    var updates = {};
    if (chakou)      updates['本体着工'] = { date: { start: chakou } };
    if (tatemae)     updates['建て方']   = { date: { start: tatemae } };
    if (shunko)      updates['竣工']     = { date: { start: shunko } };
    if (hikiwatashi) updates['引渡し']   = { date: { start: hikiwatashi } };

    var result = notionPatch('/pages/' + pageId, { properties: updates }, token);
    if (result.object === 'error') throw new Error(result.message);

    // 物件名取得（GCalイベントのタイトル用）
    var propertyName = '物件';
    try {
      var tArr = (result.properties && result.properties['物件名'] && result.properties['物件名'].title) ? result.properties['物件名'].title : [];
      if (tArr.length > 0) propertyName = tArr[0].plain_text;
    } catch(e) { /* ignore */ }

    // ── GCal マスター側 upsert（4工程、指定されたものだけ） ──
    var gcalErrs = [];
    try {
      var notionId = String(pageId).replace(/-/g, '');
      var datesByKey = {};
      if (chakou)      datesByKey.chakou      = chakou;
      if (tatemae)     datesByKey.tatemae     = tatemae;
      if (shunko)      datesByKey.shunko      = shunko;
      if (hikiwatashi) datesByKey.hikiwatashi = hikiwatashi;
      bulkUpsertMshubEvents(notionId, datesByKey, propertyName);
    } catch(gcalErr) {
      gcalErrs.push(gcalErr.message);
      Logger.log('⚠ GCal upsert 失敗: ' + gcalErr.message);
    }

    Logger.log('✓ 日付更新完了: ' + pageId + (gcalErrs.length ? ' (GCal警告: ' + gcalErrs.join('; ') + ')' : ''));
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      gcalWarnings: gcalErrs
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('❌ updateDates エラー: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false, message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
// ★ v2追加: ポータルからのfetch送信を処理してJSONを返す
// ============================================================
function handlePortalSubmit(p) {
  try {
    var propertyName = (p['物件名'] || '').trim();
    if (!propertyName) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false, message: '物件名が入力されていません'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var city         = p['市町村'] || '';
    var startDate    = p['本体着工'] || null;
    var frameDate    = p['建て方']   || null;
    var shunkoDate   = p['竣工']     || null;
    var deliveryDate = p['引渡し']   || null;
    var sakiSoto    = p['先外']    === 'on';
    var kairo       = p['改良']    === 'on';
    var gaiko       = p['外構']    === 'on';
    var shizumono   = p['鎮物']    === 'on';
    var munefuda    = p['棟札']    === 'on';
    var tegata      = p['手形']    === 'on';
    var shikyuhin   = p['支給品']  === 'on';
    var kansetsuShomei = p['間接照明'] === 'on';

    Logger.log('▶ ポータル登録開始: ' + propertyName);

    var propertyPageId = createPropertyPage(
      propertyName, city,
      startDate || null, frameDate || null, shunkoDate || null, deliveryDate || null,
      sakiSoto, kairo, gaiko, shizumono, munefuda,
      tegata, shikyuhin, kansetsuShomei
    );
    Logger.log('✓ 物件作成完了 (ID: ' + propertyPageId + ')');

    Logger.log('✅ 完了: 「' + propertyName + '」');
    return ContentService.createTextOutput(JSON.stringify({
      success: true, name: propertyName, pageId: propertyPageId
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('❌ エラー: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false, message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
// ★ v6追加: 物件ページロールバック（Notionページをアーカイブ）
// カレンダー登録失敗時にフロント側から呼ばれる
// ============================================================
function handleRollbackProperty(p) {
  try {
    var pageId = (p.pageId || '').trim();
    if (!pageId) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false, message: 'pageId が未指定です'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var token = getToken();

    // ★ v7: GCal 側の iCalUID イベントを先に削除
    try {
      var notionId = String(pageId).replace(/-/g, '');
      deleteAllMshubEvents(notionId);
    } catch(gcalErr) {
      Logger.log('⚠ GCal削除失敗（Notionアーカイブは続行）: ' + gcalErr.message);
    }

    // Notion ページのアーカイブ（archived: true）
    var result = notionPatch('/pages/' + pageId, { archived: true }, token);
    if (result.object === 'error') {
      throw new Error(result.message);
    }

    Logger.log('✓ 物件ページをアーカイブ: ' + pageId);
    return ContentService.createTextOutput(JSON.stringify({
      success: true, pageId: pageId
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('❌ rollback エラー: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false, message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
// Web アプリ: フォーム送信を受け取る (POST) ※既存フォーム用
// ============================================================
function doPost(e) {
  try {
    const p = e.parameter;

    const propertyName = (p['物件名'] || '').trim();
    if (!propertyName) {
      return HtmlService.createHtmlOutput(getResultHtml(false, '物件名が入力されていません'));
    }

    const city         = p['市町村'] || '';
    const startDate    = p['本体着工'] || null;
    const frameDate    = p['建て方']   || null;
    const shunkoDate   = p['竣工']     || null;
    const deliveryDate = p['引渡し']   || null;
    const sakiSoto      = p['先外']     === 'on';
    const kairo         = p['改良']     === 'on';
    const gaiko         = p['外構']     === 'on';
    const shizumono     = p['鎮物']     === 'on';
    const munefuda      = p['棟札']     === 'on';
    const tegata        = p['手形']     === 'on';
    const shikyuhin     = p['支給品']   === 'on';
    const kansetsuShomei = p['間接照明'] === 'on';

    Logger.log('▶ Web登録開始: ' + propertyName);

    const propertyPageId = createPropertyPage(
      propertyName, city, startDate, frameDate, shunkoDate, deliveryDate,
      sakiSoto, kairo, gaiko, shizumono, munefuda,
      tegata, shikyuhin, kansetsuShomei
    );
    Logger.log('✓ 物件作成完了 (ID: ' + propertyPageId + ')');

    Logger.log('✅ 完了: 「' + propertyName + '」');
    return HtmlService.createHtmlOutput(
      getResultHtml(true, propertyName, 0)
    );

  } catch (err) {
    Logger.log('❌ エラー: ' + err.message);
    return HtmlService.createHtmlOutput(getResultHtml(false, err.message));
  }
}


// ============================================================
// サーバー関数: google.script.run から呼ばれる
// ============================================================
function processForm(formData) {
  var propertyName = (formData['物件名'] || '').trim();
  if (!propertyName) return { success: false, message: '物件名が入力されていません' };
  var city = formData['市町村'] || '';
  var startDate = formData['本体着工'] || null;
  var frameDate = formData['建て方'] || null;
  var shunkoDate = formData['竣工'] || null;
  var deliveryDate = formData['引渡し'] || null;
  var sakiSoto     = formData['先外']     === true;
  var kairo        = formData['改良']     === true;
  var gaiko        = formData['外構']     === true;
  var shizumono    = formData['鎮物']     === true;
  var munefuda     = formData['棟札']     === true;
  var tegata       = formData['手形']     === true;
  var shikyuhin    = formData['支給品']   === true;
  var kansetsuShomei = formData['間接照明'] === true;
  Logger.log('▶ Web登録開始: ' + propertyName);
  try {
    var propertyPageId = createPropertyPage(propertyName, city, startDate, frameDate, shunkoDate, deliveryDate, sakiSoto, kairo, gaiko, shizumono, munefuda, tegata, shikyuhin, kansetsuShomei);
    return { success: true, name: propertyName };
  } catch (err) {
    Logger.log('✘ エラー: ' + err.message);
    return { success: false, message: err.message };
  }
}


// ============================================================
// HTML: 登録フォーム（iPhoneから直接アクセスした場合に表示）
// ============================================================
function getFormHtml() {
  var gasUrl = ScriptApp.getService().getUrl();
  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="物件登録">
<title>新規物件登録</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Hiragino Kaku Gothic ProN', sans-serif;
    font-size: 16px;
    background: #f0f4f2;
    min-height: 100vh;
    padding: 0 0 20px;
  }
  header { display: none; }
  .card {
    background: #fff;
    border-radius: 14px;
    margin: 8px 16px;
    padding: 12px 14px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
  }
  .card h2 {
    font-size: 11px;
    font-weight: 700;
    color: #1D6B40;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-bottom: 10px;
    padding-bottom: 8px;
    border-bottom: 1px solid #e8f0eb;
  }
  .field { margin-bottom: 10px; }
  .field:last-child { margin-bottom: 0; }
  label {
    display: block;
    font-size: 13px;
    font-weight: 600;
    color: #333;
    margin-bottom: 5px;
  }
  label .required {
    color: #e05555;
    font-size: 11px;
    margin-left: 3px;
    font-weight: 500;
  }
  input[type="text"],
  input[type="date"],
  select {
    display: block;
    width: 100%;
    padding: 8px 10px;
    font-size: 14px;
    font-family: inherit;
    border: 1.5px solid #dde8e2;
    border-radius: 10px;
    background: #fff;
    color: #222;
    -webkit-appearance: none;
    appearance: none;
    transition: border-color 0.2s;
  }
  input[type="text"]:focus,
  input[type="date"]:focus,
  select:focus {
    outline: none;
    border-color: #2D8A4E;
    box-shadow: 0 0 0 3px rgba(45,138,78,0.12);
  }
  select {
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%23999' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-position: right 12px center;
    padding-right: 32px;
  }
  input[type="date"] {
    -webkit-appearance: auto;
    appearance: auto;
  }
  .date-info {
    font-size: 12px;
    margin-top: 4px;
    padding-left: 2px;
    font-weight: 600;
  }
  .date-fields { display: grid; grid-template-columns: 1fr; gap: 10px; }
  .check-grid { display: grid; grid-template-columns: 1fr; gap: 7px; }
  .check-item {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 9px 12px;
    border: 1.5px solid #dde8e2;
    border-radius: 10px;
    cursor: pointer;
    transition: background 0.15s;
    -webkit-tap-highlight-color: transparent;
  }
  .check-item:active { background: #f0f8f4; }
  .check-item input[type="checkbox"] {
    width: 20px; height: 20px;
    border-radius: 6px;
    border: 2px solid #bbb;
    appearance: none; -webkit-appearance: none;
    background: #fff;
    cursor: pointer;
    flex-shrink: 0;
    transition: all 0.15s;
    position: relative;
  }
  .check-item input[type="checkbox"]:checked {
    background: #2D8A4E;
    border-color: #2D8A4E;
  }
  .check-item input[type="checkbox"]:checked::after {
    content: '';
    position: absolute;
    left: 5px; top: 2px;
    width: 6px; height: 10px;
    border: 2.5px solid #fff;
    border-top: none; border-left: none;
    transform: rotate(45deg);
  }
  .check-item span {
    font-size: 13px;
    font-weight: 500;
    color: #333;
  }
  .btn-submit {
    display: block;
    width: calc(100% - 32px);
    margin: 12px 16px;
    padding: 12px;
    background: #1D6B40;
    color: #fff;
    font-size: 14px;
    font-weight: 700;
    font-family: inherit;
    border: none;
    border-radius: 12px;
    cursor: pointer;
    box-shadow: 0 4px 14px rgba(29,107,64,0.3);
    -webkit-tap-highlight-color: transparent;
    transition: opacity 0.15s;
  }
  .btn-submit:active { opacity: 0.8; }
  .btn-submit:disabled { opacity: 0.6; cursor: not-allowed; }
</style>
</head>
<body>
<header>
  <div>
    <h1>新規物件登録</h1>
    <p>Notion 担当物件一覧へ登録</p>
  </div>
</header>

<form method="POST" id="regForm" onsubmit="handleSubmit(event)">

  <div class="card">
    <h2>基本情報</h2>
    <div class="field">
      <label>物件名 <span class="required">*</span></label>
      <input type="text" name="物件名" id="propName" placeholder="〇〇様邸" required>
    </div>
    <div class="field">
      <label>市町村</label>
      <select name="市町村">
        <option value="">（選択してください）</option>
        <option>岐阜市</option><option>大垣市</option><option>高山市</option>
        <option>多治見市</option><option>関市</option><option>中津川市</option>
        <option>美濃市</option><option>瑞浪市</option><option>羽島市</option>
        <option>恵那市</option><option>美濃加茂市</option><option>土岐市</option>
        <option>各務原市</option><option>可児市</option><option>山県市</option>
        <option>瑞穂市</option><option>飛騨市</option><option>本巣市</option>
        <option>郡上市</option>
      </select>
    </div>
  </div>

  <div class="card">
    <h2>工程</h2>
    <div class="date-fields">
      <div class="field">
        <label>本体着工</label>
        <input type="date" name="本体着工" id="d_chakou" oninput="updateDay('d_chakou','dw_chakou')">
        <div id="dw_chakou" class="date-info"></div>
      </div>
      <div class="field">
        <label>建て方</label>
        <input type="date" name="建て方" id="d_tatemae" oninput="updateDay('d_tatemae','dw_tatemae')">
        <div id="dw_tatemae" class="date-info"></div>
      </div>
      <div class="field">
        <label>引渡し</label>
        <input type="date" name="引渡し" id="d_hiki" oninput="updateDay('d_hiki','dw_hiki')">
        <div id="dw_hiki" class="date-info"></div>
      </div>
    </div>
  </div>

  <div class="card">
    <h2>チェック項目</h2>
    <div class="check-grid">
      <label class="check-item">
        <input type="checkbox" name="棟札" value="on"> <span>棟札</span>
      </label>
      <label class="check-item">
        <input type="checkbox" name="鎮物" value="on"> <span>鎮め物</span>
      </label>
      <label class="check-item">
        <input type="checkbox" name="先外" value="on"> <span>先行外構</span>
      </label>
      <label class="check-item">
        <input type="checkbox" name="改良" value="on"> <span>地盤改良</span>
      </label>
      <label class="check-item">
        <input type="checkbox" name="外構" value="on"> <span>外構工事</span>
      </label>
      <label class="check-item">
        <input type="checkbox" name="手形" value="on"> <span>手形</span>
      </label>
      <label class="check-item">
        <input type="checkbox" name="支給品" value="on"> <span>支給品</span>
      </label>
      <label class="check-item">
        <input type="checkbox" name="間接照明" value="on"> <span>間接照明</span>
      </label>
    </div>
  </div>

  <button type="submit" class="btn-submit" id="submitBtn">Notion に登録する</button>

</form>

<script>
var DAY_NAMES = ['\u65e5','\u6708','\u706b','\u6c34','\u6728','\u91d1','\u571f'];
var GAS_BASE_URL = '${gasUrl}';

function updateDay(inputId, infoId) {
  var val = document.getElementById(inputId).value;
  var info = document.getElementById(infoId);
  if (!val) { info.textContent = ''; return; }
  var d   = new Date(val);
  var dow = d.getDay();
  var dayColor = dow === 0 ? '#ef4444' : dow === 6 ? '#3b82f6' : '#1e293b';
  info.innerHTML =
    '<span style="color:#1e293b">' + (d.getMonth()+1) + '\u6708' + d.getDate() + '\u65e5</span> ' +
    '<span style="color:' + dayColor + ';font-weight:700">(' + DAY_NAMES[dow] + ')</span>';
}

// フォーム送信: fetch API（iframe+iOSのホワイトアウト回避）
async function handleSubmit(e) {
  e.preventDefault();
  var btn = document.getElementById('submitBtn');
  btn.disabled = true;
  btn.textContent = '\u767b\u9332\u4e2d...';

  var formData = new FormData(e.target);
  var params = new URLSearchParams();
  params.set('mode', 'submit');
  formData.forEach(function(val, key) { params.set(key, val); });

  try {
    var res = await fetch(GAS_BASE_URL + '?' + params.toString());
    if (!res.ok) throw new Error('HTTP ' + res.status);
    var data = await res.json();
    if (data.success) {
      document.body.innerHTML =
        '<div style="font-family:-apple-system,sans-serif;text-align:center;padding:48px 24px;background:#f0f4f2;min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;">' +
        '<div style="width:56px;height:56px;border-radius:50%;background:#1D6B40;margin:0 auto 20px;display:flex;align-items:center;justify-content:center;">' +
        '<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg></div>' +
        '<h1 style="font-size:20px;font-weight:700;color:#1D6B40;margin:0 0 6px">\u767b\u9332\u5b8c\u4e86</h1>' +
        '<div style="font-size:16px;font-weight:600;color:#333;margin-bottom:32px">\u300c' + data.name + '\u300d</div>' +
        '<button onclick="window.location.href=GAS_BASE_URL" style="display:block;width:220px;padding:14px 32px;background:#1D6B40;color:#fff;font-size:15px;font-weight:700;border:none;border-radius:12px;cursor:pointer;box-shadow:0 3px 10px rgba(29,107,64,0.3)">\u7d9a\u3051\u3066\u767b\u9332</button>' +
        '</div>';
    } else {
      throw new Error(data.message || '\u767b\u9332\u306b\u5931\u6557\u3057\u307e\u3057\u305f');
    }
  } catch(err) {
    alert('\u30a8\u30e9\u30fc: ' + err.message);
    btn.disabled = false;
    btn.textContent = 'Notion \u306b\u767b\u9332\u3059\u308b';
  }
}
</script>
</body>
</html>`;
}


// ============================================================
// HTML: 結果画面
// ============================================================
function getResultHtml(success, messageOrName, taskCount) {
  if (success) {
    return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>登録完了</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, sans-serif; background: #f0f4f2; min-height: 100vh;
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    padding: 40px 24px; text-align: center;
  }
  .icon {
    width: 56px; height: 56px; border-radius: 50%;
    background: #1D6B40; margin: 0 auto 20px;
    display: flex; align-items: center; justify-content: center;
  }
  h1 { font-size: 20px; font-weight: 700; color: #1D6B40; margin-bottom: 6px; }
  .name { font-size: 16px; font-weight: 600; color: #333; margin-bottom: 32px; }
  a.btn {
    display: block; padding: 14px 32px;
    background: #1D6B40; color: #fff;
    font-size: 15px; font-weight: 700;
    border-radius: 12px; text-decoration: none;
    box-shadow: 0 3px 10px rgba(29,107,64,0.3);
  }
</style>
</head>
<body>
  <div class="icon"><svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg></div>
  <h1>登録完了</h1>
  <div class="name">「${messageOrName}」</div>
  <a class="btn" href="javascript:history.back()">続けて登録</a>
</body>
</html>`;
  } else {
    return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>エラー</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, sans-serif; background: #fff5f5; min-height: 100vh;
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    padding: 40px 24px; text-align: center;
  }
  .icon { font-size: 64px; margin-bottom: 20px; }
  h1 { font-size: 20px; font-weight: 700; color: #c0392b; margin-bottom: 12px; }
  p { font-size: 14px; color: #666; margin-bottom: 32px; line-height: 1.6; }
  a.btn {
    display: block; padding: 14px 32px;
    background: #555; color: #fff;
    font-size: 16px; font-weight: 700;
    border-radius: 12px; text-decoration: none;
  }
</style>
</head>
<body>
  <div class="icon">❌</div>
  <h1>登録エラー</h1>
  <p>${messageOrName}</p>
  <a class="btn" href="javascript:history.back()">戻る</a>
</body>
</html>`;
  }
}


// ============================================================
// Notion: 物件ページを作成
// ============================================================
function createPropertyPage(name, city, startDate, frameDate, shunkoDate, deliveryDate,
                             sakiSoto, kairo, gaiko, shizumono, munefuda,
                             tegata, shikyuhin, kansetsuShomei) {
  const token = getToken();

  const properties = {
    '物件名':   { title: [{ text: { content: name } }] },
    '進捗':     { select: { name: '着工前' } },
    '先外':     { checkbox: !!sakiSoto  },
    '改良':     { checkbox: !!kairo     },
    '外構':     { checkbox: !!gaiko     },
    '鎮物':     { checkbox: !!shizumono },
    '棟札':     { checkbox: !!munefuda  },
    '手形':     { checkbox: !!tegata    },
    '支給品':   { checkbox: !!shikyuhin },
    '間接照明': { checkbox: !!kansetsuShomei }
  };

  if (city && city !== '（選択してください）') {
    properties['市町村'] = { select: { name: city } };
  }
  if (startDate)    properties['本体着工'] = { date: { start: startDate } };
  if (frameDate)    properties['建て方']   = { date: { start: frameDate } };
  if (shunkoDate)   properties['竣工']     = { date: { start: shunkoDate } };
  if (deliveryDate) properties['引渡し']   = { date: { start: deliveryDate } };

  const result = notionPost('/pages', {
    parent:     { database_id: PROPERTY_DB_ID },
    properties: properties
  }, token);

  if (result.object === 'error') {
    throw new Error('物件作成エラー: ' + result.message);
  }

  // ★ v7.1: 物件作成では GCal には書かない。
  //         GCalイベントは工程作成ツール(registerDirect)側で iCalUID 付きで作る。
  //         → 1回の登録フローで作成経路が1本になり重複なし。
  return result.id;
}


// ============================================================
// Notion API ヘルパー: POST
// ============================================================
function notionPost(endpoint, payload, token) {
  var MAX_RETRIES = 3;

  for (var attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    var res = UrlFetchApp.fetch(NOTION_API_BASE + endpoint, {
      method:  'post',
      headers: {
        'Authorization':  'Bearer ' + token,
        'Notion-Version': NOTION_API_VERSION,
        'Content-Type':   'application/json'
      },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    var body = res.getContentText();

    if (code >= 200 && code < 500) {
      try {
        return JSON.parse(body);
      } catch (e) {
        Logger.log('⚠ JSONパース失敗 (HTTP ' + code + '): ' + body.substring(0, 200));
        return { object: 'error', message: 'レスポンスがJSONではありません (HTTP ' + code + ')' };
      }
    }

    var wait = (code === 429) ? 2000 : 1000 * attempt;
    Logger.log('⚠ HTTP ' + code + ' - ' + wait + 'ms後にリトライ (' + attempt + '/' + MAX_RETRIES + ')');
    Utilities.sleep(wait);
  }

  return { object: 'error', message: 'Notion API が応答しません (HTTP ' + res.getResponseCode() + ')' };
}


// ============================================================
// ★ v3追加: Notion API ヘルパー: PATCH
// ============================================================
function notionPatch(endpoint, payload, token) {
  var MAX_RETRIES = 3;

  for (var attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    var res = UrlFetchApp.fetch(NOTION_API_BASE + endpoint, {
      method:  'patch',
      headers: {
        'Authorization':  'Bearer ' + token,
        'Notion-Version': NOTION_API_VERSION,
        'Content-Type':   'application/json'
      },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    var body = res.getContentText();

    if (code >= 200 && code < 500) {
      try {
        return JSON.parse(body);
      } catch (e) {
        return { object: 'error', message: 'レスポンスがJSONではありません (HTTP ' + code + ')' };
      }
    }

    var wait = (code === 429) ? 2000 : 1000 * attempt;
    Logger.log('⚠ PATCH HTTP ' + code + ' - リトライ (' + attempt + '/' + MAX_RETRIES + ')');
    Utilities.sleep(wait);
  }

  return { object: 'error', message: 'Notion API が応答しません' };
}


function getToken() {
  const token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('スクリプトプロパティ「NOTION_TOKEN」が設定されていません');
  return token;
}

function parseDate(dateStr) {
  if (!dateStr || dateStr.trim() === '') return null;
  const s = dateStr.trim();
  const m1 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m1) return m1[1] + '-' + m1[2].padStart(2,'0') + '-' + m1[3].padStart(2,'0');
  const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m2) return m2[3] + '-' + m2[1].padStart(2,'0') + '-' + m2[2].padStart(2,'0');
  try {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  } catch(e) {}
  return null;
}


// ============================================================
// ★ v5追加: Googleタスクを完了にする
// mode=completeGoogleTask&taskId=xxx&listId=yyy
// ============================================================
function handleCompleteGoogleTask(params) {
  try {
    var taskId = params.taskId;
    var listId = params.listId;
    if (!taskId || !listId) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false, error: 'パラメータ不足（taskId・listIdが必要）'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    Tasks.Tasks.patch({ status: 'completed' }, listId, taskId);
    return ContentService.createTextOutput(JSON.stringify({
      success: true
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    Logger.log('❌ completeGoogleTask エラー: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false, error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// ★ v4追加: ダッシュボードデータ一括取得
// mode=getDashboardData
//   - Googleカレンダーの着工・引渡しイベントをリアルタイム取得
//   - 返値: JSON { success, timestamp, calendarData }
// ============================================================
function handleGetDashboardData() {
  try {
    var token = getToken();
    var now = new Date();
    var timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    var today     = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
    var currentMonth = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'M')) - 1; // 0-based

    // ── 1. Notion物件データ（物件ボード＋スケジュール統一データソース）──
    var propsResult = notionPost('/databases/' + PROPERTY_DB_ID + '/query', { page_size: 100 }, token);
    var propertyMap = {};
    var allNotionProperties = [];
    var startedProperties  = [];
    var monthlyStarts      = [];
    var deliveredProperties = [];
    var monthlyDeliveries  = [];
    var yearlyStartCounts    = [0,0,0,0,0,0,0,0,0,0,0,0];
    var yearlyDeliveryCounts = [0,0,0,0,0,0,0,0,0,0,0,0];
    var propIdCounter = 1;
    var thisYear = now.getFullYear();

    // ★ v7: GCal スキャン中に iCalUID=mshub-* を拾って日付マスターを組み立て
    // 形式: mshubDateMap[notionId][key] = 'YYYY-MM-DD'
    //   key: chakou | tatemae | shunko | hikiwatashi
    var mshubDateMap = {};
    var MSHUB_UID_RE = /^mshub-([a-f0-9]+)-(chakou|tatemae|shunko|hikiwatashi)@/i;

    function toSlash(d) { return d ? d.replace(/-/g, '/') : null; }
    function sortAndClean(arr) {
      arr.sort(function(a, b) { return a._sort - b._sort; });
      arr.forEach(function(item) { delete item._sort; });
      return arr;
    }

    // ── 第1パス: Notion からページ情報収集（集計はまだ） ──
    // ※ 日付の最終確定は GCal スキャン後に行う（GCal優先, Notion fallback）
    var _notionRaw = [];  // [{ pageId, name, displayName, city, shinchoku, checks, notionDates }]
    (propsResult.results || []).forEach(function(page) {
      var pr   = page.properties || {};
      var tArr = (pr['物件名'] && pr['物件名'].title) ? pr['物件名'].title : [];
      var name = tArr.length > 0 ? tArr[0].plain_text : '';
      if (!name) return;
      // テンプレート・原本ページを除外
      if (name === '原本' || name === '原本(コピー)') return;

      var chakou = (pr['本体着工'] && pr['本体着工'].date) ? pr['本体着工'].date.start : null;
      var tate   = (pr['建て方']   && pr['建て方'].date)   ? pr['建て方'].date.start   : null;
      var shunko = (pr['竣工']     && pr['竣工'].date)     ? pr['竣工'].date.start     : null;
      var hiki   = (pr['引渡し']   && pr['引渡し'].date)   ? pr['引渡し'].date.start   : null;

      // 市町村（select）
      var city = (pr['市町村'] && pr['市町村'].select) ? pr['市町村'].select.name : '';

      // 進捗（select → name を直接取得。未設定は着工前扱い）
      var shinchoku = (pr['進捗'] && pr['進捗'].select) ? pr['進捗'].select.name : '着工前';

      // チェック項目
      var checks = {
        '棟札':     pr['棟札']     ? pr['棟札'].checkbox     : false,
        '鎮物':     pr['鎮物']     ? pr['鎮物'].checkbox     : false,
        '先外':     pr['先外']     ? pr['先外'].checkbox     : false,
        '改良':     pr['改良']     ? pr['改良'].checkbox     : false,
        '外構':     pr['外構']     ? pr['外構'].checkbox     : false,
        '手形':     pr['手形']     ? pr['手形'].checkbox     : false,
        '支給品':   pr['支給品']   ? pr['支給品'].checkbox   : false,
        '間接照明': pr['間接照明'] ? pr['間接照明'].checkbox : false
      };

      var displayName = name.replace(/邸$/, '');
      var notionIdHex = page.id.replace(/-/g, '');

      _notionRaw.push({
        pageId: page.id,
        notionIdHex: notionIdHex,
        name: name,
        displayName: displayName,
        city: city,
        shinchoku: shinchoku,
        checks: checks,
        notionDates: { chakou: chakou, tatemae: tate, shunko: shunko, hikiwatashi: hiki }
      });
    });

    // ── 2. Googleカレンダー（週間スケジュールのみ）──────────────────
    var calErrorMsg = '';
    var scheduleEvents = [];
    var todayStr     = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
    var tomorrowDate = new Date(now.getTime() + 24 * 60 * 60 * 1000);
    var tomorrowStr  = Utilities.formatDate(tomorrowDate, 'Asia/Tokyo', 'yyyy-MM-dd');

    // 3月1日〜12月31日のスケジュールを収集（年間カバー）
    var dayOfWeek = now.getDay();
    var weekSunday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek);
    var schedMap = {};
    var weekDateStrs = [];
    // 過去2ヶ月〜未来12ヶ月（今日基準で動的に算出）
    var schedStart = new Date(now.getFullYear(), now.getMonth() - 2, 1); // 2ヶ月前の1日
    var schedEnd   = new Date(now.getFullYear(), now.getMonth() + 13, 0); // 12ヶ月後の月末
    for (var wd = new Date(schedStart); wd <= schedEnd; wd.setDate(wd.getDate() + 1)) {
      var wdStr = Utilities.formatDate(wd, 'Asia/Tokyo', 'yyyy-MM-dd');
      schedMap[wdStr] = [];
      weekDateStrs.push(wdStr);
    }

    // イベント個別色 → Hex変換マップ（Googleカレンダー EventColor ID）
    var evColorMap = {
      '1':'#a4bdfc','2':'#7ae7bf','3':'#dbadff','4':'#ff887c',
      '5':'#fbd75b','6':'#ffb878','7':'#46d6db','8':'#e1e1e1',
      '9':'#5484ed','10':'#51b749','11':'#dc2127'
    };

    try {
      var year       = now.getFullYear();
      // タイムゾーン明示（toISOString()はUTC変換するため+09:00で指定）
      var rangeStartISO = year + '-01-01T00:00:00+09:00';
      var rangeEndISO   = (year + 1) + '-01-01T00:00:00+09:00';

      // Calendar Advanced Service でカレンダー一覧取得
      var calListResult = Calendar.CalendarList.list();
      var calendars = calListResult.items || [];

      // 同一イベントの二重表示を防ぐため (iCalUID, 日付) で重複除去
      // （招待や共有で同じイベントが複数カレンダーに現れる場合に備える）
      var seenEventKeys = {};

      calendars.forEach(function(cal) {
        var calId    = cal.id;
        var calName  = cal.summary || '';
        var calColor = cal.backgroundColor || '#4285f4';
        var isSys    = calName.indexOf('祝日') !== -1 || calName.toLowerCase().indexOf('birthday') !== -1;

        // イベント取得（ページネーション対応）
        var pageToken = null;
        do {
          var params = {
            timeMin: rangeStartISO,
            timeMax: rangeEndISO,
            singleEvents: true,
            orderBy: 'startTime',
            maxResults: 2500
          };
          if (pageToken) params.pageToken = pageToken;
          var evResult = Calendar.Events.list(calId, params);
          var items = evResult.items || [];

          items.forEach(function(ev) {
            var evTitle   = ev.summary || '';
            var isAllDay  = !!ev.start.date;
            var evStartRaw = isAllDay ? ev.start.date : ev.start.dateTime;
            // 全日イベントはJST明示（"2026-04-09"→UTC解釈を防止）
            var evStart   = isAllDay ? new Date(evStartRaw + 'T00:00:00+09:00') : new Date(evStartRaw);
            var evDateStr = Utilities.formatDate(evStart, 'Asia/Tokyo', 'yyyy-MM-dd');

            // ★ v7: iCalUID=mshub-* をキャプチャして日付マスターマップへ
            //      （scheduleEvents への追加は通常ルートで続行 → 月間/週間表示に出る）
            if (ev.iCalUID) {
              var mshubMatch = String(ev.iCalUID).match(MSHUB_UID_RE);
              if (mshubMatch) {
                var mshubNotionId = mshubMatch[1].toLowerCase();
                var mshubKey      = mshubMatch[2].toLowerCase();
                if (!mshubDateMap[mshubNotionId]) mshubDateMap[mshubNotionId] = {};
                mshubDateMap[mshubNotionId][mshubKey] = evDateStr;
              }
            }

            // イベント個別色があればそれを使用、なければカレンダー色
            var evColorId = ev.colorId || '';
            var color = (evColorId && evColorMap[evColorId]) ? evColorMap[evColorId] : calColor;

            // 対象期間の全イベントを収集（祝日カレンダーも含む）
            if (schedMap[evDateStr] !== undefined) {
              // 重複キー: iCalUID（共有/招待で同一）+ 日付 + 開始時刻
              // iCalUID が無ければタイトル+開始時刻にフォールバック
              var dedupKey = (ev.iCalUID || evTitle) + '|' + evDateStr + '|' + (isAllDay ? 'allday' : (evStartRaw || ''));
              if (seenEventKeys[dedupKey]) {
                // 同一イベントを別カレンダーから再取得 → スキップ
                return;
              }
              seenEventKeys[dedupKey] = true;

              var startTime = null, endTime = null;
              if (!isAllDay) {
                startTime = Utilities.formatDate(evStart, 'Asia/Tokyo', 'HH:mm');
                if (ev.end && ev.end.dateTime) {
                  endTime = Utilities.formatDate(new Date(ev.end.dateTime), 'Asia/Tokyo', 'HH:mm');
                }
              }
              var evObj = {
                title:     evTitle,
                isAllDay:  isAllDay,
                startTime: startTime,
                endTime:   endTime,
                color:     color
              };
              // Googleカレンダーの「場所」フィールドがあれば追加
              if (ev.location) {
                evObj.location = ev.location;
              }
              schedMap[evDateStr].push(evObj);
            }

            // （着工・引渡しの集計はNotionデータから算出済み）
          });

          pageToken = evResult.nextPageToken;
        } while (pageToken);
      });

      // scheduleEvents を組み立て（前4週〜後5週の63日分）
      scheduleEvents = weekDateStrs.map(function(ds) {
        return { dateStr: ds, events: schedMap[ds] || [] };
      });

    } catch(calErr) {
      Logger.log('⚠ カレンダー取得エラー: ' + calErr.message);
      calErrorMsg = calErr.message || '不明なエラー';
      scheduleEvents = weekDateStrs.map(function(ds) {
        return { dateStr: ds, events: [] };
      });
    }

    // ── 第2パス: GCal 優先で日付確定 + スケジュール集計 ──
    // ★ v7: 物件の日付は GCal (iCalUID=mshub-*) が SSOT
    //       ただし GCal 側にイベントが無ければ Notion 日付で fallback
    //       （移行期間: 既存物件は GCal 同期されるまで Notion を参照）
    _notionRaw.forEach(function(r) {
      var mshubDates = mshubDateMap[r.notionIdHex] || {};
      // 解決順: GCal > Notion
      var chakou = mshubDates.chakou      || r.notionDates.chakou;
      var tate   = mshubDates.tatemae     || r.notionDates.tatemae;
      var shunko = mshubDates.shunko      || r.notionDates.shunko;
      var hiki   = mshubDates.hikiwatashi || r.notionDates.hikiwatashi;

      // propertyMap（互換性用）
      propertyMap[r.notionIdHex] = { name: r.name, hikiwatashi: hiki, tatemae: tate };

      allNotionProperties.push({
        id: propIdCounter++,
        name: r.displayName,
        location: r.city,
        shinchoku: r.shinchoku,
        notionId: r.notionIdHex,
        dates: { '着工': toSlash(chakou), '建方': toSlash(tate), '竣工': toSlash(shunko), '引渡': toSlash(hiki) },
        checks: r.checks,
        // デバッグ用: どの工程が GCal 由来かのフラグ（必要に応じてフロントで利用可）
        _src: {
          chakou:      mshubDates.chakou      ? 'gcal' : (r.notionDates.chakou      ? 'notion' : null),
          tatemae:     mshubDates.tatemae     ? 'gcal' : (r.notionDates.tatemae     ? 'notion' : null),
          shunko:      mshubDates.shunko      ? 'gcal' : (r.notionDates.shunko      ? 'notion' : null),
          hikiwatashi: mshubDates.hikiwatashi ? 'gcal' : (r.notionDates.hikiwatashi ? 'notion' : null)
        }
      });

      // ── スケジュール集計（解決済み日付を使用）──
      if (chakou) {
        var cDate = new Date(chakou + 'T00:00:00+09:00');
        var cStr  = Utilities.formatDate(cDate, 'Asia/Tokyo', 'yyyy-MM-dd');
        var cMon  = parseInt(Utilities.formatDate(cDate, 'Asia/Tokyo', 'M')) - 1;
        var cDay  = parseInt(Utilities.formatDate(cDate, 'Asia/Tokyo', 'd'));
        var cYear = parseInt(Utilities.formatDate(cDate, 'Asia/Tokyo', 'yyyy'));
        if (cYear === thisYear) yearlyStartCounts[cMon]++;
        if (cStr <= today) {
          startedProperties.push({ name: r.displayName, date: (cMon+1)+'/'+cDay, _sort: cMon*100+cDay });
        } else if (cMon === currentMonth && cYear === thisYear) {
          monthlyStarts.push({ name: r.displayName, date: (cMon+1)+'/'+cDay, _sort: cMon*100+cDay });
        }
      }
      if (hiki) {
        var hDate = new Date(hiki + 'T00:00:00+09:00');
        var hStr  = Utilities.formatDate(hDate, 'Asia/Tokyo', 'yyyy-MM-dd');
        var hMon  = parseInt(Utilities.formatDate(hDate, 'Asia/Tokyo', 'M')) - 1;
        var hDay  = parseInt(Utilities.formatDate(hDate, 'Asia/Tokyo', 'd'));
        var hYear = parseInt(Utilities.formatDate(hDate, 'Asia/Tokyo', 'yyyy'));
        if (hYear === thisYear) yearlyDeliveryCounts[hMon]++;
        if (hStr <= today) {
          deliveredProperties.push({ name: r.displayName, date: (hMon+1)+'/'+hDay, _sort: hMon*100+hDay });
        } else if (hMon === currentMonth && hYear === thisYear) {
          monthlyDeliveries.push({ name: r.displayName, date: (hMon+1)+'/'+hDay, _sort: hMon*100+hDay });
        }
      }
    });

    sortAndClean(startedProperties);
    sortAndClean(monthlyStarts);
    sortAndClean(deliveredProperties);
    sortAndClean(monthlyDeliveries);

    // ── 4. Google Tasks（未完了タスク）──────────────────────────────
    var gTasks = [];
    try {
      var taskLists = Tasks.Tasklists.list();
      if (taskLists.items) {
        taskLists.items.forEach(function(tl) {
          var result = Tasks.Tasks.list(tl.id, {
            showCompleted: false,
            showHidden: false,
            maxResults: 100
          });
          if (result.items) {
            result.items.forEach(function(task) {
              if (!task.title) return;
              var obj = { title: task.title, id: task.id, listId: tl.id };
              if (task.due) {
                // GAS Tasks APIはdue を RFC3339 で返す（例: 2026-04-10T00:00:00.000Z）
                var dueDate = new Date(task.due);
                obj.due = Utilities.formatDate(dueDate, 'Asia/Tokyo', 'yyyy-MM-dd');
              }
              // task.notes にdeadlineが書いてある場合（オプション）
              if (task.notes) {
                var dlMatch = task.notes.match(/deadline[:\s]*(\d{4}-\d{2}-\d{2})/i);
                if (dlMatch) obj.deadline = dlMatch[1];
              }
              gTasks.push(obj);
            });
          }
        });
      }
      // 期日順ソート（期日なしは末尾）
      gTasks.sort(function(a, b) {
        if (!a.due && !b.due) return 0;
        if (!a.due) return 1;
        if (!b.due) return -1;
        return a.due < b.due ? -1 : a.due > b.due ? 1 : 0;
      });
    } catch(taskErr) {
      Logger.log('⚠ Google Tasks取得エラー: ' + taskErr.message);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success:        true,
      timestamp:      timestamp,
      calError:       calErrorMsg,
      scheduleEvents: scheduleEvents,
      googleTasks:    gTasks,
      notionProperties: allNotionProperties,
      calendarData: {
        startedProperties:    startedProperties,
        monthlyStarts:        monthlyStarts,
        deliveredProperties:  deliveredProperties,
        monthlyDeliveries:    monthlyDeliveries,
        yearlyStartCounts:    yearlyStartCounts,
        yearlyDeliveryCounts: yearlyDeliveryCounts
      }
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('❌ getDashboardData エラー: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false, message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
