/**
 * gas-customer-schedule.gs
 * ===============================================================
 * お客様工程表作成ツール — GCal から日程を取得して 21項目を返す
 *
 * デプロイ後の URL を tools/customer-schedule.html の fetchScheduleDates() で叩く想定。
 *
 * 入力：?mode=getCustomerScheduleData&prop=山田T
 * 出力：{ ok: true, data: { workDates, meetingDates, publicInspDates } }
 *
 * 検索ルール（wishlist_sekou_portal.md より）：
 *   工程予定日（11個）：地縄立会い / 本体着工 / 基礎検査(旬丸め) / 建て方(旬丸め)
 *                     / 建て方 / 建て方+2週(旬丸め) / 木完検査-5日(旬丸め) / 木完検査
 *                     / 竣工検査-1週(旬丸め) / 竣工検査 / 引渡し
 *   立会い（5個）：地縄立会い / 地縄立会い(=近隣挨拶) / 構造立会い / 木完立会い / 竣工立会い
 *   公的検査（5個）：配筋検査(旬丸め) / 構造検査(旬丸め) / 雨仕舞い検査+10日(旬丸め)
 *                  / 竣工検査 / 竣工検査(=完了検査)
 *
 * 旬丸め：1〜10日=上旬 / 11〜20日=中旬 / 21日〜末日=下旬
 * ===============================================================
 */

// ==== 設定 ====
const GCAL_ID = 'primary';  // または特定のカレンダーID

function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) || '';
  if (mode === 'getCustomerScheduleData') {
    const prop = e.parameter.prop || '';
    return jsonOut(buildScheduleData(prop));
  }
  return jsonOut({ ok: false, error: 'unknown mode' });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==== メイン処理 ====
function buildScheduleData(propName) {
  if (!propName) return { ok: false, error: 'propName required' };

  // 物件の全期間（着工〜引渡し）をカバーする範囲で検索：1年前〜2年後
  const now = new Date();
  const from = new Date(now.getFullYear() - 1, 0, 1);
  const to   = new Date(now.getFullYear() + 2, 11, 31);

  const cal = CalendarApp.getCalendarById(GCAL_ID) || CalendarApp.getDefaultCalendar();
  const events = cal.getEvents(from, to);

  // 物件名で絞り込み（前方一致 / タイトル中含有 どちらでも）
  const filtered = events.filter(ev => {
    const t = ev.getTitle() || '';
    return t.indexOf(propName) >= 0;
  });

  // タイトルキーワード → 開始日 のマップを作成
  const dateMap = {};
  filtered.forEach(ev => {
    const t = ev.getTitle();
    const d = ev.getStartTime();
    KEYWORDS.forEach(kw => {
      if (t.indexOf(kw) >= 0 && !dateMap[kw]) dateMap[kw] = d;
    });
  });

  // 21項目の組み立て
  const result = {
    workDates: {
      jinawa:       fmtMD(dateMap['地縄立会い']),
      chakkou:      fmtMD(dateMap['本体着工']),
      kiso:         fmtJun(dateMap['基礎検査']),
      dodai:        fmtJun(dateMap['建て方']),
      tatekata:     fmtMD(dateMap['建て方']),
      yane:         fmtJun(addDays(dateMap['建て方'], 14)),
      gaiheki:      fmtJun(addDays(dateMap['木完検査'], -5)),
      mokkan:       fmtMD(dateMap['木完検査']),
      naibu:        fmtJun(addDays(dateMap['竣工検査'], -7)),
      shunkou:      fmtMD(dateMap['竣工検査']),
      hikiwatashi:  fmtMD(dateMap['引渡し']),
    },
    meetingDates: {
      jinawa: fmtMD(dateMap['地縄立会い']),
      kinjo:  fmtMD(dateMap['地縄立会い']),  // 近隣挨拶 = 地縄立会いと同日
      kouzou: fmtMD(dateMap['構造立会い']),
      mokkan: fmtMD(dateMap['木完立会い']),
      shunkou:fmtMD(dateMap['竣工立会い']),
    },
    publicInspDates: {
      '1':         fmtJun(dateMap['配筋検査']),
      '2':         fmtJun(dateMap['構造検査']),
      '3':         fmtJun(addDays(dateMap['雨仕舞い検査'], 10)),
      '4':         fmtMD(dateMap['竣工検査']),
      completion:  fmtMD(dateMap['竣工検査']),
    },
  };

  return { ok: true, data: result };
}

// 検索対象キーワード一覧
const KEYWORDS = [
  '地縄立会い', '本体着工', '基礎検査', '建て方', '木完検査',
  '竣工検査', '引渡し', '構造立会い', '木完立会い', '竣工立会い',
  '配筋検査', '構造検査', '雨仕舞い検査',
];

// ==== 日付ユーティリティ ====
function fmtMD(d) {
  if (!d) return '';
  const m = pad2(d.getMonth() + 1);
  const dd = pad2(d.getDate());
  return `（${m}/${dd}）`;
}

function fmtJun(d) {
  if (!d) return '';
  const m = d.getMonth() + 1;
  const dd = d.getDate();
  let label;
  if (dd <= 10) label = '上旬';
  else if (dd <= 20) label = '中旬';
  else label = '下旬';
  return `（${m}月${label}）`;
}

function addDays(d, n) {
  if (!d) return null;
  const r = new Date(d.getTime());
  r.setDate(r.getDate() + n);
  return r;
}

function pad2(n) { return n < 10 ? '0' + n : '' + n; }
