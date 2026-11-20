/**
 * 登記日網站的後端：把 RSVP 和祝福存進這份 Google 試算表。
 * 貼到「擴充功能 → Apps Script」，然後部署成網頁應用程式（見 設定教學.md）。
 *
 * 會自動建立兩個工作表：
 *   RSVP ：時間｜名字｜出席｜人數｜飲食需求｜備註
 *   祝福 ：時間｜署名｜祝福｜歌曲｜連結｜顯示｜為什麼選這首
 *   設定 ：B1 放 Spotify「邀請協作者」連結（7 天失效，過期再換新的貼上）
 * 想隱藏某則留言，把那一列「顯示」欄改成 N 即可。
 */

const RSVP_SHEET = 'RSVP';
const BLESS_SHEET = '祝福';
const CONFIG_SHEET = '設定';
const BLESS_HEADERS = ['時間', '署名', '祝福', '歌曲', '連結', '顯示', '為什麼選這首'];

function blessSheet_() {
  const sh = sheet_(BLESS_SHEET, BLESS_HEADERS);
  if (sh.getRange(1, 7).getValue() === '') sh.getRange(1, 7).setValue('為什麼選這首');
  return sh;
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function clip_(v, n) {
  v = String(v == null ? '' : v).trim();
  // 避免被試算表當成公式
  if (/^[=+\-@]/.test(v)) v = "'" + v;
  return v.slice(0, n);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const now = new Date();
    if (d.type === 'rsvp') {
      sheet_(RSVP_SHEET, ['時間', '名字', '出席', '人數', '飲食需求', '備註'])
        .appendRow([now, clip_(d.name, 40), clip_(d.attend, 20), Number(d.count) || 0,
                    clip_(d.diet, 60), clip_(d.note, 200)]);
    } else if (d.type === 'blessing') {
      if (!clip_(d.name, 40) || !clip_(d.message, 600)) throw new Error('empty');
      blessSheet_()
        .appendRow([now, clip_(d.name, 40), clip_(d.message, 600), clip_(d.song, 80),
                    clip_(d.link, 300), 'Y', clip_(d.why, 300)]);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (e.parameter.action === 'config') {
    const cs = sheet_(CONFIG_SHEET, ['Spotify 協作連結（貼在 B1）', '']);
    const url = String(cs.getRange(1, 2).getValue()).trim();
    return json_({ collabUrl: /^https:\/\/open\.spotify\.com\//.test(url) ? url : '' });
  }
  if (e.parameter.action !== 'blessings') return json_([]);
  const sh = blessSheet_();
  const rows = sh.getDataRange().getValues().slice(1);
  const list = rows
    .filter(r => String(r[5]).toUpperCase() !== 'N' && r[2])
    .reverse()   // 最新的在最上面
    .map(r => ({
      name: String(r[1]).replace(/^'/, ''),
      message: String(r[2]).replace(/^'/, ''),
      song: String(r[3]).replace(/^'/, ''),
      link: String(r[4]).replace(/^'/, ''),
      why: String(r[6] || '').replace(/^'/, '')
    }));
  return json_(list);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
