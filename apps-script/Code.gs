/**
 * 생수 정기배송 수령 확인 — Apps Script (참고용 전체 코드)
 *
 * ⚠ 이 파일은 저장소에 보관된 «참고 코드»이며 자동으로 적용되지 않습니다.
 *   Apps Script 에 붙여넣고 «새 버전»으로 다시 배포해야 사진 첨부·메일 알림이 켜집니다.
 *   기존 스크립트를 교체하기 전에 반드시 기존 코드를 따로 복사해 두세요.
 *   (이 코드는 Apps Script 환경이 아닌 곳에서는 실행해 볼 수 없어, 배포 후 README 의 점검표로 확인이 필요합니다.)
 *
 * 하는 일
 *   1) doGet  ?key=...            → 저장된 값 읽기  {value}
 *      doGet  ?action=features    → 지원 기능 알림  {features:{photos:true}}   ← 이 응답이 있어야 웹페이지에 «사진 추가»가 나타납니다
 *   2) doPost {key,value}         → 값 저장 (키 형식·크기 검사)
 *      doPost {action:'upload'…}  → 사진을 Drive 폴더에 저장하고 파일 id 반환
 *   3) sendReminders()            → 지연된 기관을 메일로 알림 (시간 기반 트리거로 매일 실행)
 */
const SHEET_NAME = '_WEBDATA';                 // 기존 시트 탭 이름과 같아야 합니다
const KEY_PATTERN = /^pulmuone-water:\d{4}-\d{2}$/;
const MAX_VALUE = 49000;                       // 구글 시트 한 칸 5만 자 제한
const PHOTO_FOLDER = '풀무원샘물_배송사진';
const MAX_PHOTO_B64 = 6 * 1024 * 1024;         // 사진 1장 상한(base64)
const REMIND = {recv: 3, reply: 2, resolve: 5}; // index.html 의 ALERT_DAYS 와 같은 기준(일)

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(['key', 'value', 'updatedAt']); }
  return sh;
}
function rowOf_(sh, key) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const keys = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) if (keys[i][0] === key) return i + 2;
  return 0;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'features') return json_({ ok: true, features: { photos: true } });
  const key = p.key;
  if (!key || !KEY_PATTERN.test(key)) return json_({});
  const sh = sheet_(), r = rowOf_(sh, key);
  return json_(r ? { value: String(sh.getRange(r, 2).getValue()) } : {});
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  if (body.action === 'upload') return upload_(body);

  const key = body.key, value = body.value;
  if (!KEY_PATTERN.test(String(key))) return json_({ ok: false, error: 'bad key' });
  if (typeof value !== 'string' || value.length > MAX_VALUE) return json_({ ok: false, error: 'bad value' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);                         // 동시에 저장해도 서로 섞이지 않도록 한 번에 하나씩 처리
  try {
    const sh = sheet_(), r = rowOf_(sh, key), now = new Date();
    if (r) sh.getRange(r, 2, 1, 2).setValues([[value, now]]);
    else sh.appendRow([key, value, now]);
  } finally { lock.releaseLock(); }
  return json_({ ok: true });
}

function upload_(b) {
  if (!KEY_PATTERN.test(String(b.key)) || typeof b.data !== 'string' || b.data.length > MAX_PHOTO_B64)
    return json_({ ok: false, error: 'bad upload' });
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(PHOTO_FOLDER);
  const name = String(b.key).replace('pulmuone-water:', '') + '_기관' + Number(b.no) + '_' + (b.name || 'photo');
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b.data), 'image/jpeg', name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);   // 링크를 아는 사람만 볼 수 있음
  return json_({ ok: true, id: file.getId() });
}

/* ── 메일 알림: 스크립트 속성 REMIND_TO(쉼표로 구분한 메일 주소)에 매일 오전 발송 ── */
function sendReminders() {
  const to = PropertiesService.getScriptProperties().getProperty('REMIND_TO');
  if (!to) return;
  const t = new Date(), m = t.getMonth() + 1, key = 'pulmuone-water:' + t.getFullYear() + '-' + ('0' + m).slice(-2);
  const sh = sheet_(), r = rowOf_(sh, key);
  if (!r) return;
  let rows; try { const j = JSON.parse(sh.getRange(r, 2).getValue()); rows = j.rows || j; } catch (err) { return; }
  const today = new Date(t.getFullYear(), t.getMonth(), t.getDate());
  const days = iso => (/^\d{4}-\d{2}-\d{2}/.test(iso || '')) ? Math.floor((today - new Date(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10))) / 86400000) : null;
  const lines = [];
  Object.keys(rows).forEach(no => {
    const d = rows[no] || {}, iss = !!d.issue || d.status === 'miss' || d.status === 'part';
    let n;
    if (d.shipStart && (d.status || 'none') === 'none' && (n = days(d.shipDate)) !== null && n >= REMIND.recv) lines.push(no + '번 기관: 배송 시작 ' + n + '일째 수령 확인 없음');
    if (iss && !String(d.pmNote || '').trim() && (n = days(String(d.atFac || '').slice(0, 10))) !== null && n >= REMIND.reply) lines.push(no + '번 기관: 이슈 등록 ' + n + '일째 조치 회신 없음');
    if (iss && String(d.pmNote || '').trim() && !d.resolved && (n = days(String(d.atPM || '').slice(0, 10))) !== null && n >= REMIND.resolve) lines.push(no + '번 기관: 조치 회신 ' + n + '일째 해결 확인 없음');
  });
  if (!lines.length) return;
  MailApp.sendEmail(to, '[생수 정기배송] ' + m + '월 확인 필요 ' + lines.length + '건',
    lines.join('\n') + '\n\n확인: https://pulmuonewater.github.io/');
}
