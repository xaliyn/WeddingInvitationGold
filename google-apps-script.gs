/**
 * Хуримын урилгын хариуг Google Sheet-д хадгалах скрипт.
 * Sheet → Extensions → Apps Script дээр энэ кодыг хуулж тавиад,
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone) болгоно.
 *
 * Баганууд: A Овог, нэр | B Утасны дугаар | C Ирэх эсэх | D Том хүн | E Хүүхэд | F Ерөөл | G Огноо
 */
const SHEET_NAME = 'Sheet1';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const coming = d.attend === 'yes';
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    sheet.appendRow([
      text_(d.name, 100),
      "'" + String(d.phone || '').replace(/[^\d+ ]/g, '').slice(0, 20), // тэргүүлэх 0-ийг хадгална
      coming ? 'Ирнэ' : 'Ирэхгүй',
      coming ? clamp_(d.adults) : 0,
      coming ? clamp_(d.kids) : 0,
      text_(d.wish, 600),
      new Date(),
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Sheet нь "=", "+", "-", "@"-аар эхэлсэн утгыг томъёо гэж уншдаг тул текст болгоно.
function text_(v, max) {
  const s = String(v || '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function clamp_(n) {
  n = Math.floor(Number(n) || 0);
  return Math.max(0, Math.min(20, n));
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
