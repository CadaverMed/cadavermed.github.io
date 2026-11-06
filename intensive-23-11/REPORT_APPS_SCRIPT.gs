/**
 * CADAVER.MED — 23.11.2026
 * Google Apps Script web app. Deploy as owner, access "Anyone".
 * Receives only invitation name, code and page events; NO phone from browser.
 * Phone submissions remain in Tally until a separate approved integration is configured.
 */
const SHEET_ID = '1TZtRVy8NRfe1TtaZT3a1RgCnEKu5xzeQW89NmD7chnM';
const TAB = 'Посетители и заявки';
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents || '{}');
    const code = String(p.code || '').trim();
    const name = String(p.name || '').trim().slice(0,80);
    const event = String(p.event || '');
    if (!/^CM-261123-\d{3,8}$/.test(code) || !['visit','open','instagram'].includes(event)) return ContentService.createTextOutput('invalid');
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);
      const last = sheet.getLastRow();
      const ids = last > 1 ? sheet.getRange(2,2,last-1,1).getValues() : [];
      let row = ids.findIndex(x => x[0] === code) + 2;
      const now = new Date();
      if (row < 2) {
        row = sheet.getLastRow()+1;
        sheet.getRange(row,1,1,10).setValues([[name,code,now,now,0,0,0,'','','']]);
      }
      const v = sheet.getRange(row,1,1,10).getValues()[0];
      if (name) v[0] = name;
      v[3] = now;
      if (event === 'visit') {if (!v[2]) v[2] = now; v[4] = Number(v[4]||0)+1;}
      if (event === 'open') v[5] = Number(v[5]||0)+1;
      if (event === 'instagram') v[6] = Number(v[6]||0)+1;
      sheet.getRange(row,1,1,10).setValues([v]);
    } finally {lock.releaseLock();}
    return ContentService.createTextOutput('ok');
  } catch(err) {return ContentService.createTextOutput('error');}
}
function doGet() {return ContentService.createTextOutput('CADAVER.MED tracker ready');}
