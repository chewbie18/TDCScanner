/** TUP Smart OMR — Google Sheets receiver.
 * 1. Create Google Sheet, Extensions > Apps Script, paste this code.
 * 2. Set SECRET to a long random passphrase (same as in app Settings).
 * 3. Deploy > New deployment > Web app, Execute as Me, access Anyone.
 * 4. Paste /exec URL in app. Restrict spreadsheet sharing to authorized teachers.
 * WARNING: Shared-secret + publicly callable endpoint is NOT enterprise authentication.
 */
const SECRET = 'CHANGE_THIS_TO_A_LONG_RANDOM_PRIVATE_PASSPHRASE';
const TAB = 'Scores';
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (!payload || payload.secret !== SECRET || SECRET.startsWith('CHANGE_THIS')) throw Error('Unauthorized');
    const r = payload.result || {};
    if (!r.id || !r.name || !r.examId || !Number.isFinite(Number(r.score))) throw Error('Invalid record');
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const book = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = book.getSheetByName(TAB) || book.insertSheet(TAB);
      const header = ['Record ID','Student name','Year & section','Date','Subject','Professor','Examination','Exam ID','Score','Maximum','Percent','Checked at'];
      if (sheet.getLastRow() === 0) { sheet.appendRow(header); sheet.setFrozenRows(1); }
      const id = String(r.id).slice(0,120);
      const ids = sheet.getLastRow()>1 ? sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat() : [];
      if (ids.includes(id)) return ContentService.createTextOutput('duplicate');
      // Leading apostrophe forces spreadsheet text, protecting against formula injection.
      const safe = x => { const s = String(x ?? '').slice(0,300); return /^[=+\-@\t\r]/.test(s) ? "'"+s : s; };
      const max = Number(r.max), score=Number(r.score);
      sheet.appendRow([id,safe(r.name),safe(r.section),safe(r.date),safe(r.subject),safe(r.professor),safe(r.exam),safe(r.examId),score,max,max>0?score/max:0,safe(r.created)]);
      sheet.getRange(sheet.getLastRow(),11).setNumberFormat('0.00%');
      return ContentService.createTextOutput('ok');
    } finally { lock.releaseLock(); }
  } catch (err) { return ContentService.createTextOutput('error: '+err.message); }
}
