const SHEET_ID = '1HuCg5keyzs4BQIa3ZjzEyEVB54j-SwtGarFP5uq48Pk';
const TAB_NAME = 'מנויים';
const HEADERS = ['שם השירות','קטגוריה','סכום לתשלום','מטבע','תדירות','תאריך החיוב הבא','תאריך סיום','אמצעי תשלום','מצב','הערות'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('המנויים שלי');
}
function sheet_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_NAME);
  if (!sheet) throw new Error('לא נמצאה לשונית בשם ״מנויים״.');
  const headers = sheet.getRange(1, 1, 1, 10).getValues()[0];
  if (HEADERS.some((h, i) => h !== headers[i])) throw new Error('כותרות הגיליון אינן תואמות לקובץ ״מנויים ותשלומים״.');
  return sheet;
}
function asDate_(value, timezone) {
  return value instanceof Date && !isNaN(value.getTime())
    ? Utilities.formatDate(value, timezone, 'yyyy-MM-dd')
    : String(value || '');
}
function record_(cells, row, timezone) {
  return {row: row, name:String(cells[0]||''), category:String(cells[1]||''), amount:Number(cells[2])||0,
    currency:String(cells[3]||''), frequency:String(cells[4]||''), nextDate:asDate_(cells[5], timezone),
    endDate:asDate_(cells[6], timezone), payment:String(cells[7]||''), status:String(cells[8]||''), notes:String(cells[9]||'')};
}
function getSubscriptions() {
  const sheet = sheet_();
  if (sheet.getLastRow() < 2) return [];
  const timezone = sheet.getParent().getSpreadsheetTimeZone();
  return sheet.getRange(2, 1, sheet.getLastRow()-1, 10).getValues()
    .map((cells, i) => record_(cells, i+2, timezone)).filter(r => r.name);
}
function parseDate_(value) {
  if (!value) return '';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error('תאריך לא תקין.');
  const [y,m,d] = value.split('-').map(Number);
  const date = new Date(y,m-1,d);
  if (date.getFullYear() !== y || date.getMonth() !== m-1 || date.getDate() !== d) throw new Error('תאריך לא תקין.');
  return date;
}
function text_(value, max) {
  const text = String(value || '').trim().slice(0,max);
  return /^[=+@-]/.test(text) ? "'"+text : text;
}
function cells_(r) {
  const amount = Number(r.amount);
  if (!String(r.name||'').trim() || !Number.isFinite(amount) || amount < 0) throw new Error('יש להזין שם וסכום תקינים.');
  if (!['שקל','דולר','אירו'].includes(r.currency) || !['חודשי','שנתי','חד פעמי'].includes(r.frequency) || !['פעיל','מושהה','בוטל'].includes(r.status)) throw new Error('יש לבחור ערכים מהרשימות.');
  const next = parseDate_(r.nextDate), end = parseDate_(r.endDate);
  if (!next) throw new Error('יש להזין תאריך חיוב.');
  if (end && end < next) throw new Error('תאריך הסיום מוקדם מתאריך החיוב.');
  return [text_(r.name,100),text_(r.category,100),amount,r.currency,r.frequency,next,end,text_(r.payment,100),r.status,text_(r.notes,500)];
}
function saveSubscription(record, expected) {
  const lock = LockService.getScriptLock();lock.waitLock(10000);
  try {
    const sheet = sheet_(), row = Number(record.row) || 0, cells = cells_(record);
    if (row) {
      if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) throw new Error('הרשומה אינה קיימת. רענן את הדף.');
      const current = record_(sheet.getRange(row,1,1,10).getValues()[0],row,sheet.getParent().getSpreadsheetTimeZone());
      if (JSON.stringify(current) !== JSON.stringify(expected)) throw new Error('הרשומה השתנתה בגיליון. רענן לפני השמירה.');
      sheet.getRange(row,1,1,10).setValues([cells]);
      sheet.getRange(row,6,1,2).setNumberFormat('dd/mm/yyyy');
    } else {
      const last = sheet.getLastRow(), target = last+1;
      sheet.getRange(target,1,1,10).setValues([cells]);
      sheet.getRange(target,6,1,2).setNumberFormat('dd/mm/yyyy');
    }
    SpreadsheetApp.flush();return getSubscriptions();
  } finally {lock.releaseLock()}
}
function deleteSubscription(expected) {
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const sheet=sheet_(), row=Number(expected && expected.row);
    if (!Number.isInteger(row)||row<2||row>sheet.getLastRow()) throw new Error('הרשומה אינה קיימת. רענן את הדף.');
    const current=record_(sheet.getRange(row,1,1,10).getValues()[0],row,sheet.getParent().getSpreadsheetTimeZone());
    if (JSON.stringify(current)!==JSON.stringify(expected)) throw new Error('הרשומה השתנתה בגיליון. רענן לפני המחיקה.');
    sheet.deleteRow(row);SpreadsheetApp.flush();return getSubscriptions();
  } finally {lock.releaseLock()}
}
