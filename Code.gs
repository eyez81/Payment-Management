const SHEET_ID = '1HuCg5keyzs4BQIa3ZjzEyEVB54j-SwtGarFP5uq48Pk';
const TAB_NAME = 'מנויים';
const HEADERS = ['שם השירות','קטגוריה','סכום לתשלום','מטבע','תדירות','תאריך החיוב הבא','תאריך סיום','אמצעי תשלום','מצב','הערות','קישור','שייך ל'];
const LIST_HEADERS = ['קטגוריות','אמצעי תשלום','שייך ל'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('המנויים שלי');
}
function sheet_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_NAME);
  if (!sheet) throw new Error('לא נמצאה לשונית בשם ״מנויים״.');
  const headers = sheet.getRange(1, 1, 1, 10).getValues()[0];
  if (HEADERS.slice(0, 10).some((h, i) => h !== headers[i])) throw new Error('כותרות הגיליון אינן תואמות לקובץ ״מנויים ותשלומים״.');
  const linkHeader = sheet.getRange(1, 11);
  if (!linkHeader.getValue()) {
    linkHeader.setValue('קישור');
    linkHeader.setBackground('#183349').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(11, 260);
  } else if (linkHeader.getValue() !== 'קישור') {
    throw new Error('עמודה K כבר בשימוש. יש לפנות אותה לפני הוספת עמודת קישור.');
  }
  const ownerHeader = sheet.getRange(1, 12);
  if (!ownerHeader.getValue()) {
    ownerHeader.setValue('שייך ל');
    ownerHeader.setBackground('#183349').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(12, 140);
  } else if (ownerHeader.getValue() !== 'שייך ל') {
    throw new Error('עמודה L כבר בשימוש. יש לפנות אותה לפני הוספת עמודת בעל התשלום.');
  }
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
    endDate:asDate_(cells[6], timezone), payment:String(cells[7]||''), status:String(cells[8]||''), notes:String(cells[9]||''), link:String(cells[10]||''), owner:String(cells[11]||'')};
}
function getSubscriptions() {
  const sheet = sheet_();
  if (sheet.getLastRow() < 2) return [];
  const timezone = sheet.getParent().getSpreadsheetTimeZone();
  return sheet.getRange(2, 1, sheet.getLastRow()-1, 12).getValues()
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
  const link = String(r.link || '').trim();
  if (link && !/^https?:\/\/[^\s]+$/i.test(link)) throw new Error('הקישור חייב להתחיל ב־https:// או http://.');
  return [text_(r.name,100),text_(r.category,100),amount,r.currency,r.frequency,next,end,text_(r.payment,100),r.status,text_(r.notes,500),link,text_(r.owner,100)];
}
function saveSubscription(record, expected) {
  const lock = LockService.getScriptLock();lock.waitLock(10000);
  try {
    const sheet = sheet_(), row = Number(record.row) || 0, cells = cells_(record);
    if (row) {
      if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) throw new Error('הרשומה אינה קיימת. רענן את הדף.');
      const current = record_(sheet.getRange(row,1,1,12).getValues()[0],row,sheet.getParent().getSpreadsheetTimeZone());
      if (JSON.stringify(current) !== JSON.stringify(expected)) throw new Error('הרשומה השתנתה בגיליון. רענן לפני השמירה.');
      sheet.getRange(row,1,1,12).setValues([cells]);
      sheet.getRange(row,6,1,2).setNumberFormat('dd/mm/yyyy');
    } else {
      const last = sheet.getLastRow(), target = last+1;
      sheet.getRange(target,1,1,12).setValues([cells]);
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
    const current=record_(sheet.getRange(row,1,1,12).getValues()[0],row,sheet.getParent().getSpreadsheetTimeZone());
    if (JSON.stringify(current)!==JSON.stringify(expected)) throw new Error('הרשומה השתנתה בגיליון. רענן לפני המחיקה.');
    sheet.deleteRow(row);SpreadsheetApp.flush();return getSubscriptions();
  } finally {lock.releaseLock()}
}

function listSheet_() {
  const book = SpreadsheetApp.openById(SHEET_ID);
  let sheet = book.getSheetByName('רשימות');
  if (!sheet) sheet = book.insertSheet('רשימות');
  const current = sheet.getRange(1, 1, 1, 3).getValues()[0];
  if (current.every(v => !v)) {
    sheet.getRange(1, 1, 1, 3).setValues([LIST_HEADERS]);
    sheet.getRange(1, 1, 1, 3).setBackground('#183349').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (LIST_HEADERS.some((h, i) => current[i] !== h)) {
    throw new Error('הלשונית ״רשימות״ כבר קיימת במבנה אחר.');
  }
  return sheet;
}
function getOptions() {
  const sheet = listSheet_(), options = {category:[], payment:[], owner:[]};
  const keys = ['category','payment','owner'];
  if (sheet.getLastRow() > 1) {
    const values = sheet.getRange(2, 1, sheet.getLastRow()-1, 3).getValues();
    values.forEach(row => row.forEach((v, i) => {if (v) options[keys[i]].push(String(v));}));
  }
  getSubscriptions().forEach(r => keys.forEach(key => {if (r[key]) options[key].push(r[key]);}));
  keys.forEach(key => options[key] = [...new Set(options[key])].sort((a,b) => a.localeCompare(b,'he')));
  return options;
}
function addOption(type, rawValue) {
  const column = {category:1, payment:2, owner:3}[type];
  if (!column) throw new Error('סוג רשימה לא תקין.');
  const value = text_(rawValue, 100);
  if (!value) throw new Error('יש לכתוב שם לרשימה.');
  const lock = LockService.getScriptLock();lock.waitLock(10000);
  try {
    const sheet = listSheet_();
    const last = sheet.getLastRow();
    const existing = last > 1 ? sheet.getRange(2,column,last-1,1).getValues().flat().map(String) : [];
    if (!existing.includes(value)) sheet.getRange(last+1,column).setValue(value);
    return getOptions();
  } finally {lock.releaseLock()}
}
function getRates() {
  const cache = CacheService.getScriptCache(), cached = cache.get('boi-rates-v1');
  if (cached) return JSON.parse(cached);
  try {
    const response = UrlFetchApp.fetch('https://boi.org.il/PublicApi/GetExchangeRates', {muteHttpExceptions:true});
    if (response.getResponseCode() !== 200) throw new Error('שירות השערים אינו זמין.');
    const body = JSON.parse(response.getContentText());
    const result = {USD:null, EUR:null, updated:null};
    for (const key of ['USD','EUR']) {
      const entry = (body.exchangeRates || []).find(r => r.key === key);
      const rate = Number(entry && entry.currentExchangeRate), unit = Number(entry && entry.unit || 1);
      if (!Number.isFinite(rate) || rate <= 0 || !Number.isFinite(unit) || unit <= 0) throw new Error('נתוני השערים חסרים.');
      result[key] = rate / unit;
      if (entry.lastUpdate && (!result.updated || entry.lastUpdate > result.updated)) result.updated = entry.lastUpdate;
    }
    cache.put('boi-rates-v1', JSON.stringify(result), 3600);
    return result;
  } catch (error) {
    return {USD:null, EUR:null, updated:null, error:'לא ניתן לטעון כרגע שער יציג מבנק ישראל.'};
  }
}
