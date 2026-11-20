// SANGEETHA'S MAKEUP ARTIST - Google Apps Script backend
// Keep this file private. Do not upload Code.gs or API keys to public GitHub.

const CONFIG = {
  TIMEZONE: 'Asia/Kolkata',
  TITHE_PERCENTAGE: 1,
  API_KEY_PROPERTY: 'API_KEY'
};

const HEADERS = {
  'Bookings': ['BookingID','EventDate','Customer','Phone','Event','Time','Location','Notes','Total','Advance','Balance','Sangeetha','Rizwana','Tithe','PaymentStatus','Confirmation','Created','Updated'],
  'Payments': ['PaymentID','BookingID','Date','Amount'],
  'Tithe Records': ['BookingID','EventDate','Customer','BookingAmount','Tithe'],
  'Expenses': ['ExpenseID','Date','Category','Description','Amount','Created']
};

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    checkKey_(p.key);
    if ((p.action || 'getData') !== 'getData') return json_({success:false,error:'Unsupported GET action'});
    return json_({success:true,data:getAllData_()});
  } catch (err) {
    return json_({success:false,error:err.message || String(err)});
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkKey_(body.key);
    let result;
    switch (body.action) {
      case 'saveBooking': result = saveBooking_(body.booking || {}); break;
      case 'deleteBooking': result = deleteBooking_(String(body.bookingId || '')); break;
      case 'addPayment': result = addPayment_(body.payment || {}); break;
      case 'saveExpense': result = saveExpense_(body.expense || {}); break;
      default: throw new Error('Unknown action');
    }
    return json_({success:true,data:result});
  } catch (err) {
    return json_({success:false,error:err.message || String(err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the Apps Script project from the Google Sheet, then run setup().');
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = HEADERS[name];
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
    else {
      const existing = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getDisplayValues()[0];
      headers.forEach((h,i) => {
        if (!existing[i]) sh.getRange(1,i+1).setValue(h);
      });
    }
  });
  // Preserve date/time as text to avoid timezone shifts in date-only and time-only fields.
  getSheet_('Bookings').getRange('B:B').setNumberFormat('@');
  getSheet_('Bookings').getRange('F:F').setNumberFormat('@');
  getSheet_('Payments').getRange('C:C').setNumberFormat('@');
  getSheet_('Tithe Records').getRange('B:B').setNumberFormat('@');
  getSheet_('Expenses').getRange('B:B').setNumberFormat('@');
  Logger.log('Setup complete. Existing rows were not cleared.');
}

function getAllData_() {
  const bookings = readObjects_('Bookings').map(r => ({
    bookingId:String(r.BookingID || ''),
    date:normalizeDate_(r.EventDate),
    customer:String(r.Customer || ''),
    phone:String(r.Phone || ''),
    event:String(r.Event || ''),
    time:normalizeTime_(r.Time),
    location:String(r.Location || ''),
    notes:String(r.Notes || ''),
    total:number_(r.Total),
    advance:number_(r.Advance),
    balance:number_(r.Balance),
    sangeetha:number_(r.Sangeetha),
    rizwana:number_(r.Rizwana),
    tithe:number_(r.Tithe),
    paymentStatus:String(r.PaymentStatus || 'Pending'),
    confirmation:String(r.Confirmation || 'Pending'),
    created:String(r.Created || ''),
    updated:String(r.Updated || '')
  })).filter(x => x.bookingId);
  const payments = readObjects_('Payments').map(r => ({
    paymentId:String(r.PaymentID || ''),
    bookingId:String(r.BookingID || ''),
    date:normalizeDate_(r.Date),
    amount:number_(r.Amount)
  })).filter(x => x.paymentId);
  const titheRecords = readObjects_('Tithe Records').map(r => ({
    bookingId:String(r.BookingID || ''),
    date:normalizeDate_(r.EventDate),
    customer:String(r.Customer || ''),
    bookingAmount:number_(r.BookingAmount),
    tithe:number_(r.Tithe)
  })).filter(x => x.bookingId);
  const expenses = readObjects_('Expenses').map(r => ({
    expenseId:String(r.ExpenseID || ''),
    date:normalizeDate_(r.Date),
    category:String(r.Category || ''),
    description:String(r.Description || ''),
    amount:number_(r.Amount)
  })).filter(x => x.expenseId);
  return {bookings:bookings,payments:payments,titheRecords:titheRecords,expenses:expenses};
}

function saveBooking_(b) {
  const id = String(b.bookingId || '').trim() || newId_('BKG');
  const date = normalizeDate_(b.date);
  const time = normalizeTime_(b.time);
  const customer = String(b.customer || '').trim();
  const phone = String(b.phone || '').trim();
  const event = String(b.event || '').trim();
  const total = number_(b.total);
  const advance = number_(b.advance);
  const sangeetha = number_(b.sangeetha);
  const rizwana = number_(b.rizwana);
  if (!date) throw new Error('Please enter a valid event date.');
  if (!time) throw new Error('Please enter a valid event time.');
  if (!customer) throw new Error('Customer name is required.');
  if (!phone) throw new Error('Phone number is required.');
  if (!event) throw new Error('Event type is required.');
  if (total < 0 || advance < 0 || advance > total) throw new Error('Check total and advance amounts. Advance cannot exceed total.');
  if (sangeetha < 0 || rizwana < 0) throw new Error('Staff allocation cannot be negative.');
  const balance = round2_(total - advance);
  const tithe = round2_(total * CONFIG.TITHE_PERCENTAGE / 100);
  const now = new Date();
  const existing = findObject_('Bookings','BookingID',id);
  const created = existing ? existing.Created : Utilities.formatDate(now,CONFIG.TIMEZONE,'yyyy-MM-dd HH:mm:ss');
  const row = {
    BookingID:id, EventDate:date, Customer:customer, Phone:phone, Event:event,
    Time:time, Location:String(b.location || '').trim(), Notes:String(b.notes || '').trim(),
    Total:total, Advance:advance, Balance:balance, Sangeetha:sangeetha, Rizwana:rizwana,
    Tithe:tithe, PaymentStatus:balance <= 0 ? 'Paid' : (advance > 0 ? 'Partially Paid' : 'Pending'),
    Confirmation:String(b.confirmation || 'Pending'), Created:created,
    Updated:Utilities.formatDate(now,CONFIG.TIMEZONE,'yyyy-MM-dd HH:mm:ss')
  };
  upsertObject_('Bookings','BookingID',id,row);
  upsertObject_('Tithe Records','BookingID',id,{
    BookingID:id,EventDate:date,Customer:customer,BookingAmount:total,Tithe:tithe
  });
  return {bookingId:id,balance:balance,tithe:tithe};
}

function addPayment_(p) {
  const bookingId = String(p.bookingId || '').trim();
  const date = normalizeDate_(p.date);
  const amount = number_(p.amount);
  if (!bookingId || !date || amount <= 0) throw new Error('Select a booking, date and valid payment amount.');
  const booking = findObject_('Bookings','BookingID',bookingId);
  if (!booking) throw new Error('Booking not found.');
  const currentBalance = number_(booking.Balance);
  if (amount > currentBalance) throw new Error('Payment cannot be greater than the remaining balance.');
  const paymentId = newId_('PAY');
  appendObject_('Payments',{PaymentID:paymentId,BookingID:bookingId,Date:date,Amount:amount});
  const newAdvance = round2_(number_(booking.Advance) + amount);
  const newBalance = round2_(number_(booking.Total) - newAdvance);
  booking.Advance = newAdvance;
  booking.Balance = newBalance;
  booking.PaymentStatus = newBalance <= 0 ? 'Paid' : 'Partially Paid';
  booking.Updated = Utilities.formatDate(new Date(),CONFIG.TIMEZONE,'yyyy-MM-dd HH:mm:ss');
  upsertObject_('Bookings','BookingID',bookingId,booking);
  return {paymentId:paymentId,balance:newBalance};
}

function saveExpense_(x) {
  const date = normalizeDate_(x.date);
  const amount = number_(x.amount);
  if (!date || amount <= 0) throw new Error('Enter a valid expense date and amount.');
  const id = String(x.expenseId || '').trim() || newId_('EXP');
  const row = {ExpenseID:id,Date:date,Category:String(x.category || 'Other'),Description:String(x.description || ''),Amount:amount,Created:Utilities.formatDate(new Date(),CONFIG.TIMEZONE,'yyyy-MM-dd HH:mm:ss')};
  appendObject_('Expenses',row);
  return {expenseId:id};
}

function deleteBooking_(id) {
  if (!id) throw new Error('Booking ID is required.');
  deleteByValue_('Bookings','BookingID',id);
  deleteByValue_('Tithe Records','BookingID',id);
  // Keep payment history as an audit trail. Payments for this ID remain visible as orphaned records.
  return {deleted:true,bookingId:id};
}

function checkKey_(provided) {
  const expected = PropertiesService.getScriptProperties().getProperty(CONFIG.API_KEY_PROPERTY);
  if (!expected) throw new Error('API_KEY is not configured in Script Properties.');
  if (!provided || String(provided) !== String(expected)) throw new Error('Access denied. Check the API key in Settings.');
}
function getSS_(){const ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw new Error('Spreadsheet not found.');return ss;}
function getSheet_(name){const sh=getSS_().getSheetByName(name);if(!sh)throw new Error('Missing sheet: '+name+'. Run setup() first.');return sh;}
function readObjects_(name){const sh=getSheet_(name),values=sh.getDataRange().getValues();if(values.length<2)return[];const headers=values[0].map(String);return values.slice(1).filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>{const o={};headers.forEach((h,i)=>o[h]=r[i]);return o;});}
function findObject_(name,keyHeader,keyValue){return readObjects_(name).find(r=>String(r[keyHeader])===String(keyValue))||null;}
function upsertObject_(name,keyHeader,keyValue,obj){const sh=getSheet_(name),values=sh.getDataRange().getValues(),headers=values[0].map(String),row=headers.map(h=>obj[h]===undefined?'':obj[h]);let rowNo=0;for(let i=1;i<values.length;i++){if(String(values[i][headers.indexOf(keyHeader)])===String(keyValue)){rowNo=i+1;break;}}if(rowNo)sh.getRange(rowNo,1,1,headers.length).setValues([row]);else sh.appendRow(row);}
function appendObject_(name,obj){const sh=getSheet_(name),headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);sh.appendRow(headers.map(h=>obj[h]===undefined?'':obj[h]));}
function deleteByValue_(name,header,value){const sh=getSheet_(name),v=sh.getDataRange().getValues(),headers=v[0].map(String),idx=headers.indexOf(header);for(let i=v.length-1;i>=1;i--)if(String(v[i][idx])===String(value))sh.deleteRow(i+1);}
function normalizeDate_(v){
  if(v===null||v===undefined||v==='')return '';
  if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,CONFIG.TIMEZONE,'yyyy-MM-dd');
  const s=String(v).trim();
  let m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if(m)return validDate_(+m[1],+m[2],+m[3]);
  m=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if(m)return validDate_(+m[3],+m[2],+m[1]);
  return '';
}
function validDate_(y,m,d){const dt=new Date(Date.UTC(y,m-1,d));if(dt.getUTCFullYear()!==y||dt.getUTCMonth()!==m-1||dt.getUTCDate()!==d)return '';return String(y).padStart(4,'0')+'-'+String(m).padStart(2,'0')+'-'+String(d).padStart(2,'0');}
function normalizeTime_(v){
  if(v===null||v===undefined||v==='')return '';
  if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,CONFIG.TIMEZONE,'HH:mm');
  if(typeof v==='number'&&v>=0&&v<1){const mins=Math.round(v*1440);return String(Math.floor(mins/60)%24).padStart(2,'0')+':'+String(mins%60).padStart(2,'0');}
  const s=String(v).trim();
  let m=s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if(m){const h=+m[1],min=+m[2];if(h>23||min>59)return '';return String(h).padStart(2,'0')+':'+String(min).padStart(2,'0');}
  m=s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if(m){let h=+m[1],min=+m[2];if(h<1||h>12||min>59)return '';h=h%12+(m[3].toUpperCase()==='PM'?12:0);return String(h).padStart(2,'0')+':'+String(min).padStart(2,'0');}
  return '';
}
function number_(v){const n=Number(v);return Number.isFinite(n)?n:0;}
function round2_(n){return Math.round((n+Number.EPSILON)*100)/100;}
function newId_(prefix){return prefix+'-'+Utilities.formatDate(new Date(),CONFIG.TIMEZONE,'yyMMdd-HHmmss')+'-'+Math.floor(Math.random()*900+100);}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
