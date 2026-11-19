// Sangeetha's Makeup Artist – Apps Script backend (bound to a Google Sheet)
const H = {
  'Bookings': ['BookingID','EventDate','Customer','Phone','Event','Time','Location','Notes','Total','Advance','Balance','Sangeetha','Rizwana','Tithe','PaymentStatus','Confirmation','Created','Updated'],
  'Payments': ['PaymentID','BookingID','Date','Amount'],
  'Tithe Records': ['BookingID','EventDate','Customer','BookingAmount','Tithe']
};
function sh(n) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(n);
  if (!s) {
    s = ss.insertSheet(n);
    s.getRange('A:Z').setNumberFormat('@');           // keep dates/phones as plain text
    s.appendRow(H[n]);
    s.setFrozenRows(1);
    s.getRange(1,1,1,H[n].length).setFontWeight('bold').setBackground('#183D32').setFontColor('#ffffff');
  }
  return s;
}
function setup() { Object.keys(H).forEach(sh); }       // run once to create the 3 tabs
function rows(n) { const v = sh(n).getDataRange().getValues(); return v.slice(1).filter(r => r[0] !== ''); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function findRow(s, id) { const c = s.getRange(1,1,s.getLastRow(),1).getValues(); for (let i=1;i<c.length;i++) if (String(c[i][0])===String(id)) return i+1; return 0; }
function upsert(s, id, row) { const r = findRow(s,id); r ? s.getRange(r,1,1,row.length).setValues([row]) : s.appendRow(row); }
function removeById(s, id, col) { const v = s.getDataRange().getValues(); for (let i=v.length-1;i>=1;i--) if (String(v[i][col])===String(id)) s.deleteRow(i+1); }

function doGet() {
  try {
    const pays = {};
    rows('Payments').forEach(p => (pays[p[1]] = pays[p[1]] || []).push({date:p[2], amount:Number(p[3])}));
    const bookings = rows('Bookings').map(r => ({
      id:r[0], date:r[1], name:r[2], phone:r[3], event:r[4], time:r[5], location:r[6], notes:r[7],
      total:Number(r[8]), advance:Number(r[9]), balance:Number(r[10]), sangeetha:Number(r[11]), rizwana:Number(r[12]),
      tithe:Number(r[13]), status:r[14], confirmation:r[15], created:r[16], payments:pays[r[0]] || []
    }));
    return out({ok:true, bookings});
  } catch (err) { return out({ok:false, error:String(err)}); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'delete') {
      removeById(sh('Bookings'), d.id, 0); removeById(sh('Payments'), d.id, 1); removeById(sh('Tithe Records'), d.id, 0);
      return out({ok:true});
    }
    if (d.action === 'save') {
      const b = d.booking, now = new Date().toISOString();
      const id = b.id || 'BK' + Date.now();
      const pays = (b.payments || []).filter(p => Number(p.amount) > 0);
      const total = Number(b.total) || 0;
      const adv = pays.reduce((a,p) => a + Number(p.amount), 0);
      const bal = total - adv;
      const tithe = Math.round(total * 0.01 * 100) / 100;      // 1% of full booking amount
      const status = bal <= 0 && total > 0 ? 'Fully Paid' : adv > 0 ? 'Partially Paid' : 'Pending';
      const bs = sh('Bookings'), old = findRow(bs, id);
      const created = old ? bs.getRange(old,17).getValue() : now;
      upsert(bs, id, [id,b.date,b.name,b.phone,b.event,b.time,b.location,b.notes,total,adv,bal,Number(b.sangeetha)||0,Number(b.rizwana)||0,tithe,status,b.confirmation||'Pending',created,now]);
      const ps = sh('Payments'); removeById(ps, id, 1);          // rewrite payments: no duplicates
      pays.forEach((p,i) => ps.appendRow([id+'-P'+(i+1), id, p.date, Number(p.amount)]));
      upsert(sh('Tithe Records'), id, [id, b.date, b.name, total, tithe]);  // one tithe row per booking
      return out({ok:true, id});
    }
    return out({ok:false, error:'Unknown action'});
  } catch (err) { return out({ok:false, error:String(err)}); }
  finally { lock.releaseLock(); }
}
