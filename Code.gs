// วางใน Google Sheet > ส่วนขยาย > Apps Script แล้ว Deploy เป็น Web app
// (Execute as: Me, Who has access: Anyone)
const PIN = '1234'; // เปลี่ยนเป็นรหัสของคุณ

function sh(n, h) {
  const s = SpreadsheetApp.getActive();
  if (!s) throw new Error('กรุณาเปิด Apps Script จาก Google Sheet โดยตรง (ส่วนขยาย > Apps Script)');
  let x = s.getSheetByName(n);
  if (!x) {
    x = s.insertSheet(n);
    x.appendRow(h);
    if (n === 'Entries') x.getRange('B:B').setNumberFormat('@');
  } else if (x.getLastRow() === 0) {
    x.appendRow(h);
    if (n === 'Entries') x.getRange('B:B').setNumberFormat('@');
  }
  return x;
}

const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const E = () => sh('Entries', ['id', 'date', 'type', 'item', 'amount']);
const I = () => sh('Items', ['item']);

function toDateStr(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  return s;
}

function handle(p) {
  if (!p) return { error: 'no_data' };
  if (String(p.pin) !== String(PIN)) return { error: 'pin' };

  // Read-only action: ไม่ต้องใช้ Lock เพื่อความรวดเร็วและป้องกัน Lock Timeout
  if (p.action === 'list') {
    const rawEntries = E().getDataRange().getValues().slice(1);
    const entries = rawEntries
      .filter(r => r[0] && String(r[0]).trim())
      .map(r => ({
        id: String(r[0]),
        date: toDateStr(r[1]),
        type: String(r[2]),
        item: String(r[3] || ''),
        amount: Number(r[4]) || 0
      }));

    const rawItems = I().getDataRange().getValues().slice(1);
    const items = rawItems
      .map(r => String(r[0] || '').trim())
      .filter(Boolean);

    return { entries, items };
  }

  // Write actions: ใช้ Script Lock ป้องกัน Concurrency Conflict
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (p.action === 'add') {
      const id = Utilities.getUuid();
      const date = toDateStr(p.date) || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
      const type = p.type === 'expense' ? 'expense' : 'income';
      const item = String(p.item || (type === 'income' ? 'ขายป๊อบคอร์น' : 'ทั่วไป')).trim();
      const amount = Number(p.amount) || 0;
      E().appendRow([id, date, type, item, amount]);
      return { ok: 1, id };
    }

    if (p.action === 'del') {
      const s = E(), v = s.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) {
        if (String(v[i][0]) === String(p.id)) {
          s.deleteRow(i + 1);
          break;
        }
      }
      return { ok: 1 };
    }

    if (p.action === 'addItem') {
      const item = String(p.item || '').trim();
      if (!item) return { error: 'empty_item' };
      const s = I(), v = s.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) {
        if (String(v[i][0]).trim().toLowerCase() === item.toLowerCase()) {
          return { ok: 1 }; // มีอยู่แล้ว ไม่เพิ่มซ้ำ
        }
      }
      s.appendRow([item]);
      return { ok: 1 };
    }

    if (p.action === 'delItem') {
      const target = String(p.item || '').trim();
      const s = I(), v = s.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) {
        if (String(v[i][0]).trim() === target) {
          s.deleteRow(i + 1);
          break;
        }
      }
      return { ok: 1 };
    }

    return { error: 'action' };
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return out(handle(e && e.parameter ? e.parameter : {}));
}

function doPost(e) {
  let p = {};
  try {
    p = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
  } catch (err) {
    return out({ error: 'invalid_json' });
  }
  return out(handle(p));
}
