/*************************************************************
 * บัญชีเพจครูพร้อมสอน  —  ฝั่งเซิร์ฟเวอร์ (Code.gs)
 * เก็บข้อมูลใน Google Sheets + เก็บรูปสลิป/ยอดเงินใน Google Drive
 *************************************************************/

const SHEET_TX       = 'Transactions';
const SHEET_SETTINGS = 'Settings';
const IMG_FOLDER     = 'บัญชีเพจครูพร้อมสอน_รูปภาพ';

// วางลิงก์โลโก้สาธารณะ (ลงท้าย .png/.jpg เท่านั้น) เพื่อใช้เป็นไอคอน เว้นว่างถ้ายังไม่มี
const LOGO_URL = '';

// วางลิงก์เว็บแอป (ที่ลงท้าย /exec) ของตัวเองตรงนี้ เพื่อให้ปุ่ม "เปิดแอป" ทำงานชัวร์
// เว้นว่างได้ ระบบจะดึงลิงก์อัตโนมัติให้
const APP_URL = '';

// รหัสลับของ API — ต้องตรงกับ TOKEN ในไฟล์ index.html ของ PWA (เปลี่ยนเป็นของตัวเองได้ แต่ต้องตรงกันทั้งสองที่)
const API_TOKEN = 'KPS-7Qm2xV9aR4tLpZ8c';

/** API: รับคำขอจาก PWA แล้วตอบกลับเป็น JSON */
function doGet(e)  { return api_(e); }
function doPost(e) { return api_(e); }

function api_(e) {
  const out = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
  try {
    setup_();
    const p = (e && e.parameter) ? e.parameter : {};
    let body = {};
    if (e && e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (_) {}
    }
    const action = body.action || p.action || '';
    const token  = body.token  || p.token  || '';

    if (action === 'ping') return out.setContent(JSON.stringify({ ok: true, data: 'pong' }));
    if (token !== API_TOKEN) return out.setContent(JSON.stringify({ ok: false, error: 'token ไม่ถูกต้อง' }));

    let data;
    switch (action) {
      case 'getConfig':       data = getConfig(); break;
      case 'getTransactions': data = JSON.parse(getTransactions()); break;
      case 'getImageData':    data = getImageData(body.fileId || p.fileId); break;
      case 'add':             data = addTransaction(body.payload || {}); break;
      case 'update':          data = updateTransaction(body.payload || {}); break;
      case 'delete':          data = deleteTransaction(body.id || p.id); break;
      case 'addSetting':      data = addSetting(body.type, body.value); break;
      case 'deleteSetting':   data = deleteSetting(body.type, body.value); break;
      case 'import':          data = importRows(body.rows || []); break;
      case 'kvGet':           data = kvGet(body.key || p.key); break;
      case 'kvSet':           data = kvSet(body.key, body.value); break;
      case 'exportUrl':       data = getExportUrl(); break;
      default: return out.setContent(JSON.stringify({ ok: false, error: 'ไม่รู้จัก action: ' + action }));
    }
    return out.setContent(JSON.stringify({ ok: true, data: data }));
  } catch (err) {
    return out.setContent(JSON.stringify({ ok: false, error: String(err && err.message || err) }));
  }
}

/** สร้างชีต Transactions / Settings พร้อมค่าตั้งต้น */
function setup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let tx = ss.getSheetByName(SHEET_TX);
  if (!tx) {
    tx = ss.insertSheet(SHEET_TX);
    tx.appendRow(['ID','วันที่','ประเภท','หมวดหมู่','จำนวนเงิน',
                  'บัญชี','คนทำ','หมายเหตุ','รูปภาพ(FileId)','บันทึกเมื่อ']);
    tx.setFrozenRows(1);
  }

  let st = ss.getSheetByName(SHEET_SETTINGS);
  if (!st) {
    st = ss.insertSheet(SHEET_SETTINGS);
    st.appendRow(['ชนิด','รายการ']);
    const def = [
      ['catIncome','ขายแผนการสอน'],
      ['catIncome','ขายระบบโรงเรียน'],
      ['catIncome','ขายกลุ่ม VIP'],
      ['catIncome','อื่นๆ'],
      ['catExpense','ค่า Ads โฆษณา'],
      ['catExpense','ค่า Meta Verify'],
      ['catExpense','ค่าจ้างผู้ช่วย'],
      ['catExpense','ค่าซอฟต์แวร์/เครื่องมือ'],
      ['catExpense','อุปกรณ์/เฟอร์นิเจอร์'],
      ['catExpense','ค่าอินเทอร์เน็ต'],
      ['catExpense','อื่นๆ'],
      ['account','ธ.กรุงเทพ'],
      ['account','K-Shop'],
      ['account','SCB'],
      ['person','ยุย'],
      ['person','น้อง']
    ];
    st.getRange(2,1,def.length,2).setValues(def);
    st.setFrozenRows(1);
  }

  let kv = ss.getSheetByName('KV');
  if (!kv) {
    kv = ss.insertSheet('KV');
    kv.appendRow(['key','value']);
    kv.setFrozenRows(1);
  }
  return { ss, tx, st, kv };
}

/** อ่านค่าตั้งค่า (หมวดหมู่ / บัญชี / คน) */
function getConfig() {
  const { st } = setup_();
  const rows = st.getDataRange().getValues().slice(1);
  const cfg = { catIncome: [], catExpense: [], account: [], person: [] };
  rows.forEach(r => {
    const type = String(r[0]).trim();
    const val  = String(r[1]).trim();
    if (cfg[type] && val) cfg[type].push(val);
  });
  return cfg;
}

/** ดึงรายการทั้งหมด (เรียงล่าสุดก่อน) */
function getTransactions() {
  const { tx } = setup_();
  const values = tx.getDataRange().getValues().slice(1);
  const list = values.map(r => ({
    id:       r[0],
    date:     formatDate_(r[1]),
    type:     r[2],
    category: r[3],
    amount:   Number(r[4]) || 0,
    account:  r[5],
    person:   r[6],
    note:     r[7],
    fileId:   r[8] || '',
    created:  r[9]
  })).filter(o => o.id);
  list.sort((a,b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  return JSON.stringify(list);
}

/** เพิ่มรายการใหม่ */
function addTransaction(p) {
  const { tx } = setup_();
  const id = 'TX' + Date.now() + Math.floor(Math.random()*1000);
  let fileId = '';
  if (p.image && p.image.base64) fileId = saveImage_(p.image, p);
  tx.appendRow([
    id,
    p.date,
    p.type,
    p.category,
    Number(p.amount) || 0,
    p.account || '',
    p.person  || '',
    p.note    || '',
    fileId,
    new Date()
  ]);
  return { ok: true, id: id, fileId: fileId };
}

/** แก้ไขรายการเดิม */
function updateTransaction(p) {
  const { tx } = setup_();
  const data = tx.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === p.id) {
      let fileId = data[i][8] || '';
      if (p.image && p.image.base64) fileId = saveImage_(p.image, p);
      if (p.removeImage) fileId = '';
      tx.getRange(i+1, 2, 1, 8).setValues([[
        p.date, p.type, p.category, Number(p.amount) || 0,
        p.account || '', p.person || '', p.note || '', fileId
      ]]);
      return { ok: true, fileId: fileId };
    }
  }
  return { ok: false, error: 'ไม่พบรายการ' };
}

/** ลบรายการ */
function deleteTransaction(id) {
  const { tx } = setup_();
  const data = tx.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      tx.deleteRow(i+1);
      return { ok: true };
    }
  }
  return { ok: false };
}

/** เพิ่ม/ลบ ค่าตั้งค่า (หมวดหมู่ บัญชี คน) */
function addSetting(type, value) {
  const { st } = setup_();
  st.appendRow([type, String(value).trim()]);
  return getConfig();
}
function deleteSetting(type, value) {
  const { st } = setup_();
  const data = st.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === type && String(data[i][1]).trim() === String(value).trim()) {
      st.deleteRow(i+1);
      break;
    }
  }
  return getConfig();
}

/** บันทึกรูปลง Drive แล้วคืน fileId (ไฟล์เป็นส่วนตัว ไม่เปิดสาธารณะ) */
function saveImage_(image, p) {
  const folder = getFolder_();
  const bytes  = Utilities.base64Decode(image.base64);
  const name   = (p.type || 'img') + '_' + (p.date || '') + '_' +
                 Date.now() + '.' + (image.ext || 'jpg');
  const blob   = Utilities.newBlob(bytes, image.mimeType || 'image/jpeg', name);
  const file   = folder.createFile(blob);
  return file.getId();
}

/** ดึงรูปแบบ data-URL (เรียกตอนกดดูเท่านั้น เพื่อความเป็นส่วนตัว) */
function getImageData(fileId) {
  if (!fileId) return '';
  const blob = DriveApp.getFileById(fileId).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' +
         Utilities.base64Encode(blob.getBytes());
}

function getFolder_() {
  const it = DriveApp.getFoldersByName(IMG_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(IMG_FOLDER);
}

/** เก็บค่า key-value (ใช้เก็บค่าลดหย่อนรายคน ฯลฯ) */
function kvGet(key) {
  const { kv } = setup_();
  const d = kv.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(key)) return String(d[i][1] || '');
  return '';
}
function kvSet(key, value) {
  const { kv } = setup_();
  const d = kv.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(key)) { kv.getRange(i + 1, 2).setValue(value); return true; }
  }
  kv.appendRow([key, value]);
  return true;
}

/** ลิงก์ส่งออกทั้งไฟล์เป็น Excel */
function getExportUrl() {
  return 'https://docs.google.com/spreadsheets/d/' +
         SpreadsheetApp.getActiveSpreadsheet().getId() + '/export?format=xlsx';
}

function formatDate_(d) {
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(d).trim();
}


/*************************************************************
 * เมนูบน Google Sheet + นำเข้า/ส่งออก/สรุป
 *************************************************************/

/** สร้างเมนู "ครูพร้อมสอน" ตอนเปิดชีต */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎓 ครูพร้อมสอน')
    .addItem('📱 เปิดแอปบัญชี', 'openApp')
    .addItem('📊 สรุปด่วนปีนี้', 'quickSummary')
    .addSeparator()
    .addItem('📥 นำเข้าบัญชี (จาก Excel)', 'importMenu')
    .addItem('📤 ส่งออกเป็น Excel', 'exportExcel')
    .addSeparator()
    .addItem('ℹ️ วิธีใช้งาน', 'showHelp')
    .addToUi();
}

/** ปุ่มเปิดแอป (ใช้ APP_URL ถ้ามี ไม่งั้นดึงอัตโนมัติ) */
function openApp() {
  let url = APP_URL;
  if (!url) {
    try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; }
  }
  const body = url
    ? '<a href="' + url + '" target="_blank" ' +
      'style="display:inline-block;background:#FF77B0;color:#1E1B16;text-decoration:none;' +
      'font-weight:700;padding:12px 22px;border:3px solid #1E1B16;border-radius:14px;' +
      'box-shadow:4px 4px 0 #1E1B16">📚 เปิดแอปบัญชีเพจ</a>' +
      '<p style="font-size:12px;color:#666;margin-top:14px">แตะปุ่มเพื่อเปิดแอปในแท็บใหม่</p>'
    : '<p style="color:#d6256b">ยังไม่พบลิงก์เว็บแอป กรุณา Deploy เป็น Web app ก่อน ' +
      'แล้วนำลิงก์ /exec มาวางในตัวแปร APP_URL</p>';
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;text-align:center;padding:18px">' + body + '</div>'
  ).setWidth(340).setHeight(170);
  SpreadsheetApp.getUi().showModalDialog(html, 'ครูพร้อมสอน');
}

/** สรุปด่วนปีปัจจุบัน */
function quickSummary() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TX);
  if (!sh || sh.getLastRow() < 2) { SpreadsheetApp.getUi().alert('ยังไม่มีข้อมูล'); return; }
  const data = sh.getDataRange().getValues().slice(1);
  const y = String(new Date().getFullYear());
  let inc = 0, exp = 0;
  data.forEach(r => {
    if (String(r[1]).slice(0, 4) === y) {
      if (r[2] === 'รับ')  inc += Number(r[4]) || 0;
      if (r[2] === 'จ่าย') exp += Number(r[4]) || 0;
    }
  });
  const f = n => n.toLocaleString('th-TH', { maximumFractionDigits: 2 });
  SpreadsheetApp.getUi().alert(
    'สรุปปี ' + (+y + 543) +
    '\n\nรายรับ: ' + f(inc) + ' บาท' +
    '\nรายจ่าย: ' + f(exp) + ' บาท' +
    '\nกำไร: ' + f(inc - exp) + ' บาท'
  );
}

/** วิธีใช้งาน */
function showHelp() {
  SpreadsheetApp.getUi().alert(
    'ครูพร้อมสอน — บัญชีเพจ\n\n' +
    '• กด "เปิดแอปบัญชี" เพื่อเข้าหน้าแอป\n' +
    '• ข้อมูลทั้งหมดเก็บในแท็บ Transactions\n' +
    '• หมวดหมู่/บัญชี/คน แก้ได้ในแท็บ Settings\n' +
    '• นำเข้าข้อมูลเก่า: ก๊อปจากไฟล์ Excel มาวางในแท็บ Transactions แล้วกด "ตรวจ + เรียงข้อมูล"'
  );
}

/** หน้าต่างนำเข้าบัญชี: เลือกไฟล์ Excel แล้วนำเข้าอัตโนมัติ */
function importMenu() {
  const html = HtmlService.createHtmlOutput(`
<div style="font-family:sans-serif;padding:18px;font-size:14px;line-height:1.6">
  <b style="font-size:15px">เลือกไฟล์ Excel เพื่อนำเข้าอัตโนมัติ</b>
  <p style="color:#666;margin:8px 0">รองรับไฟล์รูปแบบของแอป (ไฟล์ "นำเข้า_บัญชีเพจ" หรือไฟล์ที่ส่งออกจากแอป) คอลัมน์ A–J</p>
  <div style="text-align:center;margin:18px 0">
    <input type="file" id="file" accept=".xlsx,.xls,.csv" style="display:none" onchange="handle()">
    <button onclick="document.getElementById('file').click()"
      style="background:#FFDD3C;border:3px solid #1E1B16;border-radius:12px;padding:12px 20px;font-weight:700;cursor:pointer;box-shadow:3px 3px 0 #1E1B16">📁 เลือกไฟล์ Excel</button>
    <div id="fname" style="margin-top:10px;color:#333"></div>
    <div id="msg" style="margin-top:12px;font-weight:700"></div>
  </div>
  <details style="color:#777;font-size:12px"><summary>หรือก๊อปวางเอง (วิธีเดิม)</summary>
  วางข้อมูล A2–J ในแท็บ Transactions เอง แล้วกดเมนู "ตรวจ + เรียงข้อมูล" จากปุ่มนี้
  <button onclick="google.script.run.withSuccessHandler(function(r){document.getElementById('msg').innerText=r;}).cleanupTransactions()"
    style="margin-top:6px;background:#B4E84A;border:2px solid #1E1B16;border-radius:9px;padding:6px 12px;font-weight:700;cursor:pointer">ตรวจ + เรียงข้อมูล</button></details>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
function pad(n){return String(n).padStart(2,'0');}
function d2s(v){
  if(v instanceof Date) return v.getFullYear()+'-'+pad(v.getMonth()+1)+'-'+pad(v.getDate());
  v=String(v).trim();
  if(/^\\d{4}-\\d{2}-\\d{2}/.test(v)) return v.slice(0,10);
  return v;
}
function handle(){
  var f=document.getElementById('file').files[0]; if(!f) return;
  document.getElementById('fname').innerText='ไฟล์: '+f.name;
  var msg=document.getElementById('msg'); msg.style.color='#0a7'; msg.innerText='กำลังอ่านไฟล์...';
  var rd=new FileReader();
  rd.onload=function(e){
    try{
      var wb=XLSX.read(new Uint8Array(e.target.result),{type:'array',cellDates:true});
      var ws=wb.Sheets[wb.SheetNames[0]];
      var aoa=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      var rows=[], start=0;
      if(aoa.length && (String(aoa[0][2]).indexOf('ประเภท')>=0 || String(aoa[0][1]).indexOf('วันที่')>=0)) start=1;
      for(var i=start;i<aoa.length;i++){
        var r=aoa[i]; if(!r) continue;
        var amt=parseFloat(String(r[4]).replace(/,/g,''))||0;
        if(!r[1]||!r[2]||!amt) continue;
        rows.push({id:String(r[0]||''),date:d2s(r[1]),type:String(r[2]).trim(),category:String(r[3]||''),
          amount:amt,account:String(r[5]||''),person:String(r[6]||''),note:String(r[7]||''),fileId:String(r[8]||'')});
      }
      if(!rows.length){ msg.style.color='#d6256b'; msg.innerText='ไม่พบข้อมูลที่นำเข้าได้ ตรวจรูปแบบไฟล์อีกครั้ง'; return; }
      msg.innerText='กำลังนำเข้า '+rows.length+' รายการ...';
      google.script.run
        .withSuccessHandler(function(res){ msg.style.color='#0a7'; msg.innerText='✓ '+res+' — รีเฟรชแอปเพื่อดูข้อมูล'; })
        .withFailureHandler(function(err){ msg.style.color='#d6256b'; msg.innerText='ผิดพลาด: '+err.message; })
        .importRows(rows);
    }catch(err){ msg.style.color='#d6256b'; msg.innerText='อ่านไฟล์ไม่ได้: '+err.message; }
  };
  rd.readAsArrayBuffer(f);
}
</script>
`).setWidth(440).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, '📥 นำเข้าบัญชี');
}

/** เขียนข้อมูลที่นำเข้าลงแท็บ Transactions (ข้ามรายการที่ ID ซ้ำ แล้วเรียงตามวันที่) */
function importRows(rows) {
  const { tx } = setup_();
  const existing = {};
  const data = tx.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (data[i][0]) existing[data[i][0]] = true;

  const out = [];
  let skipped = 0;
  (rows || []).forEach((r, i) => {
    const amt = Number(r.amount) || 0;
    if (!r.date || !r.type || !amt) return;
    let id = r.id || ('IMP' + Date.now() + '_' + i);
    if (existing[id]) { skipped++; return; }
    existing[id] = true;
    out.push([id, r.date, r.type, r.category || '', amt,
              r.account || '', r.person || '', r.note || '', r.fileId || '', new Date()]);
  });

  if (out.length) {
    tx.getRange(tx.getLastRow() + 1, 1, out.length, 10).setValues(out);
    const all = tx.getRange(2, 1, tx.getLastRow() - 1, 10).getValues();
    all.sort((a, b) => String(a[1]) < String(b[1]) ? -1 : 1);
    tx.getRange(2, 1, all.length, 10).setValues(all);
  }
  return 'นำเข้า ' + out.length + ' รายการ' + (skipped ? ' (ข้ามรายการซ้ำ ' + skipped + ')' : '');
}

/** ตรวจ + เติม ID + แปลงวันที่ + เรียงข้อมูลในแท็บ Transactions */
function cleanupTransactions() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TX);
  if (!sh || sh.getLastRow() < 2) return 'ยังไม่มีข้อมูลในแท็บ Transactions';
  let v = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  v = v.filter(r => r[1] && r[2] && (Number(r[4]) || 0) !== 0);
  v.forEach((r, i) => {
    if (!r[0]) r[0] = 'TX' + Date.now() + i;
    if (r[1] instanceof Date) {
      r[1] = Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  });
  v.sort((a, b) => String(a[1]) < String(b[1]) ? -1 : 1);
  sh.getRange(2, 1, sh.getMaxRows() - 1, 10).clearContent();
  if (v.length) sh.getRange(2, 1, v.length, 10).setValues(v);
  return 'เรียบร้อย! ข้อมูลพร้อมใช้ ' + v.length + ' รายการ';
}

/** ส่งออกทั้งไฟล์เป็น Excel */
function exportExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;text-align:center;padding:18px">' +
    '<p style="font-size:14px">ดาวน์โหลดข้อมูลทั้งหมดเป็นไฟล์ Excel</p>' +
    '<a href="' + url + '" target="_blank" ' +
    'style="display:inline-block;background:#62C7FF;color:#1E1B16;text-decoration:none;' +
    'font-weight:700;padding:12px 22px;border:3px solid #1E1B16;border-radius:14px;' +
    'box-shadow:4px 4px 0 #1E1B16">📤 ดาวน์โหลด .xlsx</a>' +
    '<p style="font-size:12px;color:#666;margin-top:14px">ได้ไฟล์ครบทุกแท็บ เปิดด้วย Excel ได้เลย</p></div>'
  ).setWidth(340).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, '📤 ส่งออก Excel');
}
