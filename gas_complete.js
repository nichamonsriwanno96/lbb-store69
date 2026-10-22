// ═══════════════════════════════════════════════════════════════════════════
// LBB สโตร์ — Google Apps Script (สมบูรณ์)
// วางโค้ดนี้ใน Extensions → Apps Script แล้ว Deploy → New deployment
// Execute as: Me | Who has access: Anyone
// ═══════════════════════════════════════════════════════════════════════════

const S_DOCS  = 'เอกสาร';
const S_ITEMS = 'รายการ';
const S_LOG   = 'Log';
const S_STEEL = 'รายงานผูกมัดเหล็ก';
const S_PROJ  = 'โครงการ';
const S_CAT   = 'ทะเบียนอุปกรณ์';

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    ensureSheets();
    var action = String(e.parameter.action || '');
    if (action === 'getDocs')          return respond(getDocs());
    if (action === 'getProjects')      return respond(getProjects());
    if (action === 'getCatalog')       return respond(getCatalog());
    if (action === 'getSteelRecords')  return respond(getSteelRecords());
    if (action === 'ping')             return respond({ ok: true, msg: 'พร้อมใช้งาน' });
    return respond({ ok: false, error: 'Unknown: ' + action });
  } catch(err) { return respond({ ok: false, error: err.toString() }); }
}

function doPost(e) {
  try {
    ensureSheets();
    var action = String(e.parameter.action || '');
    var body = {};
    try {
      if (e.postData && e.postData.contents) {
        body = JSON.parse(e.postData.contents);
        if (body._action) action = body._action;
      }
    } catch(pe) {}
    if (action === 'addDoc')            return respond(addDoc(body));
    if (action === 'updateDoc')         return respond(updateDoc(body));
    if (action === 'updateStatus')      return respond(updateStatus(body));
    if (action === 'deleteDoc')         return respond(deleteDoc(body));
    if (action === 'saveProject')       return respond(saveProject(body));
    if (action === 'deleteProject')     return respond(deleteProject(body));
    if (action === 'saveCatalogItem')   return respond(saveCatalogItem(body));
    if (action === 'deleteCatalogItem') return respond(deleteCatalogItem(body));
    if (action === 'addSteelRecord')    return respond(addSteelRecord(body));
    if (action === 'uploadPhoto')       return respond(uploadPhoto(body));
    return respond({ ok: false, error: 'Unknown: ' + action });
  } catch(err) { return respond({ ok: false, error: err.toString() }); }
}

// ── getDocs ────────────────────────────────────────────────────────────────
function getDocs() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ds  = ss.getSheetByName(S_DOCS);
  var its = ss.getSheetByName(S_ITEMS);
  var docData = ds.getDataRange().getValues();
  var itmData = its.getDataRange().getValues();
  var imap = {};
  for (var j = 1; j < itmData.length; j++) {
    var r = itmData[j];
    var did = String(r[0]);
    if (!did) continue;
    if (!imap[did]) imap[did] = [];
    imap[did].push({
      _seq: parseInt(r[5])||j,  // คอลัมน์ลำดับที่
      รายการ:String(r[6]||''), ขนาด:String(r[7]||''), รุ่น:String(r[8]||''),
      ยี่ห้อ:String(r[9]||''), หมายเลข:String(r[10]||''),
      จำนวน:String(r[11]||''), หน่วย:String(r[12]||''), หมายเหตุ:String(r[13]||'')
    });
  }
  // เรียงรายการตามลำดับที่
  Object.keys(imap).forEach(function(k){
    imap[k].sort(function(a,b){ return a._seq - b._seq; });
    imap[k].forEach(function(it){ delete it._seq; });
  });
  var docs = [];
  for (var i = docData.length - 1; i >= 1; i--) {
    var row = docData[i];
    if (!row[0]) continue;
    var sig = {};
    try { sig = JSON.parse(String(row[10])||'{}'); } catch(e) {}
    var typeStr = String(row[1]||'');
    var statStr = String(row[2]||'');
    var ftype = typeStr==='นำออก'?'out'
              : typeStr==='นำเข้า'?'in'
              : typeStr==='โยกย้าย'?'mv'
              : (typeStr==='สูญเสีย'||typeStr==='loss')?'loss'
              : 'out';
    // detect จาก สถานะ ด้วย ถ้า typeStr ไม่ชัดเจน
    if(ftype==='out' && (statStr==='สูญเสีย'||statStr==='loss')) ftype='loss';
    var mvParts = String(row[9]||'').split(' → ');
    docs.push({
      id: row[0], ประเภท: ftype, เลขเอกสาร: String(row[3]||''),
      โครงการ: String(row[4]||''), วันที่: String(row[5]||''),
      ผู้ขออนุญาต: String(row[6]||''), ผู้นำของ: String(row[7]||''),
      ทะเบียนรถ: String(row[8]||''),
      โยกจาก: mvParts[0]||'', โยกไป: mvParts[1]||'',
      ลายเซ็น: sig, สถานะ: String(row[2]||''),
      หมายเหตุ: String(row[11]||''),
      รายการสินค้า: imap[String(row[0])] || [],
      synced: true
    });
  }
  var nextNum = 1;
  if (docs.length) {
    var nums = docs.map(function(d){ var m=String(d.เลขเอกสาร).match(/(\d+)$/); return m?parseInt(m[1]):0; });
    nextNum = Math.max.apply(null, nums) + 1;
  }
  return { ok: true, docs: docs, nextNum: nextNum };
}

// ── addDoc ─────────────────────────────────────────────────────────────────
function addDoc(body) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ds  = ss.getSheetByName(S_DOCS);
  var its = ss.getSheetByName(S_ITEMS);
  var log = ss.getSheetByName(S_LOG);
  var id  = body.id || Date.now();
  var type = String(body.ประเภท||'');
  var typeLabel  = type==='out'?'นำออก': type==='in'?'นำเข้า': type==='mv'?'โยกย้าย': type==='loss'?'สูญเสีย':type;
  var statusLabel = body.สถานะ || typeLabel;
  var existing = ds.getDataRange().getValues();
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][0]) === String(id)) return { ok: true, skipped: true };
  }
  var sig = typeof body.ลายเซ็น==='string' ? body.ลายเซ็น : JSON.stringify(body.ลายเซ็น||{});
  var mvInfo = (body.โยกจาก||body.โยกไป) ? (body.โยกจาก||'')+' → '+(body.โยกไป||'') : '';
  ds.appendRow([id, typeLabel, statusLabel, body.เลขเอกสาร||'', body.โครงการ||'',
    body.วันที่||'', body.ผู้ขออนุญาต||'', body.ผู้นำของ||'',
    body.ทะเบียนรถ||'', mvInfo, sig, body.หมายเหตุ||'', body.จำนวนรูป||0, new Date()]);
  var items = [];
  try { items = typeof body.รายการสินค้า==='string'?JSON.parse(body.รายการสินค้า):(body.รายการสินค้า||[]); } catch(e){}
  for (var j = 0; j < items.length; j++) {
    var it = items[j];
    its.appendRow([id, typeLabel, body.เลขเอกสาร||'', body.โครงการ||'', body.วันที่||'',
      j+1, it.รายการ||'', it.ขนาด||'', it.รุ่น||'', it.ยี่ห้อ||'',
      it.หมายเลข||'', it.จำนวน||'', it.หน่วย||'', it.หมายเหตุ||'']);
  }
  log.appendRow([new Date(),'บันทึก',body.เลขเอกสาร||'',body.โครงการ||'',typeLabel,items.length+' รายการ']);
  return { ok: true };
}

// ── updateDoc ────────────────────────────────────────────────────────────────
function updateDoc(body) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ds  = ss.getSheetByName(S_DOCS);
  var its = ss.getSheetByName(S_ITEMS);
  var id  = String(body.id||'');
  var type = String(body.ประเภท||'');
  var typeLabel = type==='out'?'นำออก': type==='in'?'นำเข้า': type==='mv'?'โยกย้าย': type==='loss'?'สูญเสีย':type;
  var statusLabel = body.สถานะ || typeLabel;
  var sig = typeof body.ลายเซ็น==='string' ? body.ลายเซ็น : JSON.stringify(body.ลายเซ็น||{});
  var mvInfo = (body.โยกจาก||body.โยกไป) ? (body.โยกจาก||'')+' → '+(body.โยกไป||'') : '';

  // อัปเดตแถวใน S_DOCS
  var dd = ds.getDataRange().getValues();
  var updated = false;
  for (var i = 1; i < dd.length; i++) {
    if (String(dd[i][0]) === id) {
      ds.getRange(i+1, 1, 1, 14).setValues([[
        id, typeLabel, statusLabel, body.เลขเอกสาร||'', body.โครงการ||'',
        body.วันที่||'', body.ผู้ขออนุญาต||'', body.ผู้นำของ||'',
        body.ทะเบียนรถ||'', mvInfo, sig, body.หมายเหตุ||'', body.จำนวนรูป||0, new Date()
      ]]);
      updated = true;
      break;
    }
  }
  if (!updated) return addDoc(body); // ถ้าไม่มีให้ insert ใหม่

  // ลบรายการเก่าใน S_ITEMS แล้ว insert ใหม่
  var id2 = its.getDataRange().getValues();
  for (var j = id2.length-1; j >= 1; j--) {
    if (String(id2[j][0]) === id) its.deleteRow(j+1);
  }
  var items = [];
  try { items = typeof body.รายการสินค้า==='string'?JSON.parse(body.รายการสินค้า):(body.รายการสินค้า||[]); } catch(e){}
  for (var k = 0; k < items.length; k++) {
    var it = items[k];
    its.appendRow([id, typeLabel, body.เลขเอกสาร||'', body.โครงการ||'', body.วันที่||'',
      k+1, it.รายการ||'', it.ขนาด||'', it.รุ่น||'', it.ยี่ห้อ||'',
      it.หมายเลข||'', it.จำนวน||'', it.หน่วย||'', it.หมายเหตุ||'']);
  }
  return { ok: true };
}

function updateStatus(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ds = ss.getSheetByName(S_DOCS);
  var data = ds.getDataRange().getValues();
  var id = String(body.id||'');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      ds.getRange(i+1, 3).setValue(body.สถานะ||'คืนแล้ว');
      return { ok: true };
    }
  }
  return { ok: false };
}

// ── deleteDoc ──────────────────────────────────────────────────────────────
function deleteDoc(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ds = ss.getSheetByName(S_DOCS);
  var its = ss.getSheetByName(S_ITEMS);
  var id = String(body.id||'');
  var dd = ds.getDataRange().getValues();
  for (var i = dd.length-1; i >= 1; i--) { if (String(dd[i][0])===id){ds.deleteRow(i+1);break;} }
  var id2 = its.getDataRange().getValues();
  for (var j = id2.length-1; j >= 1; j--) { if (String(id2[j][0])===id) its.deleteRow(j+1); }
  return { ok: true };
}

// ── PROJECTS ───────────────────────────────────────────────────────────────
function getProjects() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(S_PROJ);
  if (!s) return { ok: true, projects: [] };
  var data = s.getDataRange().getValues();
  var projects = [];
  for (var i = 1; i < data.length; i++) { if (data[i][0]) projects.push(String(data[i][0])); }
  return { ok: true, projects: projects };
}

function saveProject(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(S_PROJ);
  var name = String(body.name||'').trim();
  if (!name || !s) return { ok: false };
  var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase()===name.toLowerCase()) return { ok: true, skipped: true };
  }
  s.appendRow([name, new Date()]);
  return { ok: true };
}

function deleteProject(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(S_PROJ);
  var name = String(body.name||'').trim();
  if (!s) return { ok: false };
  var data = s.getDataRange().getValues();
  for (var i = data.length-1; i >= 1; i--) {
    if (String(data[i][0])===name) { s.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: false };
}

// ── CATALOG ────────────────────────────────────────────────────────────────
function getCatalog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(S_CAT);
  if (!s) return { ok: true, items: [] };
  var data = s.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      var aliasArr = [];
      try { aliasArr = data[i][4] ? JSON.parse(String(data[i][4])) : []; } catch(e) {}
      items.push({
        code: String(data[i][0]),
        name: String(data[i][1]),
        unit: String(data[i][2]||'อัน'),
        id:   String(data[i][3]||'r'+i),
        alias: aliasArr
      });
    }
  }
  return { ok: true, items: items };
}

function saveCatalogItem(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(S_CAT);
  if (!s) return { ok: false };
  var id = String(body.id||'');
  var code = String(body.code||'');
  var aliasStr = body.alias||'[]';
  if(typeof aliasStr !== 'string') aliasStr = JSON.stringify(aliasStr);
  var data = s.getDataRange().getValues();
  // ค้นหาด้วย id ก่อน ถ้าไม่เจอให้ค้นด้วย code
  for (var i = 1; i < data.length; i++) {
    var rowId = String(data[i][3]||'');
    var rowCode = String(data[i][0]||'');
    if ((id && rowId===id) || (code && rowCode===code)) {
      s.getRange(i+1,1,1,5).setValues([[body.code||'',body.name||'',body.unit||'อัน',rowId||id,aliasStr]]);
      return { ok: true, action:'updated', id:rowId||id };
    }
  }
  // ไม่เจอ — เพิ่มใหม่
  var newId = id || 'cat_'+new Date().getTime();
  s.appendRow([body.code||'', body.name||'', body.unit||'อัน', newId, aliasStr]);
  return { ok: true, action:'added', id:newId };
}

function deleteCatalogItem(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(S_CAT);
  if (!s) return { ok: false };
  var id = String(body.id||'');
  var data = s.getDataRange().getValues();
  for (var i = data.length-1; i >= 1; i--) {
    if (String(data[i][3])===id) { s.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: false };
}

// ── getSteelRecords ────────────────────────────────────────────────────────
function getSteelRecords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName(S_STEEL);
  if (!s) return { ok: true, records: [] };
  var data = s.getDataRange().getValues();
  var records = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var r = data[i];
    if (!r[0]) continue;
    var items = [];
    try { items = JSON.parse(String(r[10]||'[]')); } catch(e) {}
    var concreteItems = [];
    try { concreteItems = JSON.parse(String(r[11]||'[]')); } catch(e) {}
    records.push({
      id:            r[0],
      proj:          String(r[1]||''),
      date:          (function(v){
        if(!v) return '';
        if(v instanceof Date){
          return v.getFullYear()+'-'+('0'+(v.getMonth()+1)).slice(-2)+'-'+('0'+v.getDate()).slice(-2);
        }
        var s=String(v);
        if(/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0,10);
        var d=new Date(s);
        if(!isNaN(d.getTime())) return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);
        return s;
      })(r[2]),
      dia:           String(r[3]||''),
      bundle:        String(r[4]||''),
      priceKg:       parseFloat(r[5]||0),
      wageKg:        parseFloat(r[6]||0),
      wireQty:       parseFloat(r[7]||0),
      wirePrice:     parseFloat(r[8]||0),
      grandW:        String(r[9]||'0'),
      items:         items,
      concreteItems: concreteItems,
      steelCost:     String(r[12]||'0'),
      wageCost:      String(r[13]||'0'),
      wireCost:      String(r[14]||'0'),
      note:          String(r[15]||''),
    });
  }
  return { ok: true, records: records };
}

// ── addSteelRecord ─────────────────────────────────────────────────────────
function addSteelRecord(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName(S_STEEL);
  if (!s) { ensureSheets(); s = ss.getSheetByName(S_STEEL); }
  var id = String(body.id || Date.now());
  var items = typeof body.items === 'string' ? body.items : JSON.stringify(body.items||[]);
  var conc  = typeof body.concreteItems === 'string' ? body.concreteItems : JSON.stringify(body.concreteItems||[]);
  var rowData = [
    id, body.proj||'', body.date||'', body.dia||'', body.bundle||'',
    body.priceKg||0, body.wageKg||0, body.wireQty||0, body.wirePrice||0,
    body.grandW||0, items, conc,
    body.steelCost||0, body.wageCost||0, body.wireCost||0,
    body.note||'', new Date()
  ];
  // ถ้ามี id เดิม ให้ update แทน insert
  var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      s.getRange(i+1, 1, 1, rowData.length).setValues([rowData]);
      return { ok: true, action: 'updated' };
    }
  }
  s.appendRow(rowData);
  return { ok: true, action: 'added' };
}

// ── uploadPhoto ────────────────────────────────────────────────────────────
function uploadPhoto(body) {
  try {
    var base64 = body.data || '';
    var docno  = String(body.docno || 'unknown');
    var idx    = parseInt(body.idx || 0);
    var comma  = base64.indexOf(',');
    if (comma >= 0) base64 = base64.substring(comma + 1);
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64),
      'image/jpeg',
      docno + '_' + (idx + 1) + '.jpg'
    );
    var folders = DriveApp.getFoldersByName('LBB สโตร์ — รูปถ่าย');
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('LBB สโตร์ — รูปถ่าย');
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    return {
      ok: true, fileId: fileId,
      viewUrl:  'https://drive.google.com/file/d/' + fileId + '/view',
      thumbUrl: 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w600'
    };
  } catch(err) { return { ok: false, error: err.toString() }; }
}

// ── ensureSheets ───────────────────────────────────────────────────────────
function ensureSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(S_DOCS)) {
    var s1 = ss.insertSheet(S_DOCS);
    s1.appendRow(['ID','ประเภท','สถานะ','เลขเอกสาร','โครงการ','วันที่',
      'ผู้ขออนุญาต','ผู้นำของ','ทะเบียนรถ','โยกย้าย (จาก→ไป)',
      'ลายเซ็น(JSON)','หมายเหตุ','จำนวนรูป','บันทึกเมื่อ']);
    s1.setFrozenRows(1);
    s1.getRange('1:1').setFontWeight('bold').setBackground('#1B3A6B').setFontColor('#FFF');
    s1.setColumnWidth(1,140); s1.setColumnWidth(2,80); s1.setColumnWidth(3,80);
    s1.setColumnWidth(4,130); s1.setColumnWidth(5,180); s1.setColumnWidth(6,90);
    s1.setColumnWidth(7,140); s1.setColumnWidth(8,140); s1.setColumnWidth(9,100);
  }

  if (!ss.getSheetByName(S_ITEMS)) {
    var s2 = ss.insertSheet(S_ITEMS);
    s2.appendRow(['DocID','ประเภท','เลขเอกสาร','โครงการ','วันที่','ลำดับ',
      'รายการ','ขนาด','รุ่น/สี/ความยาว','ยี่ห้อ','หมายเลข S/N','จำนวน','หน่วย','หมายเหตุ']);
    s2.setFrozenRows(1);
    s2.getRange('1:1').setFontWeight('bold').setBackground('#1B3A6B').setFontColor('#FFF');
  }

  if (!ss.getSheetByName(S_LOG)) {
    var s3 = ss.insertSheet(S_LOG);
    s3.appendRow(['เวลา','การกระทำ','เลขเอกสาร','โครงการ','ประเภท','รายละเอียด']);
    s3.setFrozenRows(1);
    s3.getRange('1:1').setFontWeight('bold').setBackground('#555').setFontColor('#FFF');
  }

  if (!ss.getSheetByName(S_PROJ)) {
    var sp = ss.insertSheet(S_PROJ);
    sp.appendRow(['ชื่อโครงการ','เพิ่มเมื่อ']);
    sp.setFrozenRows(1);
    sp.getRange('1:1').setFontWeight('bold').setBackground('#0F6E56').setFontColor('#FFF');
    sp.setColumnWidth(1,220); sp.setColumnWidth(2,160);
    var defaults = ['Gulf','LNG หนองแฟบ','OBA/DATA',
      'สะพานข้ามแม่น้ำมูล, ศรีสะเกษ','ก่อสร้างทางแยกต่างระดับ ตาปี จ.พัทลุง',
      'ศูนย์วิทยาศาสตร์เชียงราย','ปากปรน','ช.การช่าง (ประชาสงเคราะห์)'];
    defaults.forEach(function(n){ sp.appendRow([n, new Date()]); });
  }

  if (!ss.getSheetByName(S_CAT)) {
    var sc = ss.insertSheet(S_CAT);
    sc.appendRow(['รหัส','ชื่ออุปกรณ์','หน่วย','ID (ห้ามแก้)']);
    sc.setFrozenRows(1);
    sc.getRange('1:1').setFontWeight('bold').setBackground('#B8680A').setFontColor('#FFF');
    sc.setColumnWidth(1,110); sc.setColumnWidth(2,240);
    sc.setColumnWidth(3,80); sc.setColumnWidth(4,160);
    sc.getRange('A2').setValue('★ เพิ่มรายการอุปกรณ์ที่นี่ — จะขึ้นใน dropdown ของแอปทุกเครื่อง');
    sc.getRange('A2').setFontColor('#888').setFontStyle('italic');
  }

  // ── Sheet รายงานผูกมัดเหล็ก ─────────────────────────────────────────────
  if (!ss.getSheetByName(S_STEEL)) {
    var st = ss.insertSheet(S_STEEL);
    st.appendRow(['ID','โครงการ','วันที่','ขนาด Dia.','ชุดมัด',
      'ราคา/กก','ค่าแรง/กก','ลวด(ขด)','ราคาลวด/ขด',
      'น.หนักรวม(kg)','รายการเหล็ก(JSON)','ลูกปูน(JSON)',
      'ค่าเหล็กรวม','ค่าแรงรวม','ค่าลวดรวม','หมายเหตุ','บันทึกเมื่อ']);
    st.setFrozenRows(1);
    st.getRange('1:1').setFontWeight('bold').setBackground('#B8680A').setFontColor('#FFF');
    st.setColumnWidth(1,140); st.setColumnWidth(2,180); st.setColumnWidth(3,90);
    st.setColumnWidth(4,90);  st.setColumnWidth(5,90);  st.setColumnWidth(10,120);
    st.setColumnWidth(11,300); st.setColumnWidth(12,200); st.setColumnWidth(17,140);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('LBB สโตร์')
    .addItem('ทดสอบการเชื่อมต่อ', 'testConnection')
    .addItem('สร้าง Sheets ทั้งหมด', 'ensureSheets')
    .addItem('แก้ไข loss → สูญเสีย', 'fixLossType')
    .addToUi();
}

function testConnection() {
  ensureSheets();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function(s){ return s.getName(); });
  SpreadsheetApp.getUi().alert('✓ เชื่อมต่อสำเร็จ!\n\nSheets: ' + sheets.join(', '));
}

// ── fixLossType — แก้ข้อมูลเก่าที่บันทึก "loss" แทน "สูญเสีย" ─────────────
function fixLossType() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ds = ss.getSheetByName(S_DOCS);
  var its = ss.getSheetByName(S_ITEMS);
  var data = ds.getDataRange().getValues();
  var fixed = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === 'loss') {
      ds.getRange(i+1, 2).setValue('สูญเสีย');
      fixed++;
    }
  }
  // แก้ Sheet รายการ ด้วย
  var idata = its.getDataRange().getValues();
  for (var j = 1; j < idata.length; j++) {
    if (String(idata[j][1]) === 'loss') {
      its.getRange(j+1, 2).setValue('สูญเสีย');
    }
  }
  SpreadsheetApp.getUi().alert('แก้ไขแล้ว ' + fixed + ' รายการ');
}
