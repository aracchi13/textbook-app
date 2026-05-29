// ============================================================
// サロン情報管理システム - Code.gs
// ============================================================

// ▼ スプレッドシートIDをここに貼り付けてください ▼
var SPREADSHEET_ID = '1jwJEh2Z9zpj3HwuE2IU-Tq-H5Gx60ec6N4-VaFB8HqU';
// ▲ ここまで ▲

var SH_PROJECT  = '案件';
var SH_SALON    = 'サロン情報';
var SH_TEMPLATE = 'テンプレート';

// ------------------------------------------------------------
// Webアプリ エントリーポイント
// ------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('サロン情報管理システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------------------------------------------
// シート初期化
// ------------------------------------------------------------
function initSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var sheetDefs = [
    { name: SH_PROJECT,  headers: ['ID','年度','月','案件名','作成日時'] },
    { name: SH_SALON,    headers: ['ID','案件ID','会社名','サロン名','キャッチコピー','代表者名','設立年','店舗数','スタッフ数','住所','TEL','本文テキスト','タグ','写真パス1','写真パス2','写真パス3','ロゴパス','Instagramリンク','HPリンク','レイアウト','作成日時','更新日時'] },
    { name: SH_TEMPLATE, headers: ['ID','テンプレート名','会社名','サロン名','キャッチコピー','代表者名','設立年','店舗数','スタッフ数','住所','TEL','本文テキスト','タグ','写真パス1','写真パス2','写真パス3','ロゴパス','Instagramリンク','HPリンク','レイアウト','作成日時'] }
  ];

  sheetDefs.forEach(function(def) {
    var sh = ss.getSheetByName(def.name);
    if (!sh) {
      sh = ss.insertSheet(def.name);
      sh.appendRow(def.headers);
      sh.getRange(1, 1, 1, def.headers.length)
        .setFontWeight('bold')
        .setBackground('#1a2b4c')
        .setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  });

  return { success: true };
}

// ============================================================
// 案件管理
// ============================================================
function getProjects() {
  initSheets();
  var rows = _getRows(SH_PROJECT);
  return rows.map(function(r) {
    return {
      id:        String(r[0]),
      year:      String(r[1]),
      month:     String(r[2]),
      name:      String(r[3]),
      createdAt: _fmtDate(r[4])
    };
  });
}

function saveProject(data) {
  var sh  = _sheet(SH_PROJECT);
  var all = sh.getDataRange().getValues();

  if (data.id) {
    for (var i = 1; i < all.length; i++) {
      if (String(all[i][0]) === String(data.id)) {
        sh.getRange(i + 1, 2, 1, 3).setValues([[data.year, data.month, data.name]]);
        return { success: true, id: data.id };
      }
    }
  }

  var id = 'P' + Date.now();
  sh.appendRow([id, data.year, data.month, data.name, new Date()]);
  return { success: true, id: id };
}

function deleteProject(id) {
  // 配下のサロンを削除
  var ssh  = _sheet(SH_SALON);
  var rows = ssh.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][1]) === String(id)) ssh.deleteRow(i + 1);
  }

  // 案件を削除
  var psh   = _sheet(SH_PROJECT);
  var prows = psh.getDataRange().getValues();
  for (var j = 1; j < prows.length; j++) {
    if (String(prows[j][0]) === String(id)) { psh.deleteRow(j + 1); break; }
  }

  return { success: true };
}

function copyProject(sourceId, targetYear, targetMonth) {
  var projects = getProjects();
  var src = _find(projects, 'id', sourceId);
  if (!src) throw new Error('コピー元案件が見つかりません');

  var newP   = saveProject({ year: targetYear, month: targetMonth, name: src.name });
  var salons = getSalonsByProject(sourceId);

  salons.forEach(function(s) {
    var copy     = _clone(s);
    delete copy.id;
    copy.projectId = newP.id;
    copy.createdAt = '';
    saveSalon(copy);
  });

  return { success: true, newProjectId: newP.id };
}

// ============================================================
// サロン情報管理
// ============================================================
function getSalonsByProject(projectId) {
  var rows = _getRows(SH_SALON);
  return rows
    .filter(function(r) { return String(r[1]) === String(projectId); })
    .map(_rowToSalon);
}

function getSalon(id) {
  var rows = _getRows(SH_SALON);
  var row  = _find(rows, 0, id);
  if (!row) throw new Error('サロンが見つかりません');
  return _rowToSalon(row);
}

function saveSalon(data) {
  var sh  = _sheet(SH_SALON);
  var all = sh.getDataRange().getValues();
  var now = new Date();
  var row = _salonToRow(data, now);

  if (data.id) {
    for (var i = 1; i < all.length; i++) {
      if (String(all[i][0]) === String(data.id)) {
        row[0]  = data.id;
        row[20] = all[i][20];
        sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
        return { success: true, id: data.id };
      }
    }
  }

  var id  = 'S' + Date.now();
  row[0]  = id;
  row[20] = now;
  sh.appendRow(row);
  return { success: true, id: id };
}

function deleteSalon(id) {
  var sh   = _sheet(SH_SALON);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) { sh.deleteRow(i + 1); return { success: true }; }
  }
  throw new Error('サロンが見つかりません');
}

function copySalonToProject(salonId, targetProjectId) {
  var salon       = getSalon(salonId);
  delete salon.id;
  salon.projectId = targetProjectId;
  salon.createdAt = '';
  return saveSalon(salon);
}

// ============================================================
// テンプレート管理
// ============================================================
function getTemplates() {
  var rows = _getRows(SH_TEMPLATE);
  return rows.map(_rowToTemplate);
}

function saveTemplate(data) {
  var sh  = _sheet(SH_TEMPLATE);
  var all = sh.getDataRange().getValues();
  var now = new Date();
  var row = _templateToRow(data, now);

  if (data.id) {
    for (var i = 1; i < all.length; i++) {
      if (String(all[i][0]) === String(data.id)) {
        row[0] = data.id;
        sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
        return { success: true, id: data.id };
      }
    }
  }

  var id = 'T' + Date.now();
  row[0] = id;
  sh.appendRow(row);
  return { success: true, id: id };
}

function deleteTemplate(id) {
  var sh   = _sheet(SH_TEMPLATE);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) { sh.deleteRow(i + 1); return { success: true }; }
  }
  throw new Error('テンプレートが見つかりません');
}

// サロン情報からテンプレートとして保存
function saveAsTemplate(salonId, templateName) {
  var salon = getSalon(salonId);
  return saveTemplate({
    templateName:   templateName,
    companyName:    salon.companyName,
    salonName:      salon.salonName,
    catchphrase:    salon.catchphrase,
    representative: salon.representative,
    established:    salon.established,
    stores:         salon.stores,
    staff:          salon.staff,
    address:        salon.address,
    tel:            salon.tel,
    bodyText:       salon.bodyText,
    tags:           salon.tags,
    photo1:         salon.photo1,
    photo2:         salon.photo2,
    photo3:         salon.photo3,
    logo:           salon.logo,
    instagram:      salon.instagram,
    hp:             salon.hp,
    layout:         salon.layout
  });
}

// ============================================================
// ダッシュボード
// ============================================================
function getDashboard() {
  initSheets();
  var projects  = getProjects();
  var salonRows = _getRows(SH_SALON);
  var required  = [2, 3, 9, 10]; // 会社名,サロン名,住所,TELのインデックス

  return projects.map(function(p) {
    var salons   = salonRows.filter(function(r) { return String(r[1]) === String(p.id); });
    var warnings = salons.filter(function(r) {
      return required.some(function(idx) { return !r[idx]; });
    }).length;

    return { project: p, count: salons.length, warnings: warnings };
  });
}

// ============================================================
// 全サロン一覧（案件横断）
// ============================================================
function getAllSalons() {
  initSheets();
  var projects = getProjects();
  var rows = _getRows(SH_SALON);
  return rows.map(function(r) {
    var salon = _rowToSalon(r);
    var proj = null;
    for (var i = 0; i < projects.length; i++) {
      if (projects[i].id === salon.projectId) { proj = projects[i]; break; }
    }
    salon.projectName = proj ? proj.year + ' ' + proj.month + ' ' + proj.name : '';
    return salon;
  });
}

// ============================================================
// CSV出力（InDesign用）
// ============================================================
function exportCSV(projectId, layout) {
  var salons  = getSalonsByProject(projectId);
  var perRow  = layout === '1/3' ? 3 : 2;
  var baseH   = ['会社名','サロン名','キャッチコピー','代表者名','設立年','店舗数','スタッフ数','住所','TEL','本文テキスト','タグ','写真パス1','写真パス2','写真パス3','ロゴパス','Instagramリンク','HPリンク','レイアウト'];

  var headers = [];
  for (var n = 1; n <= perRow; n++) {
    baseH.forEach(function(h) { headers.push('@' + h + '_' + n); });
  }

  var rows = [headers];

  for (var i = 0; i < salons.length; i += perRow) {
    var group = salons.slice(i, i + perRow);
    var row   = [];
    for (var j = 0; j < perRow; j++) {
      var s = group[j];
      if (s) {
        row = row.concat([s.companyName, s.salonName, s.catchphrase, s.representative, s.established, s.stores, s.staff, s.address, s.tel, s.bodyText, (s.tags||[]).join('|'), s.photo1, s.photo2, s.photo3, s.logo, s.instagram, s.hp, s.layout]);
      } else {
        for (var k = 0; k < baseH.length; k++) row.push('');
      }
    }
    rows.push(row);
  }

  var csv = rows.map(function(r) {
    return r.map(function(c) {
      var s = String(c == null ? '' : c).replace(/"/g, '""');
      return (s.indexOf(',') !== -1 || s.indexOf('\n') !== -1 || s.indexOf('"') !== -1) ? '"' + s + '"' : s;
    }).join(',');
  }).join('\r\n');

  return { success: true, csv: '﻿' + csv };
}

// ============================================================
// ユーティリティ
// ============================================================
function getYears() {
  var y = new Date().getFullYear(), list = [];
  for (var i = y - 2; i <= y + 3; i++) list.push(i + '年度');
  return list;
}

function getMonths() {
  return ['4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月','3月'];
}

function _sheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function _getRows(sheetName) {
  var sh   = _sheet(sheetName);
  var data = sh.getDataRange().getValues();
  return data.slice(1).filter(function(r) { return r[0] !== ''; });
}

function _find(arr, key, val) {
  for (var i = 0; i < arr.length; i++) {
    if (typeof key === 'number' ? String(arr[i][key]) === String(val) : String(arr[i][key]) === String(val)) return arr[i];
  }
  return null;
}

function _clone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function _fmtDate(v) {
  if (!v) return '';
  try { return Utilities.formatDate(new Date(v), 'Asia/Tokyo', 'yyyy/MM/dd'); } catch(e) { return ''; }
}

function _validLayout(v) {
  var s = String(v || '');
  return (s === '1/2' || s === '1/3') ? s : '1/2';
}

function _tagsStr(tags) {
  return Array.isArray(tags) ? tags.join(',') : (tags || '');
}

function _tagsArr(str) {
  return str ? String(str).split(',').filter(function(t){ return t; }) : [];
}

function _rowToSalon(r) {
  return {
    id: String(r[0]), projectId: String(r[1]),
    companyName: String(r[2]||''), salonName: String(r[3]||''),
    catchphrase: String(r[4]||''), representative: String(r[5]||''),
    established: String(r[6]||''), stores: String(r[7]||''),
    staff: String(r[8]||''), address: String(r[9]||''),
    tel: String(r[10]||''), bodyText: String(r[11]||''),
    tags: _tagsArr(r[12]),
    photo1: String(r[13]||''), photo2: String(r[14]||''), photo3: String(r[15]||''),
    logo: String(r[16]||''), instagram: String(r[17]||''),
    hp: String(r[18]||''), layout: _validLayout(r[19]),
    createdAt: _fmtDate(r[20]), updatedAt: _fmtDate(r[21])
  };
}

function _salonToRow(d, now) {
  return [
    d.id||'', d.projectId||'',
    d.companyName||'', d.salonName||'', d.catchphrase||'',
    d.representative||'', d.established||'', d.stores||'',
    d.staff||'', d.address||'', d.tel||'', d.bodyText||'',
    _tagsStr(d.tags),
    d.photo1||'', d.photo2||'', d.photo3||'', d.logo||'',
    d.instagram||'', d.hp||'', d.layout||'1/2',
    d.createdAt ? new Date(d.createdAt) : now, now
  ];
}

function _rowToTemplate(r) {
  return {
    id: String(r[0]), templateName: String(r[1]||''),
    companyName: String(r[2]||''), salonName: String(r[3]||''),
    catchphrase: String(r[4]||''), representative: String(r[5]||''),
    established: String(r[6]||''), stores: String(r[7]||''),
    staff: String(r[8]||''), address: String(r[9]||''),
    tel: String(r[10]||''), bodyText: String(r[11]||''),
    tags: _tagsArr(r[12]),
    photo1: String(r[13]||''), photo2: String(r[14]||''), photo3: String(r[15]||''),
    logo: String(r[16]||''), instagram: String(r[17]||''),
    hp: String(r[18]||''), layout: String(r[19]||'1/2'),
    createdAt: _fmtDate(r[20])
  };
}

function _templateToRow(d, now) {
  return [
    d.id||'', d.templateName||'',
    d.companyName||'', d.salonName||'', d.catchphrase||'',
    d.representative||'', d.established||'', d.stores||'',
    d.staff||'', d.address||'', d.tel||'', d.bodyText||'',
    _tagsStr(d.tags),
    d.photo1||'', d.photo2||'', d.photo3||'', d.logo||'',
    d.instagram||'', d.hp||'', d.layout||'1/2',
    now
  ];
}

// ============================================================
// Google Drive 画像アップロード
// ============================================================
function uploadImage(base64Data, fileName, mimeType, projectId, salonName, slot) {
  try {
    var folder  = _getOrCreateSalonFolder(projectId, salonName);
    var decoded = Utilities.base64Decode(base64Data);
    var blob    = Utilities.newBlob(decoded, mimeType, fileName);

    // 同スロットの既存ファイルを削除
    var existing = folder.getFilesByName(slot + '_' + fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    var file = folder.createFile(blob);
    file.setName(slot + '_' + fileName);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileId = file.getId();
    return {
      success: true,
      fileId:  fileId,
      url:     'https://drive.google.com/uc?export=view&id=' + fileId,
      name:    fileName
    };
  } catch(e) {
    throw new Error('画像アップロードエラー: ' + e.message);
  }
}

function _getOrCreateSalonFolder(projectId, salonName) {
  // ルートフォルダ
  var rootName = 'サロン情報管理_画像';
  var rootIt   = DriveApp.getFoldersByName(rootName);
  var root     = rootIt.hasNext() ? rootIt.next() : DriveApp.createFolder(rootName);

  // 案件フォルダ
  var projects = getProjects();
  var proj     = null;
  for (var i = 0; i < projects.length; i++) {
    if (projects[i].id === String(projectId)) { proj = projects[i]; break; }
  }
  var projName   = proj ? (proj.year + '_' + proj.month + '_' + proj.name) : (projectId || '未分類');
  var projIt     = root.getFoldersByName(projName);
  var projFolder = projIt.hasNext() ? projIt.next() : root.createFolder(projName);

  // サロンフォルダ
  var snName   = salonName || '未設定';
  var snIt     = projFolder.getFoldersByName(snName);
  var snFolder = snIt.hasNext() ? snIt.next() : projFolder.createFolder(snName);

  return snFolder;
}
