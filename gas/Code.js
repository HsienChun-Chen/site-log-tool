// ============================================================
// 工地日誌工具 - Google Apps Script
// ============================================================

// ── 設定（只需填這裡）──────────────────────────────────────
const SPREADSHEET_ID   = '1a4N2m72WqDKDbOeFVKmW2YfmNVA-vNKApQjzGwQGDbE';
const UPLOAD_FOLDER_ID = '11zjIQX0hI79OnYdLpoRZaH-0m5QhlK6s';
const PHOTO_FOLDER_ID  = '11zjIQX0hI79OnYdLpoRZaH-0m5QhlK6s'; // 目前與上傳資料夾相同，之後可獨立
// ────────────────────────────────────────────────────────────

const SHEET_CASES = '案件清單';
const SHEET_ITEMS = '工項清單';
const SHEET_DAILY = '每日記錄';

// 民國年
function getRocYear() {
  return new Date().getFullYear() - 1911;
}


// ============================================================
// STEP 1：初始化 InfoData 試算表（第一次執行）
// ============================================================
function initSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function ensureSheet(name, headers) {
    let ws = ss.getSheetByName(name);
    if (!ws) {
      ws = ss.insertSheet(name);
      ws.appendRow(headers);
      ws.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f0fe');
    }
    return ws;
  }

  ensureSheet(SHEET_CASES, ['案號', '案件名稱', '負責人', '狀態', '金額', '備註', '最後更新']);
  ensureSheet(SHEET_ITEMS, ['案號', '案件名稱', '類別', '項目', '單位', '總數量']);
  ensureSheet(SHEET_DAILY, ['時間戳', '日期', '案號', '案件名稱', '類別', '項目', '今日數量', '工種人數', '備註', '照片資料夾']);

  SpreadsheetApp.getUi().alert('✅ 初始化完成！三個工作表已建立。');
}


// ============================================================
// STEP 2：同步案件清單（從 @案件列表.xlsx 的 115年 工作表）
// ============================================================
function syncCaseList() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dest = ss.getSheetByName(SHEET_CASES);

  // 找 @案件列表.xlsx
  const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  const files  = folder.getFilesByName('@案件列表.xlsx');
  if (!files.hasNext()) {
    SpreadsheetApp.getUi().alert('❌ 找不到 @案件列表.xlsx，請上傳至指定資料夾');
    return;
  }

  const rocYear   = getRocYear();
  const sheetName = `${rocYear}年`;

  // 轉換為 Google Sheets 格式（暫存）
  const tempId = _convertToSheets(files.next().getId(), 'TEMP_案件列表');
  const ws     = SpreadsheetApp.openById(tempId).getSheetByName(sheetName);

  if (!ws) {
    DriveApp.getFileById(tempId).setTrashed(true);
    SpreadsheetApp.getUi().alert(`❌ 找不到 ${sheetName} 工作表`);
    return;
  }

  const data    = ws.getDataRange().getValues();
  const rows    = [];
  let   section = '';

  for (const row of data) {
    const b = row[1], c = row[2], d = row[3], e = row[4], f = row[5], g = row[6];

    // 辨識區段標題
    if (['在案列表', '結案中列表', '提案、報價中', '已結案'].includes(b)) {
      section = b; continue;
    }
    // 跳過欄位標題、合計列
    if (b === '編號' || b === '合計' || !e) continue;
    // 只要在案 + 結案中
    if (!['在案列表', '結案中列表'].includes(section)) continue;

    rows.push([
      c || '',                                          // 案號
      e,                                               // 案件名稱
      d || '',                                         // 負責人
      section === '在案列表' ? '在建' : '結案中',        // 狀態
      f || '',                                         // 金額
      g || '',                                         // 備註
      new Date()                                       // 最後更新
    ]);
  }

  // 寫入（保留標題列，清除舊資料）
  if (dest.getLastRow() > 1) {
    dest.getRange(2, 1, dest.getLastRow() - 1, 7).clearContent();
  }
  if (rows.length > 0) {
    dest.getRange(2, 1, rows.length, 7).setValues(rows);
  }

  DriveApp.getFileById(tempId).setTrashed(true);
  SpreadsheetApp.getUi().alert(`✅ 案件清單同步完成！共 ${rows.length} 筆`);
}


// ============================================================
// STEP 3：同步工項清單（分批執行，每次 5 份，可重複跑直到完成）
// ============================================================
const BATCH_SIZE = 5;

function syncWorkItems() {
  const props  = PropertiesService.getScriptProperties();
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dest   = ss.getSheetByName(SHEET_ITEMS);
  const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);

  // 取得所有毛利管制表檔案清單
  const allFiles = [];
  const iter = folder.getFiles();
  while (iter.hasNext()) {
    const f = iter.next();
    if (f.getName().includes('毛利管制表') && f.getName().endsWith('.xlsx')) {
      allFiles.push({ id: f.getId(), name: f.getName() });
    }
  }

  // 已完成的檔案（存在 ScriptProperties）
  const doneJson = props.getProperty('syncDone') || '[]';
  const done     = new Set(JSON.parse(doneJson));

  // 找出這批要處理的檔案
  const pending = allFiles.filter(f => !done.has(f.id));

  if (pending.length === 0) {
    props.deleteProperty('syncDone'); // 重置，下次可重新全跑
    SpreadsheetApp.getUi().alert('✅ 全部工項同步完成！');
    return;
  }

  const batch    = pending.slice(0, BATCH_SIZE);
  const newRows  = [];

  batch.forEach(f => {
    const file = DriveApp.getFileById(f.id);
    const rows = _parseWorkItems(file);
    newRows.push(...rows);
    done.add(f.id);
  });

  // 追加寫入（不清除，累積）
  if (newRows.length > 0) {
    const startRow = dest.getLastRow() + 1;
    dest.getRange(startRow, 1, newRows.length, 6).setValues(newRows);
  }

  // 儲存進度
  props.setProperty('syncDone', JSON.stringify([...done]));

  const remaining = pending.length - batch.length;
  if (remaining > 0) {
    SpreadsheetApp.getUi().alert(
      `✅ 這批完成（${batch.length} 份），還剩 ${remaining} 份。\n請再執行一次 syncWorkItems 繼續。`
    );
  } else {
    props.deleteProperty('syncDone');
    SpreadsheetApp.getUi().alert(`✅ 全部工項同步完成！共寫入 ${dest.getLastRow() - 1} 筆`);
  }
}

// 若想重頭重跑（清除進度），執行這個
function resetSyncProgress() {
  PropertiesService.getScriptProperties().deleteProperty('syncDone');
  const dest = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_ITEMS);
  if (dest.getLastRow() > 1) dest.getRange(2, 1, dest.getLastRow() - 1, 6).clearContent();
  SpreadsheetApp.getUi().alert('已重置，可重新執行 syncWorkItems');
}

function _parseWorkItems(file) {
  const fileName     = file.getName();
  const caseFolderName = fileName.replace(/_毛利管制表\.xlsx$/, '');
  const caseId       = caseFolderName.split('_')[0];

  const tempId = _convertToSheets(file.getId(), 'TEMP_' + caseId);
  const ws     = SpreadsheetApp.openById(tempId).getSheetByName('預估');

  if (!ws) {
    DriveApp.getFileById(tempId).setTrashed(true);
    return [];
  }

  const data     = ws.getDataRange().getValues();
  const items    = [];
  let   category = '';

  // 資料從第 7 列開始（index 6），第 6 列（index 5）是標題
  for (let i = 6; i < data.length; i++) {
    const row  = data[i];
    const no   = row[0]; // NO.
    const cat  = row[1]; // 類別
    const item = row[2]; // 項目
    const unit = row[3]; // 單位
    const qty  = row[5]; // 售價數量（Col F）

    // 遇到頁尾小計停止
    if (typeof no === 'string' && /小計|總計|毛利率/.test(no)) break;

    // 類別 fill-forward
    if (cat) category = cat;

    // 跳過沒有售價數量的列（成本子項）
    if (typeof qty !== 'number' || qty <= 0) continue;
    if (!item) continue;

    items.push([caseId, caseFolderName, category, item, unit || '', qty]);
  }

  DriveApp.getFileById(tempId).setTrashed(true);
  return items;
}


// ============================================================
// Web App：GET — 提供資料給 HTML 表單
// ============================================================
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getCases')   return _json(_getCases());
  if (action === 'getItems')   return _json(_getItems(e.parameter.caseId));
  if (action === 'getProgress') return _json(_getProgress(e.parameter.caseId));

  // 沒有 action 就回傳表單頁面
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('工地日誌')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _getCases() {
  const data = SpreadsheetApp.openById(SPREADSHEET_ID)
    .getSheetByName(SHEET_CASES).getDataRange().getValues().slice(1);
  return data
    .filter(r => r[3] === '在建' && r[0])
    .map(r => ({ caseId: r[0], caseName: r[1], manager: r[2] }));
}

function _getItems(caseId) {
  const data = SpreadsheetApp.openById(SPREADSHEET_ID)
    .getSheetByName(SHEET_ITEMS).getDataRange().getValues().slice(1);
  return data
    .filter(r => r[0] === caseId)
    .map(r => ({ category: r[2], item: r[3], unit: r[4], totalQty: r[5] }));
}

function _getProgress(caseId) {
  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const itemData  = ss.getSheetByName(SHEET_ITEMS).getDataRange().getValues().slice(1)
                      .filter(r => r[0] === caseId);
  const dailyData = ss.getSheetByName(SHEET_DAILY).getDataRange().getValues().slice(1)
                      .filter(r => r[2] === caseId);

  const map = {};
  itemData.forEach(r => {
    map[r[3]] = { category: r[2], item: r[3], unit: r[4], totalQty: r[5], doneQty: 0 };
  });
  dailyData.forEach(r => {
    const key = r[5];
    if (map[key]) map[key].doneQty += parseFloat(r[6]) || 0;
  });

  return Object.values(map).map(p => ({
    ...p,
    pct: p.totalQty > 0 ? Math.min(100, Math.round((p.doneQty / p.totalQty) * 100)) : 0
  }));
}


// ============================================================
// Web App：POST — 接收表單送出
// ============================================================
function doPost(e) {
  try {
    const p     = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DAILY);
    const ts    = new Date();

    // 儲存照片並取得資料夾網址
    let photoUrl = '';
    if (p.photos && p.photos.length > 0) {
      photoUrl = _savePhotos(p.caseId, p.date, p.photos);
    }

    // 每個工項一列
    const items = p.items && p.items.length > 0 ? p.items : [{ category: '', item: '', qty: '' }];
    items.forEach(item => {
      sheet.appendRow([ts, p.date, p.caseId, p.caseName,
        item.category, item.item, item.qty,
        p.workers || '', p.notes || '', photoUrl]);
    });

    return _json({ success: true });
  } catch (err) {
    return _json({ success: false, error: err.message });
  }
}

function _savePhotos(caseId, date, photos) {
  const root       = DriveApp.getFolderById(PHOTO_FOLDER_ID);
  const caseFolders = root.getFoldersByName(caseId);
  const caseFolder = caseFolders.hasNext() ? caseFolders.next() : root.createFolder(caseId);
  const dateFolders = caseFolder.getFoldersByName(date);
  const dateFolder = dateFolders.hasNext() ? dateFolders.next() : caseFolder.createFolder(date);

  photos.forEach((photo, i) => {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(photo.data),
      photo.mimeType,
      `${date}_${String(i + 1).padStart(2, '0')}.jpg`
    );
    dateFolder.createFile(blob);
  });

  return dateFolder.getUrl();
}


// ── 工具函式 ──────────────────────────────────────────────
function _convertToSheets(fileId, title) {
  const blob     = DriveApp.getFileById(fileId).getBlob();
  const resource = { title, mimeType: MimeType.GOOGLE_SHEETS };
  return Drive.Files.insert(resource, blob, { convert: true }).id;
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
