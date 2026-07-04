/**
 * LINE LIFF 打卡系統 - Google Apps Script 後端
 */

// ===== 設定區 =====
const SHEET_LOGS = 'Logs';
const SHEET_SUMMARY = 'Summary';
const SHEET_EXPENSES = 'Expenses';
const SHEET_LEAVES = 'Leaves';
const SHEET_WAREHOUSE = 'Warehouse';
const STANDARD_DAILY_HOURS = 8;
const TIMEZONE = 'Asia/Taipei';
// ==================

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    const { userId, name, type, timestamp } = data;

    if (!userId || !type || !timestamp) {
      return jsonResponse({ success: false, message: '缺少必要欄位' });
    }

    // ── 代墊 / 拿貨付款 ──
    if (type === 'EXPENSE') {
      return logExpense(data);
    }

    // ── 代墊刪除（標記已刪除）──
    if (type === 'MARK_DELETE_EXPENSE') {
      return markDeleteExpense(data);
    }

    // ── 請假申請 ──
    if (type === 'LEAVE') {
      return logLeave(data);
    }

    // ── 倉庫位置配置（新增/更新）──
    if (type === 'WAREHOUSE_UPSERT') {
      return warehouseUpsert(data);
    }

    // ── 打卡 (IN / OUT) ──
    if (type !== 'IN' && type !== 'OUT') {
      return jsonResponse({ success: false, message: 'type 必須是 IN、OUT、EXPENSE 或 LEAVE' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LOGS);

    const ts = new Date(timestamp);
    const dateStr = Utilities.formatDate(ts, TIMEZONE, 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(ts, TIMEZONE, 'HH:mm:ss');

    sheet.appendRow([
      ts,                  // A 完整時間戳記
      dateStr,             // B 日期
      timeStr,             // C 時間
      userId,              // D LINE userId
      name || '',          // E 姓名
      type,                // F IN / OUT
      data.lat || '',      // G 緯度
      data.lng || '',      // H 經度
      data.accuracy || '', // I 定位誤差(公尺)
      data.address || ''   // J 打卡地址
    ]);

    updateSummaryForDay(dateStr, userId, name);
    updatePersonSheet(userId, name);

    return jsonResponse({ success: true });

  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const params = e.parameter;
  if (params.action === 'monthly' && params.userId && params.month) {
    return getMonthlyReport(params.userId, params.month);
  }
  if (params.action === 'status' && params.userId) {
    return getTodayStatus(params.userId);
  }
  if (params.action === 'leaves' && params.userId && params.month) {
    return getLeavesForMonth(params.userId, params.month);
  }
  if (params.action === 'expenses' && params.userId && params.date) {
    return getExpensesForDate(params.userId, params.date);
  }
  if (params.action === 'warehouse_search' && params.q) {
    return warehouseSearch(params.q);
  }
  if (params.action === 'warehouse_list') {
    return warehouseList();
  }
  return jsonResponse({ status: 'LINE 打卡系統 API 運作中' });
}

/**
 * 查詢員工今天的打卡狀態
 * 回傳 { success, status: 'none'|'in'|'out', inTime, outTime }
 */
function getTodayStatus(userId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(SHEET_LOGS);
    const data = logSheet.getDataRange().getValues();
    const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');

    const records = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[3]) !== String(userId)) continue;
      const rowDate = Utilities.formatDate(new Date(row[0]), TIMEZONE, 'yyyy-MM-dd');
      if (rowDate !== today) continue;
      records.push({ time: new Date(row[0]), type: String(row[5]) });
    }
    records.sort((a, b) => a.time - b.time);

    let pendingIn = null, firstIn = null, lastOut = null;
    records.forEach(r => {
      if (r.type === 'IN') { if (!firstIn) firstIn = r.time; pendingIn = r.time; }
      else if (r.type === 'OUT' && pendingIn) { lastOut = r.time; pendingIn = null; }
    });

    if (!firstIn) {
      return jsonResponse({ success: true, status: 'none' });
    } else if (pendingIn) {
      return jsonResponse({
        success: true, status: 'in',
        inTime: Utilities.formatDate(pendingIn, TIMEZONE, 'HH:mm:ss')
      });
    } else {
      return jsonResponse({
        success: true, status: 'out',
        inTime: Utilities.formatDate(firstIn, TIMEZONE, 'HH:mm:ss'),
        outTime: Utilities.formatDate(lastOut, TIMEZONE, 'HH:mm:ss')
      });
    }
  } catch(err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

/**
 * 月報表：直接從 Logs 計算，不依賴 Summary（避免 Summary 重複列問題）
 */
function getMonthlyReport(userId, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(SHEET_LOGS);
    const data = logSheet.getDataRange().getValues();

    // 依日期分組
    const dayMap = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[3]) !== String(userId)) continue;

      const rowDate = Utilities.formatDate(new Date(row[0]), TIMEZONE, 'yyyy-MM-dd');
      if (!rowDate.startsWith(month)) continue;

      if (!dayMap[rowDate]) {
        dayMap[rowDate] = { records: [], name: String(row[4]), address: '' };
      }
      dayMap[rowDate].records.push({ time: new Date(row[0]), type: String(row[5]) });
      if (String(row[5]) === 'IN' && !dayMap[rowDate].address && row[9]) {
        dayMap[rowDate].address = String(row[9]);
      }
    }

    const rows = [];
    Object.keys(dayMap).sort().forEach(date => {
      const { records, name, address } = dayMap[date];
      records.sort((a, b) => a.time - b.time);

      let totalMs = 0, pendingIn = null, firstIn = null, lastOut = null;
      records.forEach(r => {
        if (r.type === 'IN') {
          if (!firstIn) firstIn = r.time;
          pendingIn = r.time;
        } else if (r.type === 'OUT' && pendingIn) {
          totalMs += (r.time - pendingIn);
          lastOut = r.time;
          pendingIn = null;
        }
      });

      const totalHours = Math.round(totalMs / (1000 * 60 * 60) * 10) / 10;
      const overtimeHours = Math.round(Math.max(0, totalHours - STANDARD_DAILY_HOURS) * 10) / 10;
      const status = pendingIn ? '尚未下班' : (firstIn ? '已完成' : '');

      rows.push({
        date,
        inTime: firstIn ? Utilities.formatDate(firstIn, TIMEZONE, 'HH:mm:ss') : '',
        outTime: lastOut ? Utilities.formatDate(lastOut, TIMEZONE, 'HH:mm:ss') : '',
        totalHours,
        overtimeHours,
        status,
        address
      });
    });

    return jsonResponse({ success: true, data: rows });
  } catch(err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 更新 Summary 分頁（含 dedup：找到所有相同日期/userId 的列，保留第一列並刪除多餘列）
 */
function updateSummaryForDay(dateStr, userId, name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_LOGS);
  const data = logSheet.getDataRange().getValues();

  const records = [];
  let inAddress = '';
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDate = Utilities.formatDate(new Date(row[0]), TIMEZONE, 'yyyy-MM-dd');
    if (rowDate === dateStr && String(row[3]) === String(userId)) {
      records.push({ time: new Date(row[0]), type: String(row[5]) });
      if (String(row[5]) === 'IN' && !inAddress && row[9]) {
        inAddress = String(row[9]);
      }
    }
  }
  records.sort((a, b) => a.time - b.time);

  let totalMs = 0, pendingIn = null, firstIn = null, lastOut = null;
  records.forEach(r => {
    if (r.type === 'IN') {
      if (!firstIn) firstIn = r.time;
      pendingIn = r.time;
    } else if (r.type === 'OUT' && pendingIn) {
      totalMs += (r.time - pendingIn);
      lastOut = r.time;
      pendingIn = null;
    }
  });

  const totalHours = totalMs / (1000 * 60 * 60);
  const overtimeHours = Math.max(0, totalHours - STANDARD_DAILY_HOURS);
  const status = pendingIn ? '尚未下班' : '已完成';

  const summarySheet = ss.getSheetByName(SHEET_SUMMARY);
  const sData = summarySheet.getDataRange().getValues();

  // 找出所有相同日期+userId 的列（1-indexed）
  const matchRows = [];
  for (let i = 1; i < sData.length; i++) {
    const sRowDate = sData[i][0] instanceof Date
      ? Utilities.formatDate(new Date(sData[i][0]), TIMEZONE, 'yyyy-MM-dd')
      : String(sData[i][0]);
    if (sRowDate === dateStr && String(sData[i][1]) === String(userId)) {
      matchRows.push(i + 1);
    }
  }

  const rowValues = [
    dateStr,
    userId,
    name || '',
    firstIn ? Utilities.formatDate(firstIn, TIMEZONE, 'HH:mm:ss') : '',
    lastOut ? Utilities.formatDate(lastOut, TIMEZONE, 'HH:mm:ss') : '',
    Math.round(totalHours * 10) / 10,
    Math.round(overtimeHours * 10) / 10,
    status,
    inAddress
  ];

  if (matchRows.length > 0) {
    // 更新第一列
    summarySheet.getRange(matchRows[0], 1, 1, rowValues.length).setValues([rowValues]);
    // 從後往前刪除多餘列
    for (let i = matchRows.length - 1; i >= 1; i--) {
      summarySheet.deleteRow(matchRows[i]);
    }
  } else {
    summarySheet.appendRow(rowValues);
  }
}

/**
 * 更新單一員工的個人分頁（原始打卡紀錄 + 每日彙總）
 */
function updatePersonSheet(userId, name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_LOGS);
  const logs = logSheet.getDataRange().getValues();

  // 收集該員工所有打卡紀錄
  const punches = [];
  const dayMap = {};
  for (let i = 1; i < logs.length; i++) {
    const row = logs[i];
    if (String(row[3]) !== String(userId)) continue;
    const ts = new Date(row[0]);
    const rowDate = Utilities.formatDate(ts, TIMEZONE, 'yyyy-MM-dd');
    const rowTime = Utilities.formatDate(ts, TIMEZONE, 'HH:mm:ss');
    const type = String(row[5]);
    const address = String(row[9] || '');
    punches.push([rowDate, rowTime, type === 'IN' ? '上班' : '下班', address]);

    if (!dayMap[rowDate]) dayMap[rowDate] = { records: [], address: '' };
    dayMap[rowDate].records.push({ time: ts, type });
    if (type === 'IN' && !dayMap[rowDate].address && row[9]) {
      dayMap[rowDate].address = address;
    }
  }
  punches.sort((a, b) => (a[0] + a[1]).localeCompare(b[0] + b[1]));

  // 建立或取得個人分頁
  const sheetName = (name || userId.slice(0, 10)).replace(/[\\\/\?\*\[\]:]/g, '_');
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clearContents();

  // === 區塊一：原始打卡紀錄 ===
  sheet.getRange(1, 1).setValue('【打卡紀錄】');
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(12);
  sheet.getRange(2, 1, 1, 4).setValues([['日期', '時間', '類型', '地址']]);
  sheet.getRange(2, 1, 1, 4).setFontWeight('bold').setBackground('#06C755').setFontColor('#ffffff');
  if (punches.length > 0) {
    sheet.getRange(3, 1, punches.length, 4).setValues(punches);
  }

  // === 區塊二：每日彙總（在打卡紀錄下方，空一列）===
  const summaryStart = 3 + punches.length + 2;
  sheet.getRange(summaryStart, 1).setValue('【每日彙總】');
  sheet.getRange(summaryStart, 1).setFontWeight('bold').setFontSize(12);
  sheet.getRange(summaryStart + 1, 1, 1, 6).setValues([['日期', '上班', '下班', '工時(h)', '狀態', '地址']]);
  sheet.getRange(summaryStart + 1, 1, 1, 6).setFontWeight('bold').setBackground('#4A90D9').setFontColor('#ffffff');

  const sortedDates = Object.keys(dayMap).sort();
  const summaryRows = sortedDates.map(date => {
    const { records, address } = dayMap[date];
    records.sort((a, b) => a.time - b.time);
    let totalMs = 0, pendingIn = null, firstIn = null, lastOut = null;
    records.forEach(r => {
      if (r.type === 'IN') { if (!firstIn) firstIn = r.time; pendingIn = r.time; }
      else if (r.type === 'OUT' && pendingIn) { totalMs += r.time - pendingIn; lastOut = r.time; pendingIn = null; }
    });
    const totalHours = Math.round(totalMs / 36e5 * 10) / 10;
    return [
      date,
      firstIn ? Utilities.formatDate(firstIn, TIMEZONE, 'HH:mm') : '',
      lastOut ? Utilities.formatDate(lastOut, TIMEZONE, 'HH:mm') : '',
      totalHours,
      pendingIn ? '尚未下班' : '已完成',
      address
    ];
  });

  if (summaryRows.length > 0) {
    sheet.getRange(summaryStart + 2, 1, summaryRows.length, 6).setValues(summaryRows);
  }

  // 自動調整欄寬
  sheet.autoResizeColumns(1, 6);
}

/**
 * Logs 分頁依員工姓名排序，再依時間排序
 */
function sortLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOGS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  // 先依姓名（E欄=第5欄），再依時間戳記（A欄=第1欄）
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .sort([{ column: 5, ascending: true }, { column: 1, ascending: true }]);
}

/**
 * 一次性建立所有員工的個人分頁（手動執行一次）
 */
function buildAllPersonSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_LOGS);
  const logs = logSheet.getDataRange().getValues();

  const users = {};
  for (let i = 1; i < logs.length; i++) {
    const userId = String(logs[i][3]);
    const name = String(logs[i][4]);
    if (userId && !users[userId]) users[userId] = name;
  }

  Object.keys(users).forEach(userId => {
    updatePersonSheet(userId, users[userId]);
  });

  Logger.log('完成，共建立 ' + Object.keys(users).length + ' 個員工分頁');
}

/**
 * 全部重建：排序 Logs + 重建所有員工分頁 + 清理 Summary 重複列
 */
function rebuildAll() {
  sortLogs();
  cleanupSummaryDuplicates();
  buildAllPersonSheets();
  Logger.log('全部重建完成');
}

/**
 * 一次性清理 Summary 重複列（手動執行一次即可）
 */
function cleanupSummaryDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName(SHEET_SUMMARY);
  const data = summarySheet.getDataRange().getValues();

  // 從 Logs 重新計算所有人所有日期
  const logSheet = ss.getSheetByName(SHEET_LOGS);
  const logs = logSheet.getDataRange().getValues();

  const keys = {}; // "date|userId" → name
  for (let i = 1; i < logs.length; i++) {
    const row = logs[i];
    const rowDate = Utilities.formatDate(new Date(row[0]), TIMEZONE, 'yyyy-MM-dd');
    const key = rowDate + '|' + String(row[3]);
    if (!keys[key]) keys[key] = String(row[4]);
  }

  // 清空 Summary（保留表頭）
  if (data.length > 1) {
    summarySheet.deleteRows(2, data.length - 1);
  }

  // 重新計算並寫入（每個 date+userId 只有一列）
  Object.keys(keys).sort().forEach(key => {
    const [dateStr, userId] = key.split('|');
    const name = keys[key];
    updateSummaryForDay(dateStr, userId, name);
  });

  Logger.log('清理完成，共處理 ' + Object.keys(keys).length + ' 組記錄');
}

/**
 * 初始化分頁與表頭，只需手動執行一次。
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let logSheet = ss.getSheetByName(SHEET_LOGS);
  if (!logSheet) logSheet = ss.insertSheet(SHEET_LOGS);
  logSheet.clear();
  logSheet.appendRow(['時間戳記', '日期', '時間', 'LINE userId', '姓名', '類型(IN/OUT)', '緯度', '經度', '定位誤差(m)', '打卡地址']);
  logSheet.setFrozenRows(1);

  let summarySheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!summarySheet) summarySheet = ss.insertSheet(SHEET_SUMMARY);
  summarySheet.clear();
  summarySheet.appendRow(['日期', 'LINE userId', '姓名', '上班時間', '下班時間', '總工時(小時)', '加班時數(小時)', '狀態', '上班打卡地址']);
  summarySheet.setFrozenRows(1);

  SpreadsheetApp.flush();
  Logger.log('分頁設定完成');
}

// =====================================================
// 代墊 / 拿貨付款
// =====================================================

/**
 * 記錄代墊或拿貨付款到 Expenses 工作表
 */
function logExpense(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 自動建立 Expenses 工作表（若不存在）
  let sheet = ss.getSheetByName(SHEET_EXPENSES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EXPENSES);
    sheet.appendRow(['申請時間', '日期', '時間', 'LINE userId', '姓名', '類型', '金額(元)', '說明/備註']);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 8);
  }

  const ts      = new Date(data.timestamp);
  const dateStr = Utilities.formatDate(ts, TIMEZONE, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(ts, TIMEZONE, 'HH:mm:ss');

  sheet.appendRow([
    ts,                    // A 申請時間
    dateStr,               // B 日期
    timeStr,               // C 時間
    data.userId,           // D LINE userId
    data.name || '',       // E 姓名
    data.expType || '',    // F 類型（代墊 / 拿貨付款）
    data.amount || 0,      // G 金額
    data.note || '',       // H 說明
    ''                     // I 狀態（空=正常，已刪除=刪除）
  ]);

  return jsonResponse({ success: true });
}

/**
 * 查詢某天的代墊申請（排除已刪除）
 * GET ?action=expenses&userId=xxx&date=yyyy-mm-dd
 */
function getExpensesForDate(userId, date) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_EXPENSES);
    if (!sheet) return jsonResponse({ success: true, expenses: [] });

    const data     = sheet.getDataRange().getValues();
    const expenses = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[3]) !== String(userId)) continue;

      // 比對日期（B欄可能是 Date 物件或字串）
      const rowDate = row[1] instanceof Date
        ? Utilities.formatDate(new Date(row[1]), TIMEZONE, 'yyyy-MM-dd')
        : String(row[1]);
      if (rowDate !== date) continue;

      // 排除已刪除
      if (String(row[8]) === '已刪除') continue;

      expenses.push({
        timestamp: new Date(row[0]).toISOString(),
        time:      String(row[2]),
        expType:   String(row[5]),
        amount:    row[6],
        note:      String(row[7])
      });
    }

    return jsonResponse({ success: true, expenses });
  } catch(err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

/**
 * 標記代墊申請為「已刪除」
 * POST { type: 'MARK_DELETE_EXPENSE', userId, timestamp }
 */
function markDeleteExpense(data) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_EXPENSES);
    if (!sheet) return jsonResponse({ success: false, message: 'Expenses 工作表不存在' });

    const rows     = sheet.getDataRange().getValues();
    const targetMs = new Date(data.timestamp).getTime();

    for (let i = 1; i < rows.length; i++) {
      const rowMs = new Date(rows[i][0]).getTime();
      if (String(rows[i][3]) === String(data.userId) && Math.abs(rowMs - targetMs) < 2000) {
        sheet.getRange(i + 1, 9).setValue('已刪除');
        return jsonResponse({ success: true });
      }
    }

    return jsonResponse({ success: false, message: '找不到對應記錄' });
  } catch(err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

// =====================================================
// 請假申請
// =====================================================

/**
 * 記錄請假到 Leaves 工作表
 */
function logLeave(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 自動建立 Leaves 工作表（若不存在）
  let sheet = ss.getSheetByName(SHEET_LEAVES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LEAVES);
    sheet.appendRow(['申請時間', 'LINE userId', '姓名', '開始日期', '結束日期', '請假天數']);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 6);
  }

  const ts = new Date(data.timestamp);

  sheet.appendRow([
    ts,                    // A 申請時間
    data.userId,           // B LINE userId
    data.name || '',       // C 姓名
    data.startDate || '',  // D 開始日期
    data.endDate   || '',  // E 結束日期
    data.days      || 0    // F 請假天數
  ]);

  return jsonResponse({ success: true });
}

/**
 * 查詢某月的請假記錄（供 leave.html 顯示歷史）
 * GET ?action=leaves&userId=xxx&month=2026-07
 */
function getLeavesForMonth(userId, month) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LEAVES);
    if (!sheet) return jsonResponse({ success: true, leaves: [] });

    const data   = sheet.getDataRange().getValues();
    const leaves = [];

    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      if (String(row[1]) !== String(userId)) continue;
      // 判斷開始日期是否在該月
      const start = String(row[3]);
      if (!start.startsWith(month)) continue;
      leaves.push({
        startDate: String(row[3]),
        endDate:   String(row[4]),
        days:      row[5]
      });
    }

    return jsonResponse({ success: true, leaves });
  } catch(err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

// =====================================================
// 倉庫位置配置
// =====================================================

/**
 * 新增或更新商品的倉庫位置（以商品名稱做 upsert）
 * POST { type: 'WAREHOUSE_UPSERT', itemName, itemCode, zone, shelf, note, userId, name, timestamp }
 */
function warehouseUpsert(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 自動建立 Warehouse 工作表
    let sheet = ss.getSheetByName(SHEET_WAREHOUSE);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_WAREHOUSE);
      sheet.appendRow(['更新時間', '商品名稱', '商品編號', '倉庫區域', '貨架位置', '備註', '更新者userId', '更新者姓名', '更新日期', '更新時間']);
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, 10);
    }

    const ts      = new Date(data.timestamp);
    const dateStr = Utilities.formatDate(ts, TIMEZONE, 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(ts, TIMEZONE, 'HH:mm:ss');
    const name    = String(data.itemName || '').trim();

    if (!name) return jsonResponse({ success: false, message: '商品名稱不可為空' });

    // 找現有列（依商品名稱，大小寫不敏感）
    const rows = sheet.getDataRange().getValues();
    let matchRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim().toLowerCase() === name.toLowerCase()) {
        matchRow = i + 1; // 1-indexed
        break;
      }
    }

    const rowValues = [
      ts,                       // A 更新時間
      name,                     // B 商品名稱
      data.itemCode || '',      // C 商品編號
      data.zone     || '',      // D 倉庫區域
      data.shelf    || '',      // E 貨架位置
      data.note     || '',      // F 備註
      data.userId,              // G 更新者userId
      data.name     || '',      // H 更新者姓名
      dateStr,                  // I 更新日期
      timeStr                   // J 更新時間
    ];

    if (matchRow > 0) {
      sheet.getRange(matchRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return jsonResponse({ success: true, updated: matchRow > 0 });
  } catch(err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

/**
 * 搜尋商品（商品名稱或編號模糊比對）
 * GET ?action=warehouse_search&q=關鍵字
 */
function warehouseSearch(q) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_WAREHOUSE);
    if (!sheet) return jsonResponse({ success: true, items: [] });

    const rows  = sheet.getDataRange().getValues();
    const query = String(q).trim().toLowerCase();
    const items = [];

    for (let i = 1; i < rows.length; i++) {
      const row  = rows[i];
      const name = String(row[1]).toLowerCase();
      const code = String(row[2]).toLowerCase();
      if (name.includes(query) || code.includes(query)) {
        items.push(rowToWarehouseItem(row));
      }
    }

    return jsonResponse({ success: true, items });
  } catch(err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

/**
 * 取得全部商品清單
 * GET ?action=warehouse_list
 */
function warehouseList() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_WAREHOUSE);
    if (!sheet) return jsonResponse({ success: true, items: [] });

    const rows  = sheet.getDataRange().getValues();
    const items = [];
    for (let i = 1; i < rows.length; i++) {
      items.push(rowToWarehouseItem(rows[i]));
    }

    return jsonResponse({ success: true, items });
  } catch(err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function rowToWarehouseItem(row) {
  return {
    name:      String(row[1]),
    code:      String(row[2]),
    zone:      String(row[3]),
    shelf:     String(row[4]),
    note:      String(row[5]),
    updatedBy: String(row[7]),
    updatedAt: String(row[8]) + ' ' + String(row[9])
  };
}

function rebuildYesterdaySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_LOGS);
  const data = logSheet.getDataRange().getValues();

  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const dateStr = Utilities.formatDate(yesterday, TIMEZONE, 'yyyy-MM-dd');

  const userMap = {};
  for (let i = 1; i < data.length; i++) {
    const rowDate = Utilities.formatDate(new Date(data[i][0]), TIMEZONE, 'yyyy-MM-dd');
    if (rowDate === dateStr) {
      userMap[data[i][3]] = data[i][4];
    }
  }

  Object.keys(userMap).forEach(userId => {
    updateSummaryForDay(dateStr, userId, userMap[userId]);
  });
}
