// ================================================================
// MeetingExecutionService.gs
// 處理「家事會面執行」分流，負責會面執行專屬的資料讀寫。
// 紀錄存放於「會面執行_YYYY」工作表。
// ================================================================

// ==================== 欄位定義 ====================

const MEETING_EXECUTION_FIELDS = [
  'RecordId', '執行日期', '會面分案案號', '執行方式', '案件類型', '執行人員',
  '未成年子女(男)', '未成年子女(女)',
  '會面方性別',
  '共同會面者(男)', '共同會面者(女)',
  '建立時間', '修改時間', '建立者', '修改者'
];

// ==================== 年度分表輔助函數 ====================

/**
 * 根據年份取得或建立對應的會面執行 Sheet
 */
function getMeetingExecutionSheetByYear(year) {
  if (!spreadsheet) initializeSpreadsheet();

  const sheetName = '會面執行_' + year;
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    // 使用 Init.gs 中的初始化函數
    initMeetingExecutionSheet(year);
    sheet = spreadsheet.getSheetByName(sheetName);
  }

  return sheet;
}

// ==================== 查詢紀錄 ====================

/**
 * 取得會面執行紀錄列表 (包含條件篩選與分頁)
 * @param {Object} filters - 篩選條件 (year, month, 執行社工, 個案姓名等)
 * @param {number} page - 頁碼
 * @param {number} pageSize - 每頁筆數
 * @return {Object} { records: [], totalCount: number, page: number, totalPages: number }
 */
function getMeetingExecutionRecords(filters, page, pageSize) {
  checkUserPermission();

  page = page || 1;
  pageSize = pageSize || 12;
  filters = filters || {};

  if (!spreadsheet) initializeSpreadsheet();

  const year = filters.year || new Date().getFullYear();
  const sheet = getMeetingExecutionSheetByYear(year);

  if (!sheet || sheet.getLastRow() <= 1) {
    return { records: [], totalCount: 0, page: 1, totalPages: 0 };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let records = [];

  // 解析資料列轉為物件
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let record = { rowNumber: i + 1 }; // 紀錄原始列號
    headers.forEach((header, index) => {
      let value = row[index];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, 'Asia/Taipei', 'yyyy-MM-dd');
      }
      record[header] = value;
    });

    // 多選欄位轉換為陣列 (以分號分隔)
    const arrayFields = ['執行人員', '未成年子女(男)', '未成年子女(女)', '共同會面者(男)', '共同會面者(女)'];
    arrayFields.forEach(field => {
      if (record[field] && typeof record[field] === 'string') {
        record[field] = record[field].split(';').map(s => s.trim()).filter(s => s !== '');
      } else if (!record[field]) {
        record[field] = [];
      } else if (!Array.isArray(record[field])) {
        record[field] = [String(record[field])];
      }
    });

    records.push(record);
  }

  // 依建立時間降冪排序 (最新的在最前面)
  records.sort((a, b) => {
    const dateA = new Date(a['建立時間'] || 0);
    const dateB = new Date(b['建立時間'] || 0);
    return dateB - dateA;
  });

  // 應用篩選條件
  if (filters.month) {
    records = records.filter(r => {
      if (!r['執行日期']) return false;
      const recordMonth = new Date(r['執行日期']).getMonth() + 1;
      return recordMonth === filters.month;
    });
  }

  if (filters['執行社工']) {
    records = records.filter(r => {
      if (!r['執行人員']) return false;
      return Array.isArray(r['執行人員']) ? r['執行人員'].includes(filters['執行社工']) : r['執行人員'] === filters['執行社工'];
    });
  }

  if (filters['個案姓名']) {
    const searchName = filters['個案姓名'].toLowerCase();
    records = records.filter(r => {
      const childrenM = Array.isArray(r['未成年子女(男)']) ? r['未成年子女(男)'].join(' ') : '';
      const childrenF = Array.isArray(r['未成年子女(女)']) ? r['未成年子女(女)'].join(' ') : '';
      const caseNumber = r['會面分案案號'] || '';
      const allNames = `${childrenM} ${childrenF} ${caseNumber}`.toLowerCase();
      return allNames.includes(searchName);
    });
  }

  const totalCount = records.length;
  const totalPages = Math.ceil(totalCount / pageSize);

  // 分頁裁切
  const startIndex = (page - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedRecords = records.slice(startIndex, endIndex);

  return {
    records: paginatedRecords,
    totalCount: totalCount,
    page: page,
    totalPages: totalPages
  };
}

// ==================== 新增紀錄 ====================

/**
 * 新增一筆會面執行紀錄
 */
function createMeetingExecutionRecord(recordData) {
  checkUserPermission();
  const user = getCurrentUser();
  const now = new Date();

  const serviceDate = recordData['執行日期'] ? new Date(recordData['執行日期']) : now;
  const year = serviceDate.getFullYear();
  const sheet = getMeetingExecutionSheetByYear(year);

  const recordId = Utilities.getUuid();
  const timestamp = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

  // 將陣列轉為分號分隔字串
  const arrayFields = ['執行人員', '未成年子女(男)', '未成年子女(女)', '共同會面者(男)', '共同會面者(女)'];
  arrayFields.forEach(field => {
    if (Array.isArray(recordData[field])) {
      recordData[field] = recordData[field].join('; ');
    }
  });

  const rowData = [
    recordId,                              // RecordId
    recordData['執行日期'] || '',           // 執行日期
    recordData['會面分案案號'] || '',       // 會面分案案號
    recordData['執行方式'] || '',           // 執行方式
    recordData['案件類型'] || '',           // 案件類型
    recordData['執行人員'] || '',           // 執行人員
    recordData['未成年子女(男)'] || '',     // 未成年子女(男)
    recordData['未成年子女(女)'] || '',     // 未成年子女(女)
    recordData['會面方性別'] || '',         // 會面方性別
    recordData['共同會面者(男)'] || '',     // 共同會面者(男)
    recordData['共同會面者(女)'] || '',     // 共同會面者(女)
    timestamp,                             // 建立時間
    timestamp,                             // 修改時間
    user.email,                            // 建立者
    user.email                             // 修改者
  ];

  sheet.appendRow(rowData);

  console.log('✅ 新增會面執行紀錄成功，RecordId:', recordId);
  return { success: true, message: '會面執行紀錄已成功新增', recordId: recordId };
}

// ==================== 更新紀錄 ====================

/**
 * 更新會面執行紀錄
 */
function updateMeetingExecutionRecord(recordId, recordData) {
  checkUserPermission();
  const user = getCurrentUser();
  const now = new Date();

  const found = findMeetingExecutionRecordInAllYears(recordId);
  if (found.rowIndex === -1) {
    throw new Error('找不到指定的會面執行紀錄 ID');
  }

  const sheet = found.sheet;
  const rowIndex = found.rowIndex;
  const existingRow = sheet.getRange(rowIndex, 1, 1, MEETING_EXECUTION_FIELDS.length).getValues()[0];

  const timestamp = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

  // 將陣列轉為分號分隔字串
  const arrayFields = ['執行人員', '未成年子女(男)', '未成年子女(女)', '共同會面者(男)', '共同會面者(女)'];
  arrayFields.forEach(field => {
    if (Array.isArray(recordData[field])) {
      recordData[field] = recordData[field].join('; ');
    }
  });

  const updatedRow = [
    recordId,                                                                         // RecordId
    recordData['執行日期'] !== undefined ? recordData['執行日期'] : existingRow[1],
    recordData['會面分案案號'] !== undefined ? recordData['會面分案案號'] : existingRow[2],
    recordData['執行方式'] !== undefined ? recordData['執行方式'] : existingRow[3],
    recordData['案件類型'] !== undefined ? recordData['案件類型'] : existingRow[4],
    recordData['執行人員'] !== undefined ? recordData['執行人員'] : existingRow[5],
    recordData['未成年子女(男)'] !== undefined ? recordData['未成年子女(男)'] : existingRow[6],
    recordData['未成年子女(女)'] !== undefined ? recordData['未成年子女(女)'] : existingRow[7],
    recordData['會面方性別'] !== undefined ? recordData['會面方性別'] : existingRow[8],
    recordData['共同會面者(男)'] !== undefined ? recordData['共同會面者(男)'] : existingRow[9],
    recordData['共同會面者(女)'] !== undefined ? recordData['共同會面者(女)'] : existingRow[10],
    existingRow[11],                                                                  // 建立時間 (不變)
    timestamp,                                                                        // 修改時間
    existingRow[13],                                                                  // 建立者 (不變)
    user.email                                                                        // 修改者
  ];

  sheet.getRange(rowIndex, 1, 1, MEETING_EXECUTION_FIELDS.length).setValues([updatedRow]);

  console.log('✅ 更新會面執行紀錄成功，RecordId:', recordId);
  return { success: true, message: '會面執行紀錄更新成功' };
}

// ==================== 刪除紀錄 ====================

/**
 * 刪除會面執行紀錄
 */
function deleteMeetingExecutionRecord(recordId) {
  checkUserPermission();

  const found = findMeetingExecutionRecordInAllYears(recordId);
  if (found.rowIndex === -1) {
    throw new Error('找不到指定的會面執行紀錄 ID');
  }

  found.sheet.deleteRow(found.rowIndex);

  console.log('✅ 刪除會面執行紀錄成功');
  return { success: true, message: '會面執行紀錄已刪除' };
}

// ==================== 輔助函數 ====================

/**
 * 在所有年度的會面執行表中搜尋指定 RecordId
 */
function findMeetingExecutionRecordInAllYears(recordId) {
  if (!spreadsheet) initializeSpreadsheet();
  const searchId = String(recordId).trim();

  for (const sheet of spreadsheet.getSheets()) {
    const name = sheet.getName();
    if (!name.startsWith('會面執行_') || sheet.getLastRow() <= 1) continue;

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchId) {
        return { sheet, rowIndex: i + 2 };
      }
    }
  }
  return { sheet: null, rowIndex: -1 };
}
