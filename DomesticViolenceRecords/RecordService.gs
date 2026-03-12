// ================================================================
// RecordService.gs - 服務紀錄 CRUD
// ================================================================

// ==================== 欄位定義 ====================

const RECORD_FIELDS = [
  'RecordId', '服務日期', '個案來源', '個案姓名', '個案身分', '個案國籍',
  '暴力類型', '個案性別', '有無在案', '在案社工', '接受轉介項目', '轉案股別/單位',
  '服務社工', '服務方式', '服務主題', 
  '陪同出庭', '會談處遇', '法律服務',  // 2025 以前舊欄位
  '聯繫', '轉介', '通報', 
  '人身安全', '會談服務',  // 2026 新欄位
  '服務項目', '建立時間', '修改時間', '建立者', '修改者'
];

// ==================== 年度分表輔助函數 ====================

/**
 * 根據年份取得或建立對應的服務紀錄 Sheet
 * @param {number|string} year - 年份 (如 2025)
 * @returns {Sheet} 對應年度的 Sheet
 */
function getSheetByYear(year) {
  if (!spreadsheet) initializeSpreadsheet();
  
  const sheetName = '服務紀錄_' + year;
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    // 建立新的年度 Sheet
    sheet = spreadsheet.insertSheet(sheetName);
    
    // 設定表頭
    sheet.getRange(1, 1, 1, RECORD_FIELDS.length).setValues([RECORD_FIELDS]);
    
    // 設定表頭格式
    const headerRange = sheet.getRange(1, 1, 1, RECORD_FIELDS.length);
    headerRange.setBackground('#2c3e50');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    console.log('已建立年度 Sheet:', sheetName);
  }
  
  return sheet;
}

/**
 * 從服務日期取得年份
 * @param {Date|string} dateValue - 日期 (Date 物件或字串)
 * @returns {number} 年份
 */
function getYearFromDate(dateValue) {
  if (!dateValue) return new Date().getFullYear();
  
  // 如果是 Date 物件，直接取年份
  if (dateValue instanceof Date) {
    return dateValue.getFullYear();
  }
  
  // 嘗試解析字串
  const dateStr = String(dateValue);
  
  // 嘗試 YYYY-MM-DD 格式
  if (/^\d{4}-\d{2}-\d{2}/.test(dateStr)) {
    return parseInt(dateStr.substring(0, 4));
  }
  
  // 嘗試 YYYY/MM/DD 格式
  if (/^\d{4}\/\d{2}\/\d{2}/.test(dateStr)) {
    return parseInt(dateStr.substring(0, 4));
  }
  
  // 嘗試中華民國年格式 (如 113/01/15)
  if (/^\d{2,3}\/\d{2}\/\d{2}/.test(dateStr)) {
    const rocYear = parseInt(dateStr.split('/')[0]);
    return rocYear + 1911;
  }
  
  // 嘗試用 Date 解析
  const parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) {
    return parsed.getFullYear();
  }
  
  console.log('getYearFromDate: 無法解析日期:', dateValue);
  return new Date().getFullYear();
}

// ==================== 新增紀錄 ====================

/**
 * 新增服務紀錄
 * @param {Object} recordData - 紀錄資料
 * @returns {Object} { success: boolean, recordId: string, message: string }
 */
function createRecord(recordData) {
  return executeWithRetry(function() {
    // 1. 權限驗證
    checkUserPermission();
    const user = getCurrentUser();
    
    // 2. 資料驗證
    validateRecordData(recordData);
    
    // 3. 產生 RecordId
    const recordId = Utilities.getUuid();
    
    // 4. 根據服務日期取得對應年度的 Sheet
    const serviceDate = recordData['服務日期'] || '';
    const year = getYearFromDate(serviceDate);
    const sheet = getSheetByYear(year);
    
    const now = new Date();
    const rowData = [
      recordId,                           // RecordId
      recordData['服務日期'] || '',       // 服務日期
      recordData['個案來源'] || '',       // 個案來源
      recordData['個案姓名'] || '',       // 個案姓名
      recordData['個案身分'] || '',       // 個案身分
      recordData['個案國籍'] || '',       // 個案國籍
      recordData['暴力類型'] || '',       // 暴力類型
      recordData['個案性別'] || '',       // 個案性別
      recordData['有無在案'] ? '是' : '否', // 有無在案
      recordData['在案社工'] || '',       // 在案社工
      recordData['接受轉介項目'] || '',   // 接受轉介項目
      recordData['轉案股別/單位'] || '',  // 轉案股別/單位
      recordData['服務社工'] || '',       // 服務社工
      recordData['服務方式'] || '',       // 服務方式
      recordData['服務主題'] || '',       // 服務主題
      formatMultiChoice(recordData['陪同出庭']),  // 陪同出庭 (多選)
      formatMultiChoice(recordData['會談處遇']),  // 會談處遇 (多選) - 2025 以前
      formatMultiChoice(recordData['法律服務']),  // 法律服務 (多選)
      formatMultiChoice(recordData['聯繫']),      // 聯繫 (多選)
      formatMultiChoice(recordData['轉介']),      // 轉介 (多選)
      formatMultiChoice(recordData['通報']),      // 通報 (多選)
      formatMultiChoice(recordData['人身安全']),  // 人身安全 (多選) - 2026 新增
      formatMultiChoice(recordData['會談服務']),  // 會談服務 (多選) - 2026 新增
      recordData['服務項目'] || '',       // 服務項目
      Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),  // 建立時間
      Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),  // 修改時間
      user.email,                         // 建立者
      user.email                          // 修改者
    ];
    
    // 5. 寫入
    sheet.appendRow(rowData);
    
    // 6. 更新時間戳
    updateLastModifiedTime('服務紀錄');
    
    // 7. 清除快取
    clearRecordsCache();
    
    console.log('✅ 新增服務紀錄成功，RecordId:', recordId);
    return { success: true, recordId: recordId, message: '新增成功' };
  });
}

// ==================== 更新紀錄 ====================

/**
 * 更新服務紀錄
 * @param {string} recordId - 紀錄 ID
 * @param {Object} updateData - 更新資料
 * @returns {Object} { success: boolean, message: string }
 */
function updateRecord(recordId, updateData) {
  return executeWithRetry(function() {
    // 1. 權限驗證
    checkUserPermission();
    const user = getCurrentUser();
    
    // 2. 在所有年度 Sheet 中尋找該筆紀錄
    const found = findRecordInAllYears(recordId);
    
    if (found.rowIndex === -1) {
      throw new Error('找不到該筆紀錄');
    }
    
    const sheet = found.sheet;
    const rowIndex = found.rowIndex;
    
    // 3. 檢查並遷移 Sheet 欄位結構 (如果需要)
    checkAndMigrateSheetColumns(sheet);
    
    // 4. 讀取現有資料 (現在確保有 28 欄了)
    const existingRow = sheet.getRange(rowIndex, 1, 1, RECORD_FIELDS.length).getValues()[0];
    
    // 5. 合併更新
    const now = new Date();
    
    // 根據目前的 RECORD_FIELDS 定義 (28欄) 組裝資料
    // 注意：如果是剛遷移過來的舊資料，existingRow[21] (人身安全) 和 [22] (會談服務) 會是空字串，原本的 [21] (服務項目) 會被移到 [23]
    
    const updatedRow = [
      recordId,                                      // 0: RecordId
      updateData['服務日期'] !== undefined ? updateData['服務日期'] : existingRow[1],      // 1: 服務日期
      updateData['個案來源'] !== undefined ? updateData['個案來源'] : existingRow[2],      // 2: 個案來源
      updateData['個案姓名'] !== undefined ? updateData['個案姓名'] : existingRow[3],      // 3: 個案姓名
      updateData['個案身分'] !== undefined ? updateData['個案身分'] : existingRow[4],      // 4: 個案身分
      updateData['個案國籍'] !== undefined ? updateData['個案國籍'] : existingRow[5],      // 5: 個案國籍
      updateData['暴力類型'] !== undefined ? updateData['暴力類型'] : existingRow[6],      // 6: 暴力類型
      updateData['個案性別'] !== undefined ? updateData['個案性別'] : existingRow[7],      // 7: 個案性別
      updateData['有無在案'] !== undefined ? (updateData['有無在案'] ? '是' : '否') : existingRow[8], // 8: 有無在案
      updateData['在案社工'] !== undefined ? updateData['在案社工'] : existingRow[9],   // 9: 在案社工
      updateData['接受轉介項目'] !== undefined ? updateData['接受轉介項目'] : existingRow[10], // 10: 接受轉介項目
      updateData['轉案股別/單位'] !== undefined ? updateData['轉案股別/單位'] : existingRow[11], // 11: 轉案股別/單位
      updateData['服務社工'] !== undefined ? updateData['服務社工'] : existingRow[12],     // 12: 服務社工
      updateData['服務方式'] !== undefined ? updateData['服務方式'] : existingRow[13],     // 13: 服務方式
      updateData['服務主題'] !== undefined ? updateData['服務主題'] : existingRow[14], // 14: 服務主題
      
      updateData['陪同出庭'] !== undefined ? formatMultiChoice(updateData['陪同出庭']) : existingRow[15], // 15: 陪同出庭
      updateData['會談處遇'] !== undefined ? formatMultiChoice(updateData['會談處遇']) : existingRow[16], // 16: 會談處遇
      updateData['法律服務'] !== undefined ? formatMultiChoice(updateData['法律服務']) : existingRow[17], // 17: 法律服務
      
      updateData['聯繫'] !== undefined ? formatMultiChoice(updateData['聯繫']) : existingRow[18], // 18: 聯繫
      updateData['轉介'] !== undefined ? formatMultiChoice(updateData['轉介']) : existingRow[19], // 19: 轉介
      updateData['通報'] !== undefined ? formatMultiChoice(updateData['通報']) : existingRow[20], // 20: 通報
      
      // 新欄位 (無論舊資料是否有值，都從 updateData 拿，或維持 existingRow 的空值)
      updateData['人身安全'] !== undefined ? formatMultiChoice(updateData['人身安全']) : existingRow[21], // 21: 人身安全
      updateData['會談服務'] !== undefined ? formatMultiChoice(updateData['會談服務']) : existingRow[22], // 22: 會談服務
      
      updateData['服務項目'] !== undefined ? updateData['服務項目'] : existingRow[23], // 23: 服務項目
      
      existingRow[24],                               // 24: 建立時間
      Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),  // 25: 修改時間
      existingRow[26],                               // 26: 建立者
      user.email                                     // 27: 修改者
    ];
    
    // 6. 寫入更新
    sheet.getRange(rowIndex, 1, 1, RECORD_FIELDS.length).setValues([updatedRow]);
    
    // 7. 更新時間戳
    updateLastModifiedTime('服務紀錄');
    
    // 8. 清除快取
    clearRecordsCache();
    
    console.log('✅ 更新服務紀錄成功，RecordId:', recordId);
    return { success: true, message: '更新成功' };
  });
}

/**
 * 檢查並遷移 Sheet 欄位結構
 * 如果 Sheet 只有 26 欄 (舊結構)，則插入 2 欄並更新標題，使其符合 28 欄 (新結構)
 */
function checkAndMigrateSheetColumns(sheet) {
  const lastCol = sheet.getLastColumn();
  
  if (lastCol < 28) { // 假設舊結構是 26 欄，或更少
    console.log(`偵測到舊版 Sheet 結構 (${lastCol} 欄)，開始遷移至新結構...`);
    
    // 舊結構: ... 通報(20=U), 服務項目(21=V), 建立時間(22=W) ...
    // 我們需要在 通報(U) 之後插入 2 欄 (V, W)，原本的 V 變成 X
    // 注意: insertColumnsAfter(21) 是在第 21 欄 (通報) 之後插入
    
    sheet.insertColumnsAfter(21, 2); // 在 '通報' 後面插入 '人身安全' 和 '會談服務'
    
    // 更新標題列
    sheet.getRange(1, 1, 1, RECORD_FIELDS.length).setValues([RECORD_FIELDS]);
    
    // 確保格式一致
    const headerRange = sheet.getRange(1, 1, 1, RECORD_FIELDS.length);
    headerRange.setBackground('#2c3e50');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    console.log('Sheet 欄位遷移完成。');
  }
}

// ==================== 刪除紀錄 ====================

/**
 * 刪除服務紀錄
 * @param {string} recordId - 紀錄 ID
 * @returns {Object} { success: boolean, message: string }
 */
function deleteRecord(recordId) {
  return executeWithRetry(function() {
    // 1. 權限驗證 (使用者即可刪除)
    checkUserPermission();
    
    // 2. 在所有年度 Sheet 中尋找該筆紀錄
    const found = findRecordInAllYears(recordId);
    
    if (found.rowIndex === -1) {
      throw new Error('找不到該筆紀錄');
    }
    
    const sheet = found.sheet;
    const rowIndex = found.rowIndex;
    
    // 3. 刪除該列
    sheet.deleteRow(rowIndex);
    
    // 4. 更新時間戳
    updateLastModifiedTime('服務紀錄');
    
    // 5. 清除快取
    clearRecordsCache();
    
    console.log('✅ 刪除服務紀錄成功');
    return { success: true, message: '刪除成功' };
  });
}

// ==================== 快取常數 ====================
const CACHE_CHUNK_SIZE = 200; // 每塊快取的紀錄數
const CACHE_EXPIRATION = 21600; // 6 小時 (秒)

/**
 * 查詢服務紀錄 (支援篩選和分頁，使用分塊快取)
 * @param {Object} filters - 篩選條件
 * @param {number} page - 頁碼 (1-indexed)，預設 1
 * @param {number} pageSize - 每頁筆數，預設 50
 * @returns {Object} { records: Array, totalCount: number, page: number, pageSize: number }
 */
function getRecords(filters, page, pageSize) {
  // 1. 權限驗證
  checkUserPermission();
  
  // 設定預設值
  page = page || 1;
  pageSize = pageSize || 50;
  
  // 2. 嘗試從分塊快取讀取
  const cache = CacheService.getScriptCache();
  const filterKey = JSON.stringify(filters || {});
  const version = getRecordsCacheVersion();
  const metaCacheKey = `records_meta_v${version}_${filterKey}`;
  const cachedMeta = cache.get(metaCacheKey);
  
  let allRecords;
  
  if (cachedMeta) {
    // 從分塊快取讀取
    const meta = JSON.parse(cachedMeta);
    allRecords = [];
    
    for (let i = 0; i < meta.totalChunks; i++) {
      const chunkKey = `records_chunk_v${version}_${i}_${filterKey}`;
      const chunkData = cache.get(chunkKey);
      if (chunkData) {
        allRecords = allRecords.concat(JSON.parse(chunkData));
      } else {
        // 快取不完整，需要重新讀取
        allRecords = null;
        break;
      }
    }
    
    if (allRecords) {
      console.log('使用分塊快取資料，總筆數:', allRecords.length);
    }
  }
  
  if (!allRecords) {
    // 3. 讀取 Sheet 資料
    if (!spreadsheet) initializeSpreadsheet();
    
    const targetSheets = [];
    if (filters && filters.year) {
      // 指定年份
      const sheet = spreadsheet.getSheetByName('服務紀錄_' + filters.year);
      if (sheet) targetSheets.push(sheet);
    } else {
      // 未指定年份，讀取所有年度 Sheet (排除備份)
      const allSheets = spreadsheet.getSheets();
      for (const sheet of allSheets) {
        const sheetName = sheet.getName();
        if (sheetName.startsWith('服務紀錄_')) {
          // 排除舊資料備份表
          if (sheetName.includes('備份') || sheetName.includes('舊')) continue;
          targetSheets.push(sheet);
        }
      }
    }
    
    allRecords = [];
    
    for (const sheet of targetSheets) {
      if (sheet.getLastRow() <= 1) continue;
      
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, RECORD_FIELDS.length).getDisplayValues();
      
      // 轉換為物件陣列
      const sheetRecords = data.map(row => {
        const record = {};
        RECORD_FIELDS.forEach((field, index) => {
          let value = row[index];
          record[field] = value !== null && value !== undefined ? value : '';
        });
        return record;
      });
      
      allRecords = allRecords.concat(sheetRecords);
    }
    
    // 處理資料格式 (日期、多選、布林)
    allRecords = allRecords.map(record => {
      // 轉換多選欄位為陣列
      ['陪同出庭', '會談處遇', '法律服務', '聯繫', '轉介', '通報', '人身安全', '會談服務'].forEach(field => {
        const val = record[field];
        if (val && typeof val === 'string' && val.trim() !== '') {
          record[field] = val.split(';').map(s => s.trim()).filter(Boolean);
        } else if (!Array.isArray(val)) {
          record[field] = [];
        }
      });
      // 轉換布林值
      record['有無在案'] = record['有無在案'] === '是' || record['有無在案'] === true;
      return record;
    });
    
    // 4. 套用篩選
    if (filters) {
      allRecords = applyFilters(allRecords, filters);
    }
    
    // 5. 排序 (依服務日期由新到舊)
    allRecords.sort((a, b) => {
        const dateStrA = String(a['服務日期'] || '');
        const dateStrB = String(b['服務日期'] || '');
        // YYYY-MM-DD 格式可以直接用字串比較
        if (!dateStrA) return 1;  // 無日期排最後
        if (!dateStrB) return -1;
        return dateStrB.localeCompare(dateStrA); // 從新到舊
    });
    
    // 6. 分塊快取結果
    try {
      const totalChunks = Math.ceil(allRecords.length / CACHE_CHUNK_SIZE);
      
      for (let i = 0; i < totalChunks; i++) {
        const start = i * CACHE_CHUNK_SIZE;
        const end = start + CACHE_CHUNK_SIZE;
        const chunk = allRecords.slice(start, end);
        const chunkKey = `records_chunk_v${version}_${i}_${filterKey}`;
        cache.put(chunkKey, JSON.stringify(chunk), CACHE_EXPIRATION);
      }
      
      // 儲存 metadata
      cache.put(metaCacheKey, JSON.stringify({
        totalChunks: totalChunks,
        totalCount: allRecords.length,
        version: version,
        createdAt: new Date().toISOString()
      }), CACHE_EXPIRATION);
      
      console.log('已分塊快取資料，共', totalChunks, '塊，版本:', version);
    } catch (e) {
      console.log('分塊快取儲存失敗:', e.message);
    }
  }
  
  // 7. 計算分頁
  const totalCount = allRecords.length;
  const startIndex = (page - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const pageRecords = allRecords.slice(startIndex, endIndex);
  
  console.log('查詢服務紀錄，總筆數:', totalCount, '，當前頁:', page);
  
  return {
    records: pageRecords,
    totalCount: totalCount,
    page: page,
    pageSize: pageSize
  };
}

/**
 * 清除紀錄快取 (在新增/更新/刪除後呼叫)
 * 支援分塊快取清除
 */
/**
 * 取得快取版本號
 * 用於強制失效所有相關快取
 */
function getRecordsCacheVersion() {
  const cache = CacheService.getScriptCache();
  let version = cache.get('records_cache_version');
  
  if (!version) {
    const props = PropertiesService.getScriptProperties();
    version = props.getProperty('records_cache_version') || '1';
    cache.put('records_cache_version', version, 21600);
  }
  
  return version;
}

/**
 * 遞增快取版本號
 * 當資料有變更時呼叫，這會使所有舊的 getRecords 快取失效
 */
function incrementRecordsCacheVersion() {
  const version = new Date().getTime().toString();
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();
  
  props.setProperty('records_cache_version', version);
  cache.put('records_cache_version', version, 21600);
  
  console.log('快取版本號已更新為:', version);
}

/**
 * 清除紀錄快取 (在新增/更新/刪除後呼叫)
 * 轉接至 incrementRecordsCacheVersion
 */
function clearRecordsCache() {
  incrementRecordsCacheVersion();
}

/**
 * 取得單筆紀錄
 * @param {string} recordId - 紀錄 ID
 * @returns {Object|null} 紀錄資料或 null
 */
function getRecordById(recordId) {
  checkUserPermission();
  
  const records = getRecords({ RecordId: recordId });
  return records.length > 0 ? records[0] : null;
}

// ==================== 輔助函數 ====================

/**
 * 依 RecordId 在所有年度 Sheet 中尋找紀錄
 * @param {string} recordId - 紀錄 ID
 * @returns {Object} { sheet: Sheet, rowIndex: number } 或 { sheet: null, rowIndex: -1 }
 */
function findRecordInAllYears(recordId) {
  if (!spreadsheet) initializeSpreadsheet();
  
  const searchId = String(recordId).trim();
  const allSheets = spreadsheet.getSheets();
  
  // 搜尋所有以「服務紀錄_」開頭的年度 Sheet (排除備份)
  for (const sheet of allSheets) {
    const sheetName = sheet.getName();
    if (!sheetName.startsWith('服務紀錄_')) continue;
    // 排除舊資料備份表
    if (sheetName.includes('備份') || sheetName.includes('舊')) continue;
    
    if (sheet.getLastRow() <= 1) continue;
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    
    for (let i = 0; i < data.length; i++) {
      const sheetId = String(data[i][0]).trim();
      if (sheetId === searchId) {
        console.log('findRecordInAllYears: 在', sheetName, '找到於列', i + 2);
        return { sheet: sheet, rowIndex: i + 2 };
      }
    }
  }
  
  console.log('findRecordInAllYears: 未找到 ID =', searchId);
  return { sheet: null, rowIndex: -1 };
}

/**
 * 依 RecordId 尋找紀錄的列號 (向下相容)
 * @param {string} recordId - 紀錄 ID
 * @returns {number} 列號 (1-indexed)，找不到時回傳 -1
 * @deprecated 請使用 findRecordInAllYears
 */
function findRecordRowById(recordId) {
  const result = findRecordInAllYears(recordId);
  return result.rowIndex;
}

/**
 * 格式化多選欄位 (陣列 → 分號分隔字串)
 * @param {Array|string} value - 值
 * @returns {string} 格式化後的字串
 */
function formatMultiChoice(value) {
  if (!value) return '';
  if (Array.isArray(value)) {
    return value.join('; ');
  }
  return String(value);
}

/**
 * 套用篩選條件
 * @param {Array} records - 紀錄陣列
 * @param {Object} filters - 篩選條件
 * @returns {Array} 篩選後的紀錄
 */
function applyFilters(records, filters) {
  return records.filter(record => {
    // RecordId 精確比對
    if (filters.RecordId && record.RecordId !== filters.RecordId) {
      return false;
    }
    
    // 年份篩選 (使用 getYearFromDate 確保一致性)
    if (filters.year) {
      const year = getYearFromDate(record['服務日期']);
      if (year !== filters.year) return false;
    }
    
    // 月份篩選
    if (filters.month) {
      const dateStr = String(record['服務日期']);
      let month = 0;
      if (dateStr.length >= 7) {
        month = parseInt(dateStr.substring(5, 7));
      }
      if (month !== filters.month) return false;
    }
    
    // 日期範圍 (保留相容性，但建議主要用年月)
    if (filters.startDate) {
      if (record['服務日期'] < filters.startDate) return false;
    }
    if (filters.endDate) {
      if (record['服務日期'] > filters.endDate) return false;
    }
    
    // 文字欄位模糊比對
    if (filters['個案姓名'] && !record['個案姓名'].includes(filters['個案姓名'])) {
      return false;
    }
    
    // 選擇欄位精確比對
    const exactMatchFields = ['個案來源', '個案身分', '個案國籍', '暴力類型', '個案性別', '服務社工', '服務方式', '服務主題', '服務項目'];
    for (const field of exactMatchFields) {
      if (filters[field] && record[field] !== filters[field]) {
        return false;
      }
    }
    
    // 有無在案
    if (filters['有無在案'] !== undefined && record['有無在案'] !== filters['有無在案']) {
      return false;
    }
    
    return true;
  });
}

/**
 * 驗證紀錄資料
 * @param {Object} recordData - 紀錄資料
 * @throws {Error} 如果資料不完整
 */
function validateRecordData(recordData) {
  const requiredFields = ['服務日期', '個案來源', '個案姓名', '個案身分', '個案國籍', '暴力類型', '個案性別', '服務社工', '服務方式', '服務項目'];
  
  const missingFields = requiredFields.filter(field => !recordData[field]);
  
  if (missingFields.length > 0) {
    throw new Error('缺少必填欄位：' + missingFields.join('、'));
  }
}

// ==================== 資料遷移腳本 ====================

/**
 * 將現有「服務紀錄」Sheet 的資料遷移到年度分表
 * ⚠️ 執行前請確保已備份資料！
 * 執行方式：在 Apps Script 編輯器中選擇此函數並執行
 */
function migrateRecordsByYear() {
  console.log('🚀 開始資料遷移...');
  
  if (!spreadsheet) initializeSpreadsheet();
  
  // 1. 取得原始「服務紀錄」Sheet
  const sourceSheet = spreadsheet.getSheetByName('服務紀錄');
  if (!sourceSheet) {
    console.log('❌ 找不到「服務紀錄」工作表');
    return { success: false, message: '找不到「服務紀錄」工作表' };
  }
  
  const lastRow = sourceSheet.getLastRow();
  if (lastRow <= 1) {
    console.log('⚠️ 「服務紀錄」工作表沒有資料');
    return { success: false, message: '沒有資料需要遷移' };
  }
  
  // 2. 讀取所有資料
  const data = sourceSheet.getRange(2, 1, lastRow - 1, RECORD_FIELDS.length).getValues();
  console.log('📊 共有', data.length, '筆資料需要遷移');
  
  // 3. 依年份分組
  const yearGroups = {};
  data.forEach((row, index) => {
    const serviceDate = row[1]; // 服務日期在第 2 欄 (index 1)
    const year = getYearFromDate(serviceDate);
    
    if (!yearGroups[year]) {
      yearGroups[year] = [];
    }
    yearGroups[year].push(row);
  });
  
  // 4. 寫入各年度 Sheet
  const summary = {};
  for (const year in yearGroups) {
    const yearData = yearGroups[year];
    const targetSheet = getSheetByYear(year);
    
    // 批次寫入 (效率更高)
    if (yearData.length > 0) {
      const startRow = targetSheet.getLastRow() + 1;
      targetSheet.getRange(startRow, 1, yearData.length, RECORD_FIELDS.length).setValues(yearData);
    }
    
    summary[year] = yearData.length;
    console.log('✅ 已遷移', yearData.length, '筆到「服務紀錄_' + year + '」');
  }
  
  // 5. 重新命名原始 Sheet (保留但不再使用)
  try {
    sourceSheet.setName('服務紀錄_舊資料備份');
    console.log('📝 已將原始工作表重新命名為「服務紀錄_舊資料備份」');
  } catch (e) {
    console.log('⚠️ 無法重新命名原始工作表:', e.message);
  }
  
  console.log('🎉 資料遷移完成！');
  console.log('📊 遷移摘要:', JSON.stringify(summary));
  
  return { success: true, summary: summary };
}

/**
 * 修復所有年度 Sheet 的日期格式
 * 將 M/D/YYYY 格式轉換為 YYYY-MM-DD
 */
function fixDateFormat() {
  console.log('🔧 開始修復日期格式...');
  
  if (!spreadsheet) initializeSpreadsheet();
  
  const allSheets = spreadsheet.getSheets();
  let fixedCount = 0;
  
  for (const sheet of allSheets) {
    const sheetName = sheet.getName();
    if (!sheetName.startsWith('服務紀錄_')) continue;
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;
    
    console.log('處理:', sheetName);
    
    // 讀取服務日期欄 (第 2 欄)
    const dateRange = sheet.getRange(2, 2, lastRow - 1, 1);
    const dates = dateRange.getValues();
    const newDates = [];
    
    for (let i = 0; i < dates.length; i++) {
      const val = dates[i][0];
      let formatted = val;
      
      if (val instanceof Date) {
        // Date 物件轉 YYYY-MM-DD
        formatted = Utilities.formatDate(val, 'Asia/Taipei', 'yyyy-MM-dd');
        fixedCount++;
      } else if (typeof val === 'string') {
        // 檢查是否為 M/D/YYYY 格式
        const match = val.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (match) {
          const month = match[1].padStart(2, '0');
          const day = match[2].padStart(2, '0');
          const year = match[3];
          formatted = `${year}-${month}-${day}`;
          fixedCount++;
        }
      }
      
      newDates.push([formatted]);
    }
    
    // 寫回並設定格式為純文字
    dateRange.setNumberFormat('@'); // @ 表示純文字格式
    dateRange.setValues(newDates);
  }
  
  console.log('✅ 修復完成，共修正', fixedCount, '筆日期');
  return { success: true, fixedCount: fixedCount };
}
