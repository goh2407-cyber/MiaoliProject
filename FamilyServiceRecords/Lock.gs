// ================================================================
// Lock.gs - 鎖定與重試機制
// ================================================================

// ==================== 重試機制 ====================
/**
 * 執行操作並自動重試（防止多使用者衝突）
 * @param {Function} operation - 要執行的操作
 * @param {number} maxRetries - 最大重試次數
 * @param {number} waitMs - 重試間隔（毫秒）
 */
function executeWithRetry(operation, maxRetries, waitMs) {
  maxRetries = maxRetries || 3;
  waitMs = waitMs || 2000;
  
  const lock = LockService.getScriptLock();
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const acquired = lock.tryLock(10000); // 10秒逾時
      
      if (!acquired) {
        if (attempt === maxRetries) {
          throw new Error('系統繁忙，請稍後再試（已重試 ' + maxRetries + ' 次）');
        }
        console.log('🔄 第 ' + attempt + ' 次嘗試失敗，等待重試...');
        Utilities.sleep(waitMs);
        continue;
      }
      
      try {
        const result = operation();
        return result;
      } finally {
        lock.releaseLock();
      }
      
    } catch (error) {
      if (attempt === maxRetries) {
        throw error;
      }
      console.log('⚠️ 操作失敗，準備重試: ' + error.message);
      Utilities.sleep(waitMs);
    }
  }
}

// ==================== 時間戳管理 ====================
/**
 * 取得指定工作表的最後修改時間
 * @param {string} sheetName - 工作表名稱
 * @returns {number} 時間戳（毫秒）
 */
function getLastModifiedTime(sheetName) {
  try {
    if (!spreadsheet) initializeSpreadsheet();
    const metaSheet = getOrCreateMetaSheet();
    const data = metaSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sheetName) {
        return data[i][1] ? new Date(data[i][1]).getTime() : 0;
      }
    }
    return 0;
  } catch (e) {
    console.error('取得時間戳失敗:', e);
    return 0;
  }
}

/**
 * 更新指定工作表的最後修改時間
 * @param {string} sheetName - 工作表名稱
 */
function updateLastModifiedTime(sheetName) {
  try {
    if (!spreadsheet) initializeSpreadsheet();
    const metaSheet = getOrCreateMetaSheet();
    const data = metaSheet.getDataRange().getValues();
    const now = new Date();
    
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sheetName) {
        metaSheet.getRange(i + 1, 2).setValue(now);
        found = true;
        break;
      }
    }
    
    if (!found) {
      metaSheet.appendRow([sheetName, now]);
    }
  } catch (e) {
    console.error('更新時間戳失敗:', e);
  }
}

/**
 * 取得或建立用於記錄時間戳的元數據工作表
 * @returns {Sheet} 元數據工作表
 */
function getOrCreateMetaSheet() {
  let sheet = spreadsheet.getSheetByName('_Metadata');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('_Metadata');
    sheet.getRange(1, 1, 1, 2).setValues([['SheetName', 'LastModified']]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 2)
         .setFontWeight('bold')
         .setBackground('#f3f3f3');
  }
  return sheet;
}

/**
 * 檢查客戶端資料是否過期
 * @param {string} sheetName - 工作表名稱
 * @param {number} clientTimestamp - 客戶端的時間戳
 * @returns {Object} 檢查結果
 */
function checkDataFreshness(sheetName, clientTimestamp) {
  const serverTimestamp = getLastModifiedTime(sheetName);
  
  return {
    isStale: serverTimestamp > clientTimestamp,
    serverTimestamp: serverTimestamp,
    needsRefresh: serverTimestamp > clientTimestamp
  };
}
