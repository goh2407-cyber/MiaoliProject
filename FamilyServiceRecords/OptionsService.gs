// ================================================================
// OptionsService.gs - 選項設定管理（家事服務紀錄系統）
// ================================================================

/**
 * 取得指定欄位的選項清單
 * @param {string} fieldName - 欄位名稱
 * @returns {Array} 選項清單
 */
function getOptions(fieldName) {
  try {
    if (!spreadsheet) initializeSpreadsheet();
    
    const sheet = spreadsheet.getSheetByName('選項設定');
    if (!sheet || sheet.getLastRow() <= 1) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    
    return data
      .filter(row => row[0] === fieldName && row[3] === '啟用')
      .sort((a, b) => a[2] - b[2])  // 依排序欄位排序
      .map(row => row[1]);          // 只回傳選項內容
      
  } catch (error) {
    console.error('取得選項失敗:', error);
    return [];
  }
}

/**
 * 取得所有欄位的選項 (用於前端初始化)
 * @returns {Object} { fieldName: [options] }
 */
/**
 * 取得所有欄位的選項 (用於前端初始化)
 * @returns {Object} { fieldName: [options] }
 */
function getAllOptions() {
  try {
    checkUserPermission();
    
    if (!spreadsheet) initializeSpreadsheet();
    
    const sheet = spreadsheet.getSheetByName('選項設定');
    if (!sheet || sheet.getLastRow() <= 1) {
      return {};
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    
    const options = {};
    
    data.forEach(row => {
      const fieldName = row[0];
      const optionValue = row[1];
      const sortOrder = row[2];
      const status = row[3];
      
      if (status !== '啟用') return;
      
      if (!options[fieldName]) {
        options[fieldName] = [];
      }
      
      options[fieldName].push({
        value: optionValue,
        sort: sortOrder
      });
    });
    
    // 對每個欄位的選項排序 (除了服務社工，因為之後會被覆蓋)
    Object.keys(options).forEach(field => {
      options[field] = options[field]
        .sort((a, b) => a.sort - b.sort)
        .map(item => item.value);
    });
    
    // 特殊處理：合併「權限管理」的使用者到「服務社工」選項中
    // 這樣可以保留「選項設定」中的靜態名單，又能動態加入新使用者
    const userSheet = spreadsheet.getSheetByName('權限管理');
    if (userSheet && userSheet.getLastRow() > 1) {
      // 讀取 Email, 角色, 姓名, 建立日期, 狀態
      const userData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 5).getValues();
      const socialWorkers = userData
        .filter(row => row[2] && row[4] !== '停用') // 有姓名且狀態不是停用
        .map(row => row[2]); // 取出姓名
      
      if (socialWorkers.length > 0) {
        // 取得原本已存在的選項（來自選項設定 Sheet）
        const existingOptions = options['服務社工'] || [];
        
        // 合併並去重複
        // 取得不需要顯示的社工 (在選項設定中被設為「停用」的)
        const disabledOptions = data
          .filter(row => row[0] === '服務社工' && row[3] !== '啟用')
          .map(row => row[1]);

        // 合併並去重複，同時過濾掉被停用的
        const merged = [...new Set([...existingOptions, ...socialWorkers])]
          .filter(name => !disabledOptions.includes(name));
          
        options['服務社工'] = merged.sort(); // 簡單按筆畫/字母排序
      }
    }
    
    return options;
    
  } catch (error) {
    console.error('取得所有選項失敗:', error);
    return {};
  }
}

/**
 * 取得選項設定完整清單 (包含停用選項、排序等) - 用於管理介面
 */
function getOptionsSettings() {
  return executeWithRetry(function() {
    checkAdminPermission();
    if (!spreadsheet) initializeSpreadsheet();
    const sheet = spreadsheet.getSheetByName('選項設定');
    if (!sheet || sheet.getLastRow() <= 1) return {};
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    const options = {};
    
    data.forEach(row => {
      const field = row[0];
      if (!options[field]) options[field] = [];
      options[field].push({
        value: row[1],
        sort: row[2],
        status: row[3]
      });
    });
    
    // Sort by sort order
    Object.keys(options).forEach(k => {
      options[k].sort((a,b) => a.sort - b.sort);
    });
    
    return options;
  });
}

// ==================== 選項管理 (管理員功能) ====================

/**
 * 新增選項 (僅管理員)
 * @param {string} fieldName - 欄位名稱
 * @param {string} optionValue - 選項內容
 * @param {number} sortOrder - 排序 (可選)
 */
function addOption(fieldName, optionValue, sortOrder) {
  return executeWithRetry(function() {
    checkAdminPermission();
    
    if (!fieldName || !optionValue) {
      throw new Error('欄位名稱和選項內容為必填');
    }
    
    if (!spreadsheet) initializeSpreadsheet();
    
    const sheet = spreadsheet.getSheetByName('選項設定');
    
    // 檢查是否已存在
    const existingOptions = getOptions(fieldName);
    if (existingOptions.includes(optionValue)) {
      throw new Error('此選項已存在');
    }
    
    // 計算排序值
    if (!sortOrder) {
      const allData = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 4).getValues();
      const fieldOptions = allData.filter(row => row[0] === fieldName);
      sortOrder = fieldOptions.length + 1;
    }
    
    sheet.appendRow([fieldName, optionValue, sortOrder, '啟用']);
    
    return { success: true, message: '選項新增成功' };
  });
}

/**
 * 更新選項 (僅管理員)
 * @param {string} fieldName - 欄位名稱
 * @param {string} oldValue - 舊選項內容
 * @param {Object} updateData - { value, sort, status }
 */
function updateOption(fieldName, oldValue, updateData) {
  return executeWithRetry(function() {
    checkAdminPermission();
    
    if (!spreadsheet) initializeSpreadsheet();
    
    const sheet = spreadsheet.getSheetByName('選項設定');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === fieldName && data[i][1] === oldValue) {
        if (updateData.value !== undefined) {
          sheet.getRange(i + 1, 2).setValue(updateData.value);
        }
        if (updateData.sort !== undefined) {
          sheet.getRange(i + 1, 3).setValue(updateData.sort);
        }
        if (updateData.status !== undefined) {
          sheet.getRange(i + 1, 4).setValue(updateData.status);
        }
        return { success: true, message: '選項更新成功' };
      }
    }
    
    throw new Error('找不到此選項');
  });
}

/**
 * 停用選項 (僅管理員)
 * @param {string} fieldName - 欄位名稱
 * @param {string} optionValue - 選項內容
 */
function disableOption(fieldName, optionValue) {
  return updateOption(fieldName, optionValue, { status: '停用' });
}



// ==================== 服務社工管理 (特殊功能) ====================

/**
 * 取得服務社工設定 (包含系統使用者與選項設定)
 */
function getSocialWorkerSettings() {
  return executeWithRetry(function() {
    checkAdminPermission();
    if (!spreadsheet) initializeSpreadsheet();

    // 1. 取得系統使用者 (權限管理)
    const userSheet = spreadsheet.getSheetByName('權限管理');
    const systemUsers = [];
    if (userSheet && userSheet.getLastRow() > 1) {
      const userData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 5).getValues();
      userData
        .filter(row => row[2] && row[4] !== '停用') // 有姓名且啟用
        .forEach(row => {
          if (!systemUsers.includes(row[2])) {
            systemUsers.push(row[2]);
          }
        });
    }

    // 2. 取得選項設定中的服務社工
    const optionSheet = spreadsheet.getSheetByName('選項設定');
    const manualOptions = [];
    const disabledMap = new Set(); // 記錄被停用的名字

    if (optionSheet && optionSheet.getLastRow() > 1) {
      const optionData = optionSheet.getRange(2, 1, optionSheet.getLastRow() - 1, 4).getValues();
      optionData.forEach(row => {
        if (row[0] === '服務社工') {
          const name = row[1];
          const status = row[3];
          
          // 記錄所有在選項設定裡的項目
          manualOptions.push({
            name: name,
            status: status
          });

          // 如果是停用，加入黑名單
          if (status !== '啟用') {
            disabledMap.add(name);
          }
        }
      });
    }

    // 3. 合併清單
    const allWorkers = new Map();

    // 先加入系統使用者
    systemUsers.forEach(name => {
      allWorkers.set(name, {
        name: name,
        source: 'system', // 系統使用者
        isVisible: !disabledMap.has(name) // 如果沒被停用就是顯示
      });
    });

    // 再加入手動設定的 (可能會覆蓋 source，或新增純手動的)
    manualOptions.forEach(opt => {
      if (allWorkers.has(opt.name)) {
        // 已經存在 (是系統使用者)，確認是否被手動停用
        const worker = allWorkers.get(opt.name);
        worker.source = 'system_managed'; // 系統使用者 + 手動管理
        worker.isVisible = (opt.status === '啟用');
      } else {
        // 純手動新增的
        allWorkers.set(opt.name, {
          name: opt.name,
          source: 'manual',
          isVisible: (opt.status === '啟用')
        });
      }
    });

    // 轉為陣列並排序
    const result = Array.from(allWorkers.values()).sort((a, b) => {
      return a.name.localeCompare(b.name, 'zh-Hant');
    });

    return result;
  });
}

/**
 * 儲存服務社工設定
 * @param {Array<{name: string, isVisible: boolean, isNew: boolean}>} settings
 */
function saveSocialWorkerSettings(settings) {
  return executeWithRetry(function() {
    checkAdminPermission();
    if (!spreadsheet) initializeSpreadsheet();

    const sheet = spreadsheet.getSheetByName('選項設定');
    if (!sheet) throw new Error('找不到選項設定工作表');

    const lastRow = sheet.getLastRow();
    const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 4).getValues() : [];
    
    // 建立現有選項的對照表 Map<name, rowIndex>
    const existingRows = new Map();
    data.forEach((row, index) => {
      if (row[0] === '服務社工') {
        existingRows.set(row[1], index + 2); // 儲存實際列號 (1-based)
      }
    });

    settings.forEach(setting => {
      const name = setting.name;
      const shouldBeVisible = setting.isVisible;
      const rowNum = existingRows.get(name);

      if (rowNum) {
        // 已存在於選項設定中 -> 更新狀態
        const currentStatus = sheet.getRange(rowNum, 4).getValue();
        const newStatus = shouldBeVisible ? '啟用' : '停用';
        
        if (currentStatus !== newStatus) {
           sheet.getRange(rowNum, 4).setValue(newStatus);
        }
      } else {
        // 不存在於選項設定中
        // 1. 如果是隱藏 (shouldBeVisible=false)，必須新增停用記錄以覆蓋系統預設
        // 2. 如果是標記為新 (isNew=true)，必須新增
        if (!shouldBeVisible || setting.isNew) {
           sheet.appendRow(['服務社工', name, 999, shouldBeVisible ? '啟用' : '停用']);
        }
      }
    });

    return { success: true, message: '社工名單設定已更新' };
  });
}

/**
 * 刪除服務社工 (僅限手動新增的)
 * @param {string} name - 社工姓名
 */
function deleteSocialWorker(name) {
  return executeWithRetry(function() {
    checkAdminPermission();
    if (!name) throw new Error('未提供姓名');

    if (!spreadsheet) initializeSpreadsheet();
    const sheet = spreadsheet.getSheetByName('選項設定');
    if (!sheet) throw new Error('找不到選項設定工作表');

    const data = sheet.getDataRange().getValues();
    
    // 尋找並刪除
    // 從後面往前找，避免索引偏移 (雖然理論上應該只有一筆)
    let deletedCount = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      // 檢查是否為「服務社工」且姓名相符
      if (data[i][0] === '服務社工' && data[i][1] === name) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }

    if (deletedCount === 0) {
      throw new Error('找不到該社工的設定資料');
    }

    return { success: true, message: '已刪除社工: ' + name };
  });
}
