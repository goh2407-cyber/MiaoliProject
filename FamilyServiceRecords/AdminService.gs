// ================================================================
// AdminService.gs - 系統管理功能 (使用者管理)
// ================================================================

/**
 * 取得所有使用者
 * @returns {Array} 使用者清單
 */
function getAllUsers() {
  return executeWithRetry(function() {
    checkAdminPermission();
    
    if (!spreadsheet) initializeSpreadsheet();
    
    const sheet = spreadsheet.getSheetByName('權限管理');
    if (!sheet || sheet.getLastRow() <= 1) {
      return [];
    }
    
    // 讀取 Email, 角色, 姓名, 建立日期, 狀態
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    
    return data.map(row => ({
      email: row[0],
      role: row[1],
      name: row[2],
      createdDate: row[3] instanceof Date ? Utilities.formatDate(row[3], 'Asia/Taipei', 'yyyy-MM-dd') : row[3],
      status: row[4]
    }));
  });
}

/**
 * 新增使用者
 * @param {Object} userData - { email, role, name }
 */
function addUser(userData) {
  return executeWithRetry(function() {
    checkAdminPermission();
    
    const { email, role, name } = userData;
    
    if (!email || !role || !name) {
      throw new Error('Email、姓名與角色為必填');
    }
    
    if (!spreadsheet) initializeSpreadsheet();
    const sheet = spreadsheet.getSheetByName('權限管理');
    
    // 檢查 Email 是否重複
    const existingData = sheet.getDataRange().getValues();
    const exists = existingData.slice(1).some(row => row[0] === email);
    
    if (exists) {
      throw new Error('此 Email 已存在');
    }
    
    const now = new Date();
    sheet.appendRow([email, role, name, now, '啟用']);
    
    return { success: true, message: '使用者新增成功' };
  });
}

/**
 * 更新使用者
 * @param {string} email - 目標 Email
 * @param {Object} updateData - { role, name, status }
 */
function updateUser(email, updateData) {
  return executeWithRetry(function() {
    checkAdminPermission();
    
    if (!spreadsheet) initializeSpreadsheet();
    const sheet = spreadsheet.getSheetByName('權限管理');
    
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex((row, index) => index > 0 && row[0] === email);
    
    if (rowIndex === -1) {
      throw new Error('找不到該使用者');
    }
    
    const realRowIndex = rowIndex + 1; // 1-based index
    
    if (updateData.role) sheet.getRange(realRowIndex, 2).setValue(updateData.role);
    if (updateData.name) sheet.getRange(realRowIndex, 3).setValue(updateData.name);
    if (updateData.status) sheet.getRange(realRowIndex, 5).setValue(updateData.status);
    
    return { success: true, message: '使用者更新成功' };
  });
}

/**
 * 刪除使用者 (標記為停用或物理刪除?)
 * 建議物理刪除或保留但狀態改為停用。這裡實作物理刪除。
 */
function deleteUser(email) {
  return executeWithRetry(function() {
    checkAdminPermission();
    
    if (email === Session.getActiveUser().getEmail()) {
        throw new Error('無法刪除自己');
    }

    if (!spreadsheet) initializeSpreadsheet();
    const sheet = spreadsheet.getSheetByName('權限管理');
    
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex((row, index) => index > 0 && row[0] === email);
    
    if (rowIndex === -1) {
      throw new Error('找不到該使用者');
    }
    
    sheet.deleteRow(rowIndex + 1);
    
    return { success: true, message: '使用者刪除成功' };
  });
}
