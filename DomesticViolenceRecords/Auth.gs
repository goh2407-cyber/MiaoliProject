// ================================================================
// Auth.gs - 權限管理系統
// ================================================================

// ==================== 使用者資訊 ====================

/**
 * 取得當前使用者資訊
 * @returns {Object} { email, role, name }
 */
function getCurrentUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    
    if (!email) {
      console.log('無法取得使用者 Email');
      return {
        email: 'unknown',
        role: 'guest',
        name: '訪客'
      };
    }
    
    const userInfo = getUserRole(email);
    return {
      email: email,
      role: userInfo.role,
      name: userInfo.name
    };
  } catch (error) {
    console.error('取得使用者資訊失敗:', error);
    return {
      email: 'unknown',
      role: 'guest',
      name: '訪客'
    };
  }
}

// ==================== 角色管理 ====================

/**
 * 取得使用者角色
 * @param {string} email - 使用者 Email
 * @returns {Object} { role, name }
 */
function getUserRole(email) {
  try {
    if (!spreadsheet) {
      initializeSpreadsheet();
    }
    
    const sheet = spreadsheet.getSheetByName('權限管理');
    if (!sheet) {
      console.log('找不到權限管理工作表');
      return { role: 'guest', name: '' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 如果權限表是空的，自動將當前使用者設為管理員
    if (data.length <= 1) {
      console.log('權限表為空，自動設定首位使用者為管理員:', email);
      try {
        sheet.appendRow([email, 'admin', '系統管理員', new Date(), '啟用']);
      } catch (e) {
        console.error('無法寫入權限表:', e);
      }
      return { role: 'admin', name: '系統管理員' };
    }
    
    // 查找使用者權限
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email && data[i][4] === '啟用') {
        return {
          role: data[i][1],
          name: data[i][2] || ''
        };
      }
    }
    
    // 找不到使用者，預設為訪客
    console.log('找不到使用者權限，設為訪客:', email);
    return { role: 'guest', name: '' };
    
  } catch (error) {
    console.error('取得使用者角色失敗:', error);
    return { role: 'guest', name: '' };
  }
}

// ==================== 權限檢查 ====================

/**
 * 檢查使用者是否為有效使用者 (admin 或 user)
 * @throws {Error} 如果權限不足
 */
function checkUserPermission() {
  const user = getCurrentUser();
  
  if (user.email === 'unknown') {
    throw new Error('無法驗證使用者身份，請重新登入');
  }
  
  if (user.role === 'guest') {
    throw new Error(`權限不足：您 (${user.email}) 尚未被授權使用此系統`);
  }
  
  return true;
}

/**
 * 檢查使用者是否為管理員
 * @throws {Error} 如果不是管理員
 */
function checkAdminPermission() {
  const user = getCurrentUser();
  
  if (user.email === 'unknown') {
    throw new Error('無法驗證使用者身份，請重新登入');
  }
  
  if (user.role !== 'admin') {
    throw new Error('權限不足：此操作需要管理員權限');
  }
  
  return true;
}

// ==================== 使用者管理 (管理員功能) ====================

/**
 * 取得所有使用者清單 (僅管理員)
 * @returns {Array} 使用者清單
 */
function getAllUsers() {
  try {
    console.log('開始執行 getAllUsers');
    checkAdminPermission();
    
    if (typeof spreadsheet === 'undefined' || !spreadsheet) {
       console.log('初始化 spreadsheet');
       initializeSpreadsheet();
    }
    
    const sheet = spreadsheet.getSheetByName('權限管理');
    if (!sheet || sheet.getLastRow() <= 1) {
      console.log('權限管理表為空');
      return [];
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    console.log('取得使用者資料筆數:', data.length);
    
    return data.map(row => ({
      email: row[0],
      role: row[1],
      name: row[2],
      createdAt: row[3] ? String(row[3]) : '',
      status: row[4]
    }));
  } catch (e) {
    console.error('getAllUsers 發生錯誤:', e);
    throw e;
  }
}

/**
 * 新增使用者 (僅管理員)
 * @param {Object} userData - { email, role, name }
 */
function addUser(userData) {
  return executeWithRetry(function() {
    checkAdminPermission();
    
    if (!userData.email || !userData.role) {
      throw new Error('Email 和角色為必填欄位');
    }
    
    if (!spreadsheet) initializeSpreadsheet();
    
    const sheet = spreadsheet.getSheetByName('權限管理');
    
    // 檢查 Email 是否已存在
    const existingUsers = getAllUsers();
    if (existingUsers.some(u => u.email === userData.email)) {
      throw new Error('此 Email 已存在');
    }
    
    sheet.appendRow([
      userData.email,
      userData.role,
      userData.name || '',
      new Date(),
      '啟用'
    ]);
    
    return { success: true, message: '使用者新增成功' };
  });
}

/**
 * 更新使用者 (僅管理員)
 * @param {string} email - 使用者 Email
 * @param {Object} updateData - { role, name, status }
 */
function updateUser(email, updateData) {
  return executeWithRetry(function() {
    checkAdminPermission();
    
    if (!spreadsheet) initializeSpreadsheet();
    
    const sheet = spreadsheet.getSheetByName('權限管理');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        if (updateData.role !== undefined) {
          sheet.getRange(i + 1, 2).setValue(updateData.role);
        }
        if (updateData.name !== undefined) {
          sheet.getRange(i + 1, 3).setValue(updateData.name);
        }
        if (updateData.status !== undefined) {
          sheet.getRange(i + 1, 5).setValue(updateData.status);
        }
        return { success: true, message: '使用者更新成功' };
      }
    }
    
    throw new Error('找不到此使用者');
  });
}
