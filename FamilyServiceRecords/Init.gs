// ================================================================
// Init.gs - 試算表初始化（家事服務紀錄系統）
// ================================================================

// ==================== 試算表初始化 ====================

/**
 * 初始化試算表
 */
function initializeSpreadsheet() {
  try {
    if (!SPREADSHEET_ID || SPREADSHEET_ID === '請填入你的試算表ID') {
      spreadsheet = SpreadsheetApp.create('家事服務紀錄系統');
      console.log('創建新試算表，ID:', spreadsheet.getId());
      console.log('請將此 ID 更新到 Code.gs 的 SPREADSHEET_ID 常數');
    } else {
      spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    
    // 確保工作表存在
    createSheetsIfNotExist();
    
    // 初始化或增量更新選項設定資料
    initializeOptionsData();
    
    return '試算表初始化成功';
  } catch (error) {
    console.error('初始化試算表失敗:', error);
    throw new Error('初始化試算表失敗: ' + error.message);
  }
}

/**
 * 創建必要的工作表
 */
function createSheetsIfNotExist() {
  const currentYear = new Date().getFullYear();
  initServiceRecordSheet(currentYear);
  initMeetingRecordSheet(currentYear);
  initMeetingExecutionSheet(currentYear);
  initParentingReportSheet();
  initParentingSurveyResponseSheet();
  initParentingPostSurveyResponseSheet();
  
  const sheetsConfig = [
    {
      name: '權限管理',
      headers: ['Email', '角色', '姓名', '建立日期', '狀態']
    },
    {
      name: '選項設定',
      headers: ['欄位名稱', '選項內容', '排序', '狀態']
    },
    {
      name: '_Metadata',
      headers: ['SheetName', 'LastModified']
    }
  ];
  
  sheetsConfig.forEach(config => {
    let sheet = spreadsheet.getSheetByName(config.name);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(config.name);
      
      // 設定表頭
      sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
      
      // 設定表頭格式
      const headerRange = sheet.getRange(1, 1, 1, config.headers.length);
      headerRange.setBackground('#1A1A1A');
      headerRange.setFontColor('#FAFAF8');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
      
      console.log('已建立工作表:', config.name);
    }
  });
  
  // 刪除預設的 Sheet1 (如果存在)
  const defaultSheet = spreadsheet.getSheetByName('工作表1');
  if (defaultSheet && spreadsheet.getSheets().length > 1) {
    spreadsheet.deleteSheet(defaultSheet);
  }
}

/**
 * 初始化服務紀錄表 (依年度)
 */
function initServiceRecordSheet(year) {
  if (!spreadsheet) initializeSpreadsheet();
  const sheetName = '服務紀錄_' + year;
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    const headers = [
      'RecordId', '服務日期', '個案來源', '個案姓名', '個案身分', '國籍別',
      '指定性別', '在案與否', '在案單位社工', '轉介單位', '轉介項目',
      '服務社工', '服務方式', '服務主題', '司法案號',
      '陪同出庭', '人身安全', '專業諮詢-法律', '專業諮詢-社福',
      '會談服務', '聯繫', '轉介', '通報',
      '服務項目', '建立時間', '修改時間', '建立者', '修改者'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1A1A1A').setFontColor('#FAFAF8').setFontWeight('bold');
    sheet.setFrozenRows(1);
    console.log('已建立工作表:', sheetName);
  } else {
    console.log('年度服務紀錄表已存在:', sheetName);
  }
}

/**
 * 初始化會面紀錄表 (依年度)
 */
function initMeetingRecordSheet(year) {
  if (!spreadsheet) initializeSpreadsheet();
  const sheetName = '會面紀錄_' + year;
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    const headers = [
      'RecordId', '服務日期', '會面分案案號', '個案姓名', '指定性別', '服務方式',
      '個案身分', '司法案號', '在案與否', '在案單位社工', '國籍別', '轉介單位',
      '服務社工', '服務主題',
      '交往交付', '專業諮詢-法律', '專業諮詢-社福', '會談服務',
      '聯繫', '轉介', '通報',
      '建立時間', '修改時間', '建立者', '修改者'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1A1A1A').setFontColor('#FAFAF8').setFontWeight('bold');
    sheet.setFrozenRows(1);
    console.log('已建立工作表:', sheetName);
  } else {
    console.log('年度會面紀錄表已存在:', sheetName);
  }
}

/**
 * 初始化會面執行紀錄表 (依年度)
 */
function initMeetingExecutionSheet(year) {
  if (!spreadsheet) initializeSpreadsheet();
  const sheetName = '會面執行_' + year;
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    const headers = [
      'RecordId', '執行日期', '會面分案案號', '執行方式', '案件類型', '執行人員',
      '未成年子女(男)', '未成年子女(女)',
      '會面方性別',
      '共同會面者(男)', '共同會面者(女)',
      '建立時間', '修改時間', '建立者', '修改者'
    ];
    sheet.appendRow(headers);
    
    // 設定標題列樣式
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#FDEADD').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    
    console.log('已建立新年度會面執行表: ' + sheetName);
  } else {
    console.log('年度會面執行表已存在: ' + sheetName);
  }
}

/**
 * 初始化初階親職總表
 */
function initParentingReportSheet() {
  if (!spreadsheet) initializeSpreadsheet();
  const sheetName = '親職_人數統計';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    const headers = [
      'RecordId', '年度', '場次日期', '講題', '聯繫人數', '報名人數', '出席人數', 
      '出席率', '聲請人(男)', '聲請人(女)', '相對人(男)', '相對人(女)', 
      '關係人(男)', '關係人(女)', '建立時間', '修改時間', '建立者', '修改者'
    ];
    sheet.appendRow(headers);
    
    // 設定標題列樣式
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#D9EAD3').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    
    console.log('已建立親職_人數統計表: ' + sheetName);
  } else {
    console.log('親職_人數統計表已存在: ' + sheetName);
  }
}


/**
 * 初始化親職問卷填答工作表（每列 = 一位填答者）
 */
function initParentingSurveyResponseSheet() {
  if (!spreadsheet) initializeSpreadsheet();
  const sheetName = '親職_問卷填答';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    const headers = [
      'RecordId', '年度', '月份', '填答序號',
      '1. 課程使我更加瞭解「調解案件的司法程序與法律知識」',
      '2. 課程讓我認識「孩子在家庭衝突中會有的身心反應」',
      '3. 課程使我知道什麼是「友善合作的父母」',
      '4. 課程使我學習到「合作照顧孩子的方式及技巧」',
      '5. 課程讓我知道遇到家庭紛爭時可尋求協助的管道',
      '建立時間', '修改時間', '建立者', '修改者'
    ];
    sheet.appendRow(headers);
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#D5A6BD').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    
    console.log('已建立親職_問卷填答表: ' + sheetName);
  } else {
    console.log('親職_問卷填答表已存在: ' + sheetName);
  }
}

/**
 * 初始化親職課後問卷填答工作表（每列 = 一位填答者）
 */
function getParentingPostSurveyHeaders_() {
  return [
    'RecordId', '年度', '月份', '填答序號',
    '1. 我會避免父母或親屬間的衝突和紛爭，讓孩子身心健康受到傷害',
    '2. 我認同要讓孩子自在的來回兩個家',
    '3. 當孩子受到家庭衝突和紛爭影響時，我會用不指責另一方的方式安撫孩子',
    '4. 我願意支持孩子和另一方父母及其他親屬有健康的親情關係',
    '5. 我不會過度追問孩子與另一方相處的情況',
    '6. 我願意善用調解制度解決我和對方的紛爭',
    '7. 我願意成為友善及合作的父母(或重要親屬)',
    '轉介_社工諮詢', '轉介_婚姻諮商', '轉介_個人心理', '轉介_家事商談', '轉介_兒少團體',
    '建立時間', '修改時間', '建立者', '修改者'
  ];
}

function ensureParentingPostSurveyHeaders_(sheet) {
  const expectedHeaders = getParentingPostSurveyHeaders_();
  const currentLastCol = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, currentLastCol).getValues()[0];
  const existingMap = {};
  currentHeaders.forEach(h => { if (h) existingMap[h] = true; });

  const missingHeaders = expectedHeaders.filter(h => !existingMap[h]);
  if (missingHeaders.length === 0) return;

  const startCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]);
  sheet.getRange(1, startCol, 1, missingHeaders.length)
    .setBackground('#FCE5CD')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, startCol, lastRow - 1, missingHeaders.length).setValue('');
  }

  console.log('親職_課後問卷填答表已補齊欄位: ' + missingHeaders.join(', '));
}

function initParentingPostSurveyResponseSheet() {
  if (!spreadsheet) initializeSpreadsheet();
  const sheetName = '親職_課後問卷填答';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    const headers = getParentingPostSurveyHeaders_();
    sheet.appendRow(headers);
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#FCE5CD').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    
    console.log('已建立親職_課後問卷填答表: ' + sheetName);
  } else {
    ensureParentingPostSurveyHeaders_(sheet);
    console.log('親職_課後問卷填答表已存在: ' + sheetName);
  }
}
function initializeOptionsData() {
  if (!spreadsheet) initializeSpreadsheet();
  
  const sheet = spreadsheet.getSheetByName('選項設定');
  
  // 預設選項資料
  const optionsData = [
    // 個案來源 (沿用家暴系統)
    ['個案來源', '自行求助', 1, '啟用'],
    ['個案來源', '自行求助-網絡諮詢', 2, '啟用'],
    ['個案來源', '接受轉介-司法系統', 3, '啟用'],
    ['個案來源', '接受轉介-社會處', 4, '啟用'],
    ['個案來源', '接受轉介-民間團體', 5, '啟用'],
    
    // 個案身分 (家事系統：當事人)
    ['個案身分', '當事人', 1, '啟用'],
    ['個案身分', '當事人-未成年', 2, '啟用'],
    ['個案身分', '相對人', 3, '啟用'],
    ['個案身分', '相對人-未成年', 4, '啟用'],
    ['個案身分', '關係人', 5, '啟用'],
    ['個案身分', '關係人-未成年', 6, '啟用'],
    ['個案身分', '民眾', 7, '啟用'],
    ['個案身分', '民眾-未成年', 8, '啟用'],
    
    // 國籍別 (沿用家暴系統)
    ['國籍別', '本國籍', 1, '啟用'],
    ['國籍別', '本國籍原住民', 2, '啟用'],
    ['國籍別', '外國籍', 3, '啟用'],
    ['國籍別', '大陸籍', 4, '啟用'],
    ['國籍別', '外國籍-印尼', 5, '啟用'],
    ['國籍別', '外國籍-越南', 6, '啟用'],
    ['國籍別', '外國籍-緬甸', 7, '啟用'],
    ['國籍別', '外國籍-菲律賓', 8, '啟用'],
    
    // 指定性別
    ['指定性別', '男', 1, '啟用'],
    ['指定性別', '女', 2, '啟用'],
    
    // 服務方式
    ['服務方式', '面談', 1, '啟用'],
    ['服務方式', '電談', 2, '啟用'],
    ['服務方式', '電談-聯繫未果', 3, '啟用'],
    ['服務方式', '書面', 4, '啟用'],
    
    // 服務社工 (動態管理)
    ['服務社工', '張雨涵', 1, '啟用'],
    ['服務社工', '范美鈴', 2, '啟用'],
    ['服務社工', '湛盈涵', 3, '啟用'],
    ['服務社工', '蔡均苹', 4, '啟用'],
    ['服務社工', '彭芮琪', 5, '啟用'],
    ['服務社工', '袁愷君', 6, '啟用'],

    
    // 服務主題 (家事系統專用)
    ['服務主題', '婚姻', 1, '啟用'],
    ['服務主題', '監護', 2, '啟用'],
    ['服務主題', '扶養', 3, '啟用'],
    ['服務主題', '會面交往', 4, '啟用'],
    ['服務主題', '收出養', 5, '啟用'],
    ['服務主題', '變更姓氏', 6, '啟用'],
    ['服務主題', '宣告', 7, '啟用'],
    ['服務主題', '繼承', 8, '啟用'],
    ['服務主題', '司法', 9, '啟用'],
    ['服務主題', '社福資訊', 10, '啟用'],
    ['服務主題', '其他', 11, '啟用'],
    
    // 服務項目
    ['服務項目', '陪同出庭', 1, '啟用'],
    ['服務項目', '諮詢服務', 2, '啟用'],
    
    // 轉介單位 (法院股別，平面單選)
    ['轉介單位', '恩', 1, '啟用'],
    ['轉介單位', '馨', 2, '啟用'],
    ['轉介單位', '真', 3, '啟用'],
    ['轉介單位', '誠', 4, '啟用'],
    ['轉介單位', '涵', 5, '啟用'],
    ['轉介單位', '菊', 6, '啟用'],
    ['轉介單位', '松', 7, '啟用'],
    ['轉介單位', '柏', 8, '啟用'],
    ['轉介單位', '家', 9, '啟用'],
    ['轉介單位', '社會處', 10, '啟用'],
    ['轉介單位', '民間單位', 11, '啟用'],
    
    // 轉介項目
    ['轉介項目', '案件訴訟協助', 1, '啟用'],
    ['轉介項目', '資源諮詢或轉介', 2, '啟用'],
    ['轉介項目', '親職教育課程', 3, '啟用'],
    ['轉介項目', '未成年人陪同出庭', 4, '啟用'],
    ['轉介項目', '成年人陪同出庭', 5, '啟用'],
    ['轉介項目', '查詢', 6, '啟用'],
    
    // ==================== 複選欄位 ====================
    
    // 陪同出庭 (多選)
    ['陪同出庭', '出庭前評估', 1, '啟用'],
    ['陪同出庭', '庭前準備', 2, '啟用'],
    ['陪同出庭', '陪同出庭', 3, '啟用'],
    ['陪同出庭', '庭後協助', 4, '啟用'],
    ['陪同出庭', '家長或監護權人諮詢協談', 5, '啟用'],
    ['陪同出庭', '其他', 6, '啟用'],
    
    // 人身安全 (多選)
    ['人身安全', '人身安全', 1, '啟用'],
    ['人身安全', '分別報到/退庭', 2, '啟用'],
    ['人身安全', '協助隔離', 3, '啟用'],
    ['人身安全', '協助安全通道', 4, '啟用'],
    ['人身安全', '陪同安全離庭', 5, '啟用'],
    ['人身安全', '叫車', 6, '啟用'],
    ['人身安全', '連結法警', 7, '啟用'],
    ['人身安全', '其他', 8, '啟用'],
    
    // 專業諮詢-法律 (多選) - 家事新欄位
    ['專業諮詢-法律', '程序說明', 1, '啟用'],
    ['專業諮詢-法律', '法律知識', 2, '啟用'],
    ['專業諮詢-法律', '書狀說明', 3, '啟用'],
    ['專業諮詢-法律', '書狀提供', 4, '啟用'],
    ['專業諮詢-法律', '證物整理', 5, '啟用'],
    ['專業諮詢-法律', '撰狀', 6, '啟用'],
    ['專業諮詢-法律', '查詢', 7, '啟用'],
    ['專業諮詢-法律', '其他', 8, '啟用'],
    
    // 專業諮詢-社福 (多選) - 家事新欄位
    ['專業諮詢-社福', '親職教育', 1, '啟用'],
    ['專業諮詢-社福', '律師諮詢', 2, '啟用'],
    ['專業諮詢-社福', '社會福利', 3, '啟用'],
    ['專業諮詢-社福', '心理諮商', 4, '啟用'],
    ['專業諮詢-社福', '收出養', 5, '啟用'],
    ['專業諮詢-社福', '親職手冊', 6, '啟用'],
    ['專業諮詢-社福', '身心障礙', 7, '啟用'],
    ['專業諮詢-社福', '就醫', 8, '啟用'],
    ['專業諮詢-社福', '居家照顧', 9, '啟用'],
    ['專業諮詢-社福', '早療', 10, '啟用'],
    ['專業諮詢-社福', '失業救濟', 11, '啟用'],
    ['專業諮詢-社福', '新住民', 12, '啟用'],
    
    // 會談服務 (多選) - 家事專用選項
    ['會談服務', '服務說明', 1, '啟用'],
    ['會談服務', '情緒支持', 2, '啟用'],
    ['會談服務', '充權', 3, '啟用'],
    ['會談服務', '未來生活討論', 4, '啟用'],
    ['會談服務', '親職認知', 5, '啟用'],
    ['會談服務', '性別概念', 6, '啟用'],
    ['會談服務', '其他認知', 7, '啟用'],
    ['會談服務', '擬定親權計畫', 8, '啟用'],
    ['會談服務', '友善父母', 9, '啟用'],
    ['會談服務', '溝通技巧', 10, '啟用'],
    ['會談服務', '物資', 11, '啟用'],
    
    // 聯繫 (多選，階層式)
    ['聯繫', '民間團體-家事商談', 1, '啟用'],
    ['聯繫', '民間團體-社區型會面', 2, '啟用'],
    ['聯繫', '民間團體-家事服務', 3, '啟用'],
    ['聯繫', '民間團體-相對人服務', 4, '啟用'],
    ['聯繫', '政府機關-戶政/地政', 5, '啟用'],
    ['聯繫', '政府機關-社福中心', 6, '啟用'],
    ['聯繫', '政府機關-家防中心', 7, '啟用'],
    ['聯繫', '法院', 8, '啟用'],
    ['聯繫', '法扶', 9, '啟用'],
    ['聯繫', '律師', 10, '啟用'],
    ['聯繫', '醫療', 11, '啟用'],
    ['聯繫', '教育', 12, '啟用'],
    ['聯繫', '就業', 13, '啟用'],
    ['聯繫', '原住民', 14, '啟用'],
    ['聯繫', '新住民', 15, '啟用'],
    ['聯繫', '其他', 16, '啟用'],
    
    // 轉介 (多選，階層式)
    ['轉介', '民間-家事商談', 1, '啟用'],
    ['轉介', '民間-社區型會面', 2, '啟用'],
    ['轉介', '民間-團體課程', 3, '啟用'],
    ['轉介', '民間-家事服務', 4, '啟用'],
    ['轉介', '民間-相對人服務', 5, '啟用'],
    ['轉介', '政府-自殺意念', 6, '啟用'],
    ['轉介', '政府-精神疾患', 7, '啟用'],
    ['轉介', '法扶', 8, '啟用'],
    ['轉介', '心理諮商', 9, '啟用'],
    ['轉介', '兒少團體', 10, '啟用'],
    ['轉介', '醫療', 11, '啟用'],
    ['轉介', '教育', 12, '啟用'],
    ['轉介', '就業', 13, '啟用'],
    ['轉介', '原住民', 14, '啟用'],
    ['轉介', '新住民', 15, '啟用'],
    ['轉介', '其他-', 16, '啟用'],
    
    // 通報 (多選)
    ['通報', '家庭暴力', 1, '啟用'],
    ['通報', '性侵害', 2, '啟用'],
    ['通報', '兒少保護', 3, '啟用'],
    ['通報', '社福中心', 4, '啟用'],
    ['通報', '自殺防治', 5, '啟用'],
    ['通報', '早療', 6, '啟用'],
    ['通報', '其他', 7, '啟用'],
    
    // ==================== 會面紀錄專用欄位 ====================
    // 會面-服務方式 (增加訪視)
    ['會面-服務方式', '面談', 1, '啟用'],
    ['會面-服務方式', '電談', 2, '啟用'],
    ['會面-服務方式', '電談-聯繫未果', 3, '啟用'],
    ['會面-服務方式', '書面', 4, '啟用'],
    ['會面-服務方式', '訪視', 5, '啟用'],
    
    // 會面-個案身分
    ['會面-個案身分', '申請方', 1, '啟用'],
    ['會面-個案身分', '照顧方', 2, '啟用'],
    ['會面-個案身分', '未成年人', 3, '啟用'],
    ['會面-個案身分', '關係人', 4, '啟用'],
    ['會面-個案身分', '網絡人員', 5, '啟用'],
    
    // 會面-服務主題
    ['會面-服務主題', '婚姻', 1, '啟用'],
    ['會面-服務主題', '監護', 2, '啟用'],
    ['會面-服務主題', '扶養', 3, '啟用'],
    ['會面-服務主題', '會面交往', 4, '啟用'],
    ['會面-服務主題', '收出養', 5, '啟用'],
    ['會面-服務主題', '履行勸告', 6, '啟用'],
    ['會面-服務主題', '變更姓氏', 7, '啟用'],
    ['會面-服務主題', '其他', 8, '啟用'],

    // 交往交付 (會面專用)
    ['交往交付', '聯繫時間', 1, '啟用'],
    ['交往交付', '說明會面注意事項', 2, '啟用'],
    ['交往交付', '說明會面進行方式', 3, '啟用'],
    ['交往交付', '了解會面期待與想法', 4, '啟用'],
    ['交往交付', '建立關係', 5, '啟用'],
    ['交往交付', '案件追蹤', 6, '啟用'],
    ['交往交付', '其他', 7, '啟用'],

    // 會面-轉介單位
    ['會面-轉介單位', '恩', 1, '啟用'],
    ['會面-轉介單位', '馨', 2, '啟用'],
    ['會面-轉介單位', '真', 3, '啟用'],
    ['會面-轉介單位', '誠', 4, '啟用'],
    ['會面-轉介單位', '涵', 5, '啟用'],
    ['會面-轉介單位', '菊', 6, '啟用'],
    ['會面-轉介單位', '松', 7, '啟用'],
    ['會面-轉介單位', '柏', 8, '啟用'],
    ['會面-轉介單位', '家', 9, '啟用'],
    ['會面-轉介單位', '縣府', 10, '啟用'],
    ['會面-轉介單位', '民間單位', 11, '啟用'],
    
    // 會面-會談服務
    ['會面-會談服務', '服務說明', 1, '啟用'],
    ['會面-會談服務', '情緒支持', 2, '啟用'],
    ['會面-會談服務', '充權', 3, '啟用'],
    ['會面-會談服務', '未來生活討論', 4, '啟用'],
    ['會面-會談服務', '親職認知', 5, '啟用'],
    ['會面-會談服務', '性別概念', 6, '啟用'],
    ['會面-會談服務', '其他認知', 7, '啟用'],
    ['會面-會談服務', '擬定親權計畫', 8, '啟用'],
    ['會面-會談服務', '友善父母', 9, '啟用'],
    ['會面-會談服務', '溝通技巧', 10, '啟用'],
    ['會面-會談服務', '物資提供', 11, '啟用'],
    ['會面-會談服務', '其他', 12, '啟用'],

    // 會面-專業諮詢-社福
    ['會面-專業諮詢-社福', '社會福利', 1, '啟用'],
    ['會面-專業諮詢-社福', '身心障礙', 2, '啟用'],
    ['會面-專業諮詢-社福', '就醫', 3, '啟用'],
    ['會面-專業諮詢-社福', '居家照顧', 4, '啟用'],
    ['會面-專業諮詢-社福', '早療', 5, '啟用'],
    ['會面-專業諮詢-社福', '兒少收出養', 6, '啟用'],
    ['會面-專業諮詢-社福', '失業救濟', 7, '啟用'],
    ['會面-專業諮詢-社福', '新住民', 8, '啟用'],
    ['會面-專業諮詢-社福', '親職教育', 9, '啟用'],
    ['會面-專業諮詢-社福', '律師諮詢', 10, '啟用'],
    ['會面-專業諮詢-社福', '心理諮商', 11, '啟用'],
    ['會面-專業諮詢-社福', '親職手冊說明', 12, '啟用'],

    // ==================== 會面執行專用欄位 ====================
    ['執行方式', '監督會面', 1, '啟用'],
    ['執行方式', '監督交付', 2, '啟用'],
    ['執行方式', '監督會面+交付', 3, '啟用'],
    ['執行方式', '視訊會面', 4, '啟用'],

    ['案件類型', '婚姻', 1, '啟用'],
    ['案件類型', '監護', 2, '啟用'],
    ['案件類型', '扶養', 3, '啟用'],
    ['案件類型', '會面交往', 4, '啟用'],
    ['案件類型', '收出養', 5, '啟用'],
    ['案件類型', '履行勸告', 6, '啟用'],
    ['案件類型', '變更姓氏', 7, '啟用'],
    ['案件類型', '其他', 8, '啟用']
  ];
  
  if (optionsData.length > 0) {
    if (sheet.getLastRow() > 1) {
      // 增量更新：比對並打入缺少的選項
      const existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      const existingSet = new Set(existingData.map(r => r[0] + '|' + r[1]));
      
      const missingOptions = optionsData.filter(opt => !existingSet.has(opt[0] + '|' + opt[1]));
      
      if (missingOptions.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, missingOptions.length, 4).setValues(missingOptions);
        console.log('已增量更新選項設定資料，共新增', missingOptions.length, '筆');
      } else {
        console.log('選項設定皆已存在，無需新增');
      }
    } else {
      // 第一次初始化
      sheet.getRange(2, 1, optionsData.length, 4).setValues(optionsData);
      console.log('已完成全新選項設定資料初始化，共', optionsData.length, '筆');
    }
  }
}

/**
 * 強制重新初始化選項資料（會清除現有資料）
 */
function forceReinitializeOptions() {
  if (!spreadsheet) initializeSpreadsheet();
  
  const sheet = spreadsheet.getSheetByName('選項設定');
  
  // 清除現有資料（保留表頭）
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    console.log('已清除現有選項資料');
  }
  
  // 重新初始化
  initializeOptionsData();
  
  console.log('✅ 強制重新初始化完成！');
  return '選項資料已重新初始化';
}
