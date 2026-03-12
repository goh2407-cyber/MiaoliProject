// ================================================================
// ParentingService.gs - 初階親職報表管理
// ================================================================

/**
 * 取得指定年度的親職報表講題
 * 如果該年度無資料，自動初始化 12 個月份的空白紀錄並回傳
 * @param {string} year 年度 (例如: 115)
 * @returns {Object} { status: string, message: string, data: Array }
 */
function getParentingTopics(year) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('親職_人數統計');
    if (!sheet) {
      throw new Error('找不到工作表「親職_人數統計」');
    }
    
    const lastRow = sheet.getLastRow();
    // 預期回傳的資料結構
    let topicsData = [];
    
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
      const data = dataRange.getValues();
      
      // 尋找符合該年度的資料
      const yearStr = String(year);
      // 假設第二欄 (index 1) 是年度，第三欄 (index 2) 是月份
      const yearRows = [];
      const yearRowIndices = []; // 紀錄在 sheet 中的實際行數 (1-based)，方便之後更新
      
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][1]) === yearStr) {
          yearRows.push(data[i]);
          yearRowIndices.push(i + 2); // 資料從第 2 行開始，所以 index + 2
        }
      }
      
      if (yearRows.length > 0) {
        // 該年度已有資料，將資料轉成物件陣列回傳
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const recordIdIdx = headers.indexOf('RecordId');
        const dateIdx = headers.indexOf('場次日期');
        const topicIdx = headers.indexOf('講題');
        const contactIdx = headers.indexOf('聯繫人數');
        const signIdx = headers.indexOf('報名人數');
        const attendIdx = headers.indexOf('出席人數');
        const rateIdx = headers.indexOf('出席率');
        const appMaleIdx = headers.indexOf('聲請人(男)');
        const appFemIdx = headers.indexOf('聲請人(女)');
        const resMaleIdx = headers.indexOf('相對人(男)');
        const resFemIdx = headers.indexOf('相對人(女)');
        const relMaleIdx = headers.indexOf('關係人(男)');
        const relFemIdx = headers.indexOf('關係人(女)');
        
        topicsData = yearRows.map((row, index) => {
          let dateVal = row[dateIdx];
          if (dateVal && dateVal instanceof Date) {
            // 轉換為 YYYY-MM-DD
            const y = dateVal.getFullYear();
            const m = String(dateVal.getMonth() + 1).padStart(2, '0');
            const d = String(dateVal.getDate()).padStart(2, '0');
            dateVal = `${y}-${m}-${d}`;
          }

          return {
            rowId: yearRowIndices[index], // sheet row number for updating
            recordId: row[recordIdIdx],
            month: index + 1, // 直接用陣列順序作為月份 (1~12)
            date: dateVal || '',
            topic: row[topicIdx] || '',
            contactCount: row[contactIdx] || 0,
            signCount: row[signIdx] || 0,
            attendCount: row[attendIdx] || 0,
            attendRate: row[rateIdx] || '',
            applicantMale: row[appMaleIdx] || 0,
            applicantFemale: row[appFemIdx] || 0,
            respondentMale: row[resMaleIdx] || 0,
            respondentFemale: row[resFemIdx] || 0,
            relatedMale: row[relMaleIdx] || 0,
            relatedFemale: row[relFemIdx] || 0
          };
        });
        
        // 檢查問卷狀態
        const surveySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('親職_問卷填答');
        let surveyRecordIds = new Set();
        if (surveySheet && surveySheet.getLastRow() > 1) {
          const surveyData = surveySheet.getRange(2, 1, surveySheet.getLastRow() - 1, 1).getValues();
          surveyData.forEach(r => surveyRecordIds.add(String(r[0])));
        }

        topicsData.forEach(t => {
          t.hasSurvey = t.recordId ? surveyRecordIds.has(String(t.recordId)) : false;
        });

        return {
          status: 'success',
          message: '讀取成功',
          data: topicsData
        };
      }
    }
    
    // 如果走到這裡，代表該年度還沒有任何資料
    return {
      status: 'not_found',
      message: '該年度尚無資料，請點擊「建立表格」新增',
      data: []
    };
    
  } catch (error) {
    console.error('getParentingTopics 錯誤：', error);
    return {
      status: 'error',
      message: '讀取失敗：' + error.message,
      data: null
    };
  }
}

/**
 * 手動建立指定年度的親職報表全新表格 (供前端呼叫)
 * @param {string} year 年度 (例如: 115)
 * @returns {Object} { status: string, message: string }
 */
function createParentingYearData(year) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('親職_人數統計');
    if (!sheet) {
      throw new Error('找不到工作表「親職_人數統計」');
    }
    
    // 檢查是否真的沒有資料，避免重複建立
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
      const data = dataRange.getValues();
      const yearStr = String(year);
      
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][1]) === yearStr) {
           return {
             status: 'error',
             message: '該年度已經存在資料，請勿重複建立'
           };
        }
      }
    }

    // 呼叫原本的初始化邏輯
    const result = initializeYearTopics(year, sheet);
    return {
      status: 'success',
      message: '已成功建立 ' + year + ' 年度表格',
      data: result.data // 若前端需要直接繪製，可一併回傳
    };
    
  } catch (error) {
    console.error('createParentingYearData 錯誤：', error);
    return {
      status: 'error',
      message: '建立失敗：' + error.message
    };
  }
}

/**
 * 初始化指定年度 1~12 月的空白紀錄
 * @param {string} year 年度
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet '親職_人數統計' 工作表
 */
function initializeYearTopics(year, sheet) {
  const currentUser = Session.getActiveUser().getEmail() || 'System';
  const now = new Date();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRows = [];
  const startRow = sheet.getLastRow() + 1;
  const topicsData = [];
  
  const defaultTopics = [
    '別讓孩子捲入父母的衝突',
    '讓孩子知道我愛他',
    '如何做友善的父母'
  ];
  
  for (let month = 1; month <= 12; month++) {
    const recordId = Utilities.getUuid();
    // 準備要插入的新行，長度與 headers 相同，皆為空白字串
    const rowData = new Array(headers.length).fill('');
    
    // 依序循環帶入預設講題
    const topicText = defaultTopics[(month - 1) % 3];
    
    rowData[headers.indexOf('RecordId')] = recordId;
    rowData[headers.indexOf('年度')] = year;
    rowData[headers.indexOf('建立時間')] = now;
    rowData[headers.indexOf('建立者')] = currentUser;
    rowData[headers.indexOf('講題')] = topicText;
    
    newRows.push(rowData);
    
    topicsData.push({
      rowId: startRow + month - 1,
      recordId: recordId,
      month: month,
      date: '',
      topic: topicText,
      contactCount: 0,
      signCount: 0,
      attendCount: 0,
      attendRate: '',
      applicantMale: 0,
      applicantFemale: 0,
      respondentMale: 0,
      respondentFemale: 0,
      relatedMale: 0,
      relatedFemale: 0
    });
  }
  
  // 一次性寫入 12 行
  sheet.getRange(startRow, 1, 12, headers.length).setValues(newRows);
  
  return {
    status: 'success',
    message: '已初始化該年度資料',
    data: topicsData,
    isNew: true
  };
}

/**
 * 儲存指定年度的講題更新
 * @param {Object} formData 包含 year 與 topics (陣列)
 */
function saveParentingTopics(formData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('親職_人數統計');
    if (!sheet) {
      throw new Error('找不到工作表「親職_人數統計」');
    }
    
    const { year, topics } = formData;
    const currentUser = Session.getActiveUser().getEmail() || 'System';
    const now = new Date();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 建立 header index map 方便查找
    const hMap = {};
    headers.forEach((h, idx) => hMap[h] = idx + 1); // 1-based index for getRange

    // 讀取當前年度的行物件，並更新
    for (const data of topics) {
      if (data.rowId) {
        // 直接更新那一行的所有數值
        // 將 date 字串轉回 Date 物件
        let dateObj = '';
        if (data.date) {
            dateObj = new Date(data.date);
        }

        sheet.getRange(data.rowId, hMap['場次日期']).setValue(dateObj);
        sheet.getRange(data.rowId, hMap['講題']).setValue(data.topic);
        sheet.getRange(data.rowId, hMap['聯繫人數']).setValue(data.contactCount);
        sheet.getRange(data.rowId, hMap['報名人數']).setValue(data.signCount);
        sheet.getRange(data.rowId, hMap['出席人數']).setValue(data.attendCount);
        sheet.getRange(data.rowId, hMap['出席率']).setValue(data.attendRate);
        sheet.getRange(data.rowId, hMap['聲請人(男)']).setValue(data.applicantMale);
        sheet.getRange(data.rowId, hMap['聲請人(女)']).setValue(data.applicantFemale);
        sheet.getRange(data.rowId, hMap['相對人(男)']).setValue(data.respondentMale);
        sheet.getRange(data.rowId, hMap['相對人(女)']).setValue(data.respondentFemale);
        sheet.getRange(data.rowId, hMap['關係人(男)']).setValue(data.relatedMale);
        sheet.getRange(data.rowId, hMap['關係人(女)']).setValue(data.relatedFemale);

        sheet.getRange(data.rowId, hMap['修改時間']).setValue(now);
        sheet.getRange(data.rowId, hMap['修改者']).setValue(currentUser);
      }
    }
    
    return {
      status: 'success',
      message: '儲存完成'
    };
    
  } catch (error) {
    console.error('saveParentingTopics 錯誤：', error);
    return {
      status: 'error',
      message: '儲存失敗：' + error.message
    };
  }
}
/**
 * 取得指定 RecordId 的所有問卷填答列（新版：每列 = 一位填答者）
 * @param {string} recordId 
 * @returns {Object} { status, data: [ {seq, Q1..Q5, svc_*}, ... ] }
 */
const PARENTING_SURVEY_QUESTION_HEADERS = [
  '1. 課程使我更加瞭解「調解案件的司法程序與法律知識」',
  '2. 課程讓我認識「孩子在家庭衝突中會有的身心反應」',
  '3. 課程使我知道什麼是「友善合作的父母」',
  '4. 課程使我學習到「合作照顧孩子的方式及技巧」',
  '5. 課程讓我知道遇到家庭紛爭時可尋求協助的管道'
];

function getParentingSurveyHeadersForService_() {
  return [
    'RecordId', '年度', '月份', '填答序號'
  ]
  .concat(PARENTING_SURVEY_QUESTION_HEADERS)
  .concat(['建立時間', '修改時間', '建立者', '修改者']);
}

function ensureParentingSurveySheetHeaders_(sheet) {
  const expectedHeaders = getParentingSurveyHeadersForService_();
  const currentLastCol = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, currentLastCol).getValues()[0];
  const existingMap = {};
  currentHeaders.forEach(h => { if (h) existingMap[h] = true; });

  const missingHeaders = expectedHeaders.filter(h => !existingMap[h]);
  if (missingHeaders.length === 0) return;

  const startCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]);
  sheet.getRange(1, startCol, 1, missingHeaders.length)
    .setBackground('#D5A6BD')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, startCol, lastRow - 1, missingHeaders.length).setValue('');
  }
}

function getParentingSurvey(recordId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('親職_問卷填答');
    if (!sheet) return { status: 'success', data: [] };
    ensureParentingSurveySheetHeaders_(sheet);
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { status: 'success', data: [] };
    
    const headers = data[0];
    const ridIdx = headers.indexOf('RecordId');
    
    const responses = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][ridIdx] === recordId) {
        const row = {};
        headers.forEach((h, idx) => {
          let v = data[i][idx];
          if (v instanceof Date) v = v.toISOString();
          row[h] = v;
        });

        for (let q = 0; q < 5; q++) {
          const questionHeader = PARENTING_SURVEY_QUESTION_HEADERS[q];
          const v = row[questionHeader];
          row['Q' + (q + 1)] = (v !== '' && v != null) ? v : 0;
        }

        responses.push(row);
      }
    }
    
    // 依填答序號排序
    responses.sort((a, b) => (a['填答序號'] || 0) - (b['填答序號'] || 0));
    
    return { status: 'success', data: responses };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

/**
 * 儲存問卷填答資料（整批覆寫）
 * @param {Object} payload { recordId, year, month, responses: [{Q1..Q5, svc_*}, ...] }
 */
function saveParentingSurvey(payload) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('親職_問卷填答');
    if (!sheet) {
      initParentingSurveyResponseSheet();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('親職_問卷填答');
    }
    ensureParentingSurveySheetHeaders_(sheet);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const ridIdx = headers.indexOf('RecordId');
    const recordId = payload.recordId;
    
    const currentUser = Session.getActiveUser().getEmail() || 'System';
    const now = new Date();
    
    // 1. 刪除該 RecordId 的所有舊列（從底部刪起避免 index 偏移）
    const rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][ridIdx] === recordId) {
        rowsToDelete.push(i + 1); // Sheet row is 1-based
      }
    }
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    
    // 2. 寫入新的填答列
    const responses = payload.responses || [];
    if (responses.length === 0) {
      return { status: 'success', message: '問卷填答已清空' };
    }
    
    const hMap = {};
    headers.forEach((h, idx) => hMap[h] = idx);
    
    const newRows = responses.map((resp, idx) => {
      const row = new Array(headers.length).fill('');
      const setCol = (name, value) => {
        const cIdx = hMap[name];
        if (cIdx !== undefined && cIdx >= 0) row[cIdx] = value;
      };
      setCol('RecordId', recordId);
      setCol('年度', payload.year);
      setCol('月份', payload.month);
      setCol('填答序號', idx + 1);
      for (let q = 0; q < 5; q++) {
        const score = resp['Q' + (q + 1)] || 0;
        setCol(PARENTING_SURVEY_QUESTION_HEADERS[q], score);
      }
      setCol('建立時間', now);
      setCol('修改時間', now);
      setCol('建立者', currentUser);
      setCol('修改者', currentUser);
      return row;
    });
    
    // 批次寫入效能優化
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, headers.length).setValues(newRows);
    
    return { status: 'success', message: '已儲存 ' + newRows.length + ' 筆問卷填答' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

/**
 * 取得指定 RecordId 的所有課後問卷填答列
 * @param {string} recordId 
 */
const PARENTING_POST_SURVEY_QUESTION_HEADERS = [
  '1. 我會避免父母或親屬間的衝突和紛爭，讓孩子身心健康受到傷害',
  '2. 我認同要讓孩子自在的來回兩個家',
  '3. 當孩子受到家庭衝突和紛爭影響時，我會用不指責另一方的方式安撫孩子',
  '4. 我願意支持孩子和另一方父母及其他親屬有健康的親情關係',
  '5. 我不會過度追問孩子與另一方相處的情況',
  '6. 我願意善用調解制度解決我和對方的紛爭',
  '7. 我願意成為友善及合作的父母(或重要親屬)'
];

const PARENTING_POST_SURVEY_LEGACY_Q_HEADERS = ['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7'];
const PARENTING_POST_SURVEY_DEPRECATED_SOCIAL_DETAIL_HEADERS = [
  '社工諮詢_司法程序諮詢',
  '社工諮詢_書狀諮詢',
  '社工諮詢_未成年人陪同出庭',
  '社工諮詢_親職計畫討論',
  '社工諮詢_會面計畫討論'
];

function getParentingPostSurveyHeadersForService_() {
  return [
    'RecordId', '年度', '月份', '填答序號'
  ]
  .concat(PARENTING_POST_SURVEY_QUESTION_HEADERS)
  .concat(['轉介_社工諮詢', '轉介_婚姻諮商', '轉介_個人心理', '轉介_家事商談', '轉介_兒少團體', '讓子女參加心理團體'])
  .concat(['建立時間', '修改時間', '建立者', '修改者']);
}

function ensureParentingPostSurveySheetHeaders_(sheet) {
  const expectedHeaders = getParentingPostSurveyHeadersForService_();
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
}

function getParentingPostSurvey(recordId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('親職_課後問卷填答');
    if (!sheet) return { status: 'success', data: [] };
    ensureParentingPostSurveySheetHeaders_(sheet);
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { status: 'success', data: [] };
    
    const headers = data[0];
    const ridIdx = headers.indexOf('RecordId');
    
    const responses = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][ridIdx] === recordId) {
        const row = {};
        headers.forEach((h, idx) => {
          let v = data[i][idx];
          if (v instanceof Date) v = v.toISOString();
          row[h] = v;
        });

        for (let q = 0; q < 7; q++) {
          const questionHeader = PARENTING_POST_SURVEY_QUESTION_HEADERS[q];
          const legacyHeader = PARENTING_POST_SURVEY_LEGACY_Q_HEADERS[q];
          const v = row[questionHeader];
          row['Q' + (q + 1)] = (v !== '' && v != null) ? v : (row[legacyHeader] || 0);
        }

        let socialCount = parseInt(row['轉介_社工諮詢'], 10);
        if (isNaN(socialCount) || socialCount < 0) socialCount = 0;
        if (socialCount > 5) socialCount = 5;
        if (socialCount === 0) {
          socialCount = PARENTING_POST_SURVEY_DEPRECATED_SOCIAL_DETAIL_HEADERS.reduce((sum, key) => {
            const raw = row[key];
            return sum + ((raw === 1 || raw === '1' || raw === true) ? 1 : 0);
          }, 0);
        }
        row['轉介_社工諮詢'] = socialCount > 0 ? socialCount : '';

        responses.push(row);
      }
    }
    
    // 依填答序號排序
    responses.sort((a, b) => (a['填答序號'] || 0) - (b['填答序號'] || 0));
    
    return { status: 'success', data: responses };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

/**
 * 儲存課後問卷填答資料（整批覆寫）
 * @param {Object} payload { recordId, year, month, responses: [{Q1..Q7, 轉介_*}, ...] }
 */
function saveParentingPostSurvey(payload) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('親職_課後問卷填答');
    if (!sheet) {
      initParentingPostSurveyResponseSheet();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('親職_課後問卷填答');
    }
    ensureParentingPostSurveySheetHeaders_(sheet);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const ridIdx = headers.indexOf('RecordId');
    const recordId = payload.recordId;
    
    const currentUser = Session.getActiveUser().getEmail() || 'System';
    const now = new Date();
    
    // 1. 刪除舊列
    const rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][ridIdx] === recordId) {
        rowsToDelete.push(i + 1);
      }
    }
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    
    // 2. 寫入新列
    const responses = payload.responses || [];
    if (responses.length === 0) {
      return { status: 'success', message: '課後問卷填答已清空' };
    }

    // 2-1. 驗證資料：Q1~Q7 必填且僅能 1~5；社工圈選數僅能 0~5（可空白）
    for (let i = 0; i < responses.length; i++) {
      const resp = responses[i] || {};
      for (let q = 1; q <= 7; q++) {
        const score = parseInt(resp['Q' + q], 10);
        if (isNaN(score) || score < 1 || score > 5) {
          return { status: 'error', message: `第 ${i + 1} 列 Q${q} 必須為 1~5（且不可空白）` };
        }
      }

      const rawSocial = resp['轉介_社工諮詢'];
      if (!(rawSocial === '' || rawSocial == null)) {
        const socialCount = parseInt(rawSocial, 10);
        if (isNaN(socialCount) || socialCount < 0 || socialCount > 5) {
          return { status: 'error', message: `第 ${i + 1} 列社工圈選數必須為 0~5` };
        }
      }

      const rawChildManual = resp['讓子女參加心理團體'];
      if (!(rawChildManual === '' || rawChildManual == null)) {
        const childManualCount = parseInt(rawChildManual, 10);
        if (isNaN(childManualCount) || childManualCount < 0 || childManualCount > 999) {
          return { status: 'error', message: `讓子女參加心理團體必須為 0~999` };
        }
      }
    }
    
    const hMap = {};
    headers.forEach((h, idx) => hMap[h] = idx);
    
    const newRows = responses.map((resp, idx) => {
      const row = new Array(headers.length).fill('');
      const setCol = (name, value) => {
        const cIdx = hMap[name];
        if (cIdx !== undefined && cIdx >= 0) row[cIdx] = value;
      };

      setCol('RecordId', recordId);
      setCol('年度', payload.year);
      setCol('月份', payload.month);
      setCol('填答序號', idx + 1);

      for (let q = 0; q < 7; q++) {
        const score = parseInt(resp['Q' + (q + 1)], 10);
        setCol(PARENTING_POST_SURVEY_QUESTION_HEADERS[q], score);
        setCol(PARENTING_POST_SURVEY_LEGACY_Q_HEADERS[q], score);
      }

      setCol('轉介_婚姻諮商', resp['轉介_婚姻諮商'] || '');
      setCol('轉介_個人心理', resp['轉介_個人心理'] || '');
      setCol('轉介_家事商談', resp['轉介_家事商談'] || '');
      setCol('轉介_兒少團體', resp['轉介_兒少團體'] || '');
      const childManualRaw = resp['讓子女參加心理團體'];
      const childManualCount = (childManualRaw === '' || childManualRaw == null)
        ? ''
        : parseInt(childManualRaw, 10);
      setCol('讓子女參加心理團體', childManualCount);

      const socialCountRaw = resp['轉介_社工諮詢'];
      const socialCount = (socialCountRaw === '' || socialCountRaw == null)
        ? ''
        : parseInt(socialCountRaw, 10);
      setCol('轉介_社工諮詢', socialCount);

      PARENTING_POST_SURVEY_DEPRECATED_SOCIAL_DETAIL_HEADERS.forEach(h => setCol(h, ''));

      setCol('建立時間', now);
      setCol('修改時間', now);
      setCol('建立者', currentUser);
      setCol('修改者', currentUser);
      return row;
    });
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, headers.length).setValues(newRows);
    
    return { status: 'success', message: '已儲存 ' + newRows.length + ' 筆課後問卷填答' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

/**
 * 取得指定 RecordId 的問卷統計結果（自動計算）
 * @param {string} recordId
 * @returns {Object} { status, data: { totalResponses, questions: [{q, counts, percents}], services } }
 */
function getParentingSurveyStats(recordId) {
  try {
    const result = getParentingSurvey(recordId);
    if (result.status !== 'success') return result;
    
    const responses = result.data;
    const total = responses.length;

    const services = {
      '婚姻諮商': 0,
      '個人心理': 0,
      '家事商談': 0,
      '兒少團體': 0,
      '社工諮詢': 0
    };
    const postResult = getParentingPostSurvey(recordId);
    if (postResult && postResult.status === 'success') {
      const postResponses = postResult.data || [];
      postResponses.forEach(r => {
        if (parseInt(r['轉介_婚姻諮商'], 10) > 0) services['婚姻諮商']++;
        // 需求：轉介個人心理 = 個人心理勾選 + 兒少團體勾選
        let personalCount = 0;
        if (parseInt(r['轉介_個人心理'], 10) > 0) personalCount += 1;
        if (parseInt(r['轉介_兒少團體'], 10) > 0) personalCount += 1;
        services['個人心理'] += personalCount;
        if (parseInt(r['轉介_家事商談'], 10) > 0) services['家事商談']++;
        let socialCount = parseInt(r['轉介_社工諮詢'], 10);
        if (!isNaN(socialCount) && socialCount > 0) {
          if (socialCount > 5) socialCount = 5;
          services['社工諮詢'] += socialCount;
        }
      });
      // 讓子女參加心理團體統計：採手寫欄位（每月同值，取第一個有效值）
      const manualChildStat = postResponses.reduce((acc, r) => {
        if (acc > 0) return acc;
        const n = parseInt(r['讓子女參加心理團體'], 10);
        return (!isNaN(n) && n > 0) ? n : 0;
      }, 0);
      services['兒少團體'] = manualChildStat;
    }

    if (total === 0) {
      return { status: 'success', data: { totalResponses: 0, questions: [], services: services } };
    }

    const questions = [];
    for (let q = 1; q <= 5; q++) {
      const key = 'Q' + q;
      const counts = [0, 0, 0, 0, 0, 0]; // index 0=未填答, 1=非常不同意, ..., 5=非常同意

      responses.forEach(r => {
        const score = parseInt(r[key]) || 0;
        if (score >= 0 && score <= 5) counts[score]++;
      });

      questions.push({
        q: q,
        counts: {
          '非常同意': counts[5],
          '同意': counts[4],
          '普通': counts[3],
          '不同意': counts[2],
          '非常不同意': counts[1],
          '未填答': counts[0]
        },
        percents: {
          '非常同意': (counts[5] / total * 100).toFixed(2),
          '同意': (counts[4] / total * 100).toFixed(2),
          '普通': (counts[3] / total * 100).toFixed(2),
          '不同意': (counts[2] / total * 100).toFixed(2),
          '非常不同意': (counts[1] / total * 100).toFixed(2),
          '未填答': (counts[0] / total * 100).toFixed(2)
        }
      });
    }

    return {
      status: 'success',
      data: {
        totalResponses: total,
        questions: questions,
        services: services
      }
    };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}
