// ================================================================
// ExportService.gs - 匯出功能 (Excel / PDF)
// ================================================================

/**
 * 匯出紀錄為 Excel 格式
 * @param {Object} filters - 篩選條件
 * @returns {Object} { url: string, filename: string }
 */
function exportToExcel(filters) {
  checkUserPermission();
  
  // 取得篩選後的紀錄
  const records = getRecords(filters);
  
  if (records.length === 0) {
    throw new Error('沒有符合條件的紀錄可匯出');
  }
  
  // 建立新的試算表
  const filename = '家暴服務紀錄_' + formatDateForExport(new Date());
  const exportSpreadsheet = SpreadsheetApp.create(filename);
  const sheet = exportSpreadsheet.getActiveSheet();
  
  // 定義匯出欄位 (排除系統欄位)
  const exportFields = [
    '服務日期', '個案來源', '個案姓名', '個案身分', '個案國籍',
    '暴力類型', '個案性別', '有無在案', '在案社工', '接受轉介項目', '轉案股別/單位',
    '服務社工', '服務方式', '服務主題', '陪同出庭', '會談處遇', '法律服務',
    '聯繫', '轉介', '通報', '服務項目', '建立時間', '建立者'
  ];
  
  // 寫入表頭
  sheet.getRange(1, 1, 1, exportFields.length).setValues([exportFields]);
  
  // 設定表頭樣式
  const headerRange = sheet.getRange(1, 1, 1, exportFields.length);
  headerRange.setBackground('#2c3e50');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 寫入資料
  const exportData = records.map(record => {
    return exportFields.map(field => {
      let value = record[field];
      
      // 處理多選欄位
      if (Array.isArray(value)) {
        value = value.join('; ');
      }
      
      // 處理布林值
      if (typeof value === 'boolean') {
        value = value ? '是' : '否';
      }
      
      // 處理日期
      if (value instanceof Date) {
        value = formatDateForExport(value);
      }
      
      return value || '';
    });
  });
  
  if (exportData.length > 0) {
    sheet.getRange(2, 1, exportData.length, exportFields.length).setValues(exportData);
  }
  
  // 自動調整欄寬
  for (let i = 1; i <= exportFields.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // 取得 URL
  const url = exportSpreadsheet.getUrl();
  
  // 強制寫入資料，確保產生 PDF 時有內容
  SpreadsheetApp.flush();
  
  console.log('✅ 匯出 Excel 成功，檔案:', filename);
  
  return {
    success: true,
    url: url,
    filename: filename,
    recordCount: records.length
  };
}

/**
 * 匯出紀錄為 PDF 格式
 * @param {Object} filters - 篩選條件
 * @returns {Object} { url: string, filename: string }
 */
function exportToPdf(filters) {
  checkUserPermission();
  
  // 先產生 Excel
  const excelResult = exportToExcel(filters);
  
  // 取得試算表 ID
  const spreadsheetId = excelResult.url.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheetId = spreadsheet.getActiveSheet().getSheetId();
  
  // 產生 PDF URL
  const pdfUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + 
                 '/export?format=pdf' +
                 '&size=A4' +
                 '&portrait=false' +  // 橫向
                 '&fitw=true' +       // 符合寬度
                 '&gridlines=true' +  // 顯示格線
                 '&gid=' + sheetId;
  
  // 取得 PDF Blob
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(pdfUrl, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  
  const pdfBlob = response.getBlob().setName(excelResult.filename + '.pdf');
  
  // 儲存到 Google Drive
  const file = DriveApp.createFile(pdfBlob);
  const downloadUrl = file.getUrl();
  
  console.log('✅ 匯出 PDF 成功，檔案:', excelResult.filename + '.pdf');
  
  return {
    success: true,
    url: downloadUrl,
    filename: excelResult.filename + '.pdf',
    recordCount: excelResult.recordCount
  };
}

/**
 * 格式化日期用於匯出
 * @param {Date} date - 日期
 * @returns {string} 格式化後的日期字串
 */
function formatDateForExport(date) {
  if (!date) return '';
  return Utilities.formatDate(date, 'Asia/Taipei', 'yyyy-MM-dd');
}
