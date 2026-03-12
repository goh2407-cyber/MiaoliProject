// ================================================================
// RepairService.gs - 系統維護與修復功能
// ================================================================

/**
 * 補填缺漏的 RecordId
 * 掃描所有年度 Sheet，找出有「服務日期」但沒有「RecordId」的資料，並不回填 UUID
 */
function generateMissingIds() {
  console.log('🔧 開始執行補填缺漏 ID...');
  
  if (!spreadsheet) initializeSpreadsheet();
  
  const allSheets = spreadsheet.getSheets();
  let totalFixed = 0;
  let summary = [];
  
  for (const sheet of allSheets) {
    const sheetName = sheet.getName();
    
    // 只處理「服務紀錄_」開頭的 Sheet
    if (!sheetName.startsWith('服務紀錄_')) continue;
    // 排除備份
    if (sheetName.includes('備份') || sheetName.includes('舊')) continue;
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue; // 只有標題或空的
    
    //讀取前兩欄：RecordId (A), 服務日期 (B)
    // getRange(row, col, numRows, numCols)
    const range = sheet.getRange(2, 1, lastRow - 1, 2);
    const data = range.getValues();
    let sheetFixedCount = 0;
    
    for (let i = 0; i < data.length; i++) {
      const recordId = data[i][0];
      const serviceDate = data[i][1];
      
      // 條件：RecordId 為空 且 服務日期不為空
      if (!recordId && serviceDate) {
        const newId = Utilities.getUuid();
        
        // 寫回 RecordId (Row 索引是 i + 2，因為 data 是從第 2 列開始)
        // Col 1 是 RecordId
        sheet.getRange(i + 2, 1).setValue(newId);
        
        sheetFixedCount++;
      }
    }
    
    if (sheetFixedCount > 0) {
      console.log(`✅ ${sheetName}: 已補填 ${sheetFixedCount} 筆 ID`);
      summary.push(`${sheetName}: ${sheetFixedCount} 筆`);
      totalFixed += sheetFixedCount;
    }
  }
  
  if (totalFixed > 0) {
    const msg = `🎉 執行完成！共補填 ${totalFixed} 筆 ID。\n詳細：\n${summary.join('\n')}`;
    console.log(msg);
    return msg;
  } else {
    console.log('✅ 檢查完成，沒有發現缺漏 ID 的資料。');
    return '檢查完成，所有資料都有 ID，無需補填。';
  }
}
