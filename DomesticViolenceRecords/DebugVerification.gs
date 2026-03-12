/**
 * 驗證統計數據用
 * 請在編輯器上方選擇此函式，並按「執行」
 */
function verifyStatsCounts() {
  const year = 2026; // 請修改為您要檢查的年份 (西元)
  const month = '1';  // 請修改為您要檢查的月份
  
  console.log(`開始驗證 ${year} 年 ${month} 月的數據...`);
  
  const allRecords = getAllRecordsForStats(year);
  const records = filterRecordsByPeriod(allRecords, year, month);
  
  console.log(`該月份總資料筆數: ${records.length}`);
  
  // 計算各別項目的數量
  let countAccompany = 0;
  let countSafety = 0;
  
  records.forEach(r => {
    // 檢查服務項目欄位 (假設是多選，用 include 判斷)
    const items = r['服務項目'] || '';
    
    if (items.includes('陪同出庭')) {
      countAccompany++;
    }
    if (items.includes('人身安全')) {
      countSafety++;
    }
  });
  
  console.log('--- 驗證結果 ---');
  console.log(`1. 包含「陪同出庭」的紀錄數: ${countAccompany}`);
  console.log(`2. 包含「人身安全」的紀錄數: ${countSafety}`);
  console.log(`3. 兩者加總 (報表應顯示): ${countAccompany + countSafety}`);
  console.log('----------------');
  console.log('注意：如果同一筆紀錄同時包含兩者，上述簡單加總可能會重複計算項次。');
  console.log('目前的報表邏輯是計算「項次」，也就是說如果一筆紀錄選了「陪同出庭」和「人身安全」，在「服務量統計」中應該會被算成 2 次 (針對該欄位的貢獻)。');
  console.log('詳細檢查 StatsService.gs 中的 countServiceVolume 邏輯：');
  
  // 4. 詳細檢查
  console.log('--- 詳細欄位檢查 ---');
  let totalItems = 0;
  
  records.forEach(r => {
    const acc = r['陪同出庭'] || '';
    const safe = r['人身安全'] || '';
    
    if (acc || safe) {
      const accCount = acc ? acc.split(';').filter(Boolean).length : 0;
      const safeCount = safe ? safe.split(';').filter(Boolean).length : 0;
      const subtotal = accCount + safeCount;
      totalItems += subtotal;
      
      console.log(`ID: ${r.RecordId}`);
      console.log(`  陪同出庭: [${acc}] (${accCount})`);
      console.log(`  人身安全: [${safe}] (${safeCount})`);
      console.log(`  小計: ${subtotal}`);
    }
  });
  
  console.log(`--- 總計 ---`);
  console.log(`報表邏輯計算總項次: ${totalItems}`);
}
