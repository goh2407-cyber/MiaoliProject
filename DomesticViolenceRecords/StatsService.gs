// ================================================================
// StatsService.gs - 統計報表服務
// ================================================================

/**
 * 根據年份取得對應的服務欄位名稱
 * 2025 以前使用舊欄位，2026 起使用新欄位
 * @param {number} year - 西元年份
 * @returns {Array} 服務欄位名稱陣列
 */
function getServiceFields(year) {
  if (year >= 2026) {
    // 2026 新結構：人身安全從會談處遇獨立出來
    // 但為了相容舊資料或誤填到舊欄位的資料，保留 '會談處遇'
    return ['陪同出庭', '人身安全', '會談服務', '會談處遇', '法律服務', '聯繫', '轉介', '通報'];
  }
  // 2025 以前舊結構
  return ['陪同出庭', '法律服務', '會談處遇', '通報', '轉介', '聯繫'];
}

/**
 * 根據年份取得會談欄位名稱
 * @param {number} year - 西元年份
 * @returns {string} 欄位名稱
 */
function getMeetingFieldName(year) {
  return year >= 2026 ? '會談服務' : '會談處遇';
}

/**
 * 產生統計報表 HTML
 * @param {number} year - 西元年份
 * @param {string} month - 月份 (1-12 或 '全年')
 * @returns {Object} { html: string, totalCount: number }
 */
function generateStatisticsReport(year, month) {
  checkUserPermission();
  
  // 取得該年度紀錄 (不分頁)
  const allRecords = getAllRecordsForStats(year);
  
  // 篩選指定期間的資料
  const records = filterRecordsByPeriod(allRecords, year, month);
  
  // 計算各種統計數據
  const stats = {
    totalCount: records.length,
    rocYear: year - 1911,
    monthText: month === '全年' ? '全年' : month + '月',
    
    // 個案統計
    consultCases: countDistinctCases(records, '諮詢服務'),
    accompCases: countDistinctCases(records, '陪同出庭'),
    totalDistinctCases: countTotalDistinctCases(records),
    
    // 用於標題的總人次與項次
    totalPersonTimes: records.length,
    totalItemTimes: 0, // 稍後計算
    
    // 服務主題統計
    serviceTopic: countByServiceTopic(records),
    
    // 服務項目統計 (多選欄位)
    serviceItems: countServiceItems(records),
    
    // 4. 服務方式
    serviceMethod: countServiceMethodCrossTab(records, year),
    
    // 5. 案件來源統計
    caseSource: countCaseSourceCrossTab(records),
    
    // 非撤回案件統計
    nonWithdrawal: countByIdentityGender(records, false),
    
    // 撤回案件統計
    withdrawal: countByIdentityGender(records, true),
    
    // 服務量統計表
    serviceVolume: countServiceVolume(records),
    
    // 服務案件類型統計 (第五張表)
    serviceTypeStats: countServiceTypes(records)
  };
  
  // 計算總項次 (基於服務方式統計的總計)
  stats.totalItemTimes = stats.serviceMethod.總計.小計;
  
  // 產生 HTML
  const html = generateReportHtml(stats, year, month);
  
  return {
    html: html,
    totalCount: stats.totalCount,
    stats: stats
  };
}

/**
 * 取得所有紀錄（用於統計，不分頁）
 * @param {number} year - 指定年份
 */
function getAllRecordsForStats(year) {
  if (!spreadsheet) initializeSpreadsheet();
  
  // 使用 RecordService 的 getSheetByYear 函數
  const sheet = getSheetByYear(year);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  
  const headers = data[0];
  const records = [];
  
  for (let i = 1; i < data.length; i++) {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = data[i][index];
    });
    records.push(record);
  }
  
  return records;
}

/**
 * 依期間篩選資料
 */
function filterRecordsByPeriod(records, year, month) {
  return records.filter(record => {
    const dateValue = record['服務日期'];
    if (!dateValue) return false;
    
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else {
      date = new Date(dateValue);
    }
    
    if (isNaN(date.getTime())) return false;
    
    const recordYear = date.getFullYear();
    const recordMonth = date.getMonth() + 1;
    
    if (recordYear !== year) return false;
    if (month !== '全年' && recordMonth !== parseInt(month)) return false;
    
    return true;
  });
}

/**
 * 計算不重複個案數
 */
function countDistinctCases(records, serviceItem) {
  const filtered = records.filter(r => r['服務項目'] === serviceItem);
  const names = new Set(filtered.map(r => r['個案姓名']).filter(Boolean));
  return names.size;
}

/**
 * 計算總不重複個案數 (不分服務項目)
 */
function countTotalDistinctCases(records) {
  const names = new Set(records.map(r => r['個案姓名']).filter(Boolean));
  return names.size;
}

/**
 * 服務主題統計
 */
function countByServiceTopic(records) {
  const topics = ['暫時', '延長', '變更', '撤銷', '撤回', '追加', '通常', '違反', '抗告', '其他'];
  const counts = {};
  
  topics.forEach(topic => {
    if (topic === '其他') {
      counts[topic] = records.filter(r => {
        const value = r['服務主題'] || '';
        return value.includes('其他');
      }).length;
    } else {
      counts[topic] = records.filter(r => r['服務主題'] === topic).length;
    }
  });
  
  counts['總計'] = records.filter(r => r['服務主題']).length;
  
  return counts;
}

/**
 * 計算多選欄位項目數
 */
function countMultiSelectItems(records, fieldName, items) {
  let total = 0;
  
  items.forEach(item => {
    records.forEach(record => {
      const value = record[fieldName] || '';
      const values = typeof value === 'string' ? value.split(';').map(s => s.trim()) : [];
      if (values.includes(item)) {
        total++;
      }
    });
  });
  
  return total;
}

/**
 * 計算單筆紀錄的多選欄位項目數
 */
function countRecordMultiSelectItems(record, fieldName) {
  const value = record[fieldName] || '';
  if (!value) return 0;
  const items = typeof value === 'string' ? value.split(';').filter(Boolean) : [];
  return items.length;
}

/**
 * 服務項目統計
 */
function countServiceItems(records) {
  const 陪同出庭Options = ["出庭前評估", "庭前準備", "庭後協助", "陪同出庭", "其他"];
  const 會談處遇Options = ["人身安全", "充權", "其他", "性別概念", "情緒支持", "社福資訊", "澄清迷思", "初層心諮"];
  const 法律服務Options = ["其他", "法律知識", "查詢", "書狀說明", "程序說明", "撰狀", "聲請保護令", "提供律師"];
  const 聯繫Options = ["民間-經扶", "民團-目睹服務", "民團-相對人服務", "民團-家事商談", "民團-家暴後追單位", "民團-家處", "民團-監督會面", "其他", "法扶", "法院", "律師", "政府-公所", "政府-心衛相關單位", "政府-戶政/地政", "政府-地檢署", "政府-身障", "政府-社福中心", "政府-家防中心", "政府-警政", "原住民", "教育", "就業", "新住民", "醫療"];
  const 轉介Options = ["心理諮商", "民間-社區型會面", "民團-相對人服務", "民團-家事商談", "團體課程", "兒少團體", "其他單位", "其他服務處", "法扶", "政府-自殺意念", "政府-精神疾患", "原住民", "教育", "就業", "新住民", "醫療"];
  const 通報Options = ["早期療育", "自殺防治", "兒少保護", "性侵害", "社福中心", "家庭暴力", "精障個案", "藥酒癮"];
  
  return {
    陪同出庭: countMultiSelectItems(records, '陪同出庭', 陪同出庭Options),
    會談處遇: countMultiSelectItems(records, '會談處遇', 會談處遇Options),
    法律服務: countMultiSelectItems(records, '法律服務', 法律服務Options),
    聯繫: countMultiSelectItems(records, '聯繫', 聯繫Options),
    轉介: countMultiSelectItems(records, '轉介', 轉介Options),
    通報: countMultiSelectItems(records, '通報', 通報Options)
  };
}

/**
 * 服務方式交叉分析
 */
function countServiceMethodCrossTab(records, year) {
  const result = {
    諮詢服務: { 電談: 0, 聯繫未果: 0, 面談: 0, 書面: 0, 小計: 0 },
    陪同出庭: { 電談: 0, 聯繫未果: 0, 面談: 0, 書面: 0, 小計: 0 },
    總計: { 電談: 0, 聯繫未果: 0, 面談: 0, 書面: 0, 小計: 0 }
  };
  
  // 取得該年度的服務欄位列表
  const fields = getServiceFields(year);

  records.forEach(record => {
    const 服務項目 = record['服務項目'];
    const 服務方式 = record['服務方式'];
    
    if (!服務項目) return;
    
    // 計算該紀錄的服務項目數 (動態根據年度欄位加總)
    let serviceCount = 0;
    fields.forEach(field => {
      serviceCount += countRecordMultiSelectItems(record, field);
    });
    
    if (服務項目 === '諮詢服務') {
      if (服務方式 === '電談') {
        result.諮詢服務.電談 += serviceCount;
      } else if (服務方式 === '電談-聯繫未果') {
        result.諮詢服務.聯繫未果 += serviceCount;
      } else if (服務方式 === '面談') {
        result.諮詢服務.面談 += serviceCount;
      } else if (服務方式 === '書面') {
        result.諮詢服務.書面 += serviceCount;
      }
      result.諮詢服務.小計 += serviceCount;
    } else if (服務項目 === '陪同出庭') {
      if (服務方式 === '電談') {
        result.陪同出庭.電談 += serviceCount;
      } else if (服務方式 === '電談-聯繫未果') {
        result.陪同出庭.聯繫未果 += serviceCount;
      } else if (服務方式 === '面談') {
        result.陪同出庭.面談 += serviceCount;
      } else if (服務方式 === '書面') {
        result.陪同出庭.書面 += serviceCount;
      }
      result.陪同出庭.小計 += serviceCount;
    }
  });
  
  // 計算總計
  result.總計.電談 = result.諮詢服務.電談 + result.陪同出庭.電談;
  result.總計.聯繫未果 = result.諮詢服務.聯繫未果 + result.陪同出庭.聯繫未果;
  result.總計.面談 = result.諮詢服務.面談 + result.陪同出庭.面談;
  result.總計.書面 = result.諮詢服務.書面 + result.陪同出庭.書面;
  result.總計.小計 = result.諮詢服務.小計 + result.陪同出庭.小計;
  
  return result;
}

/**
 * 個案來源交叉分析
 */
function countCaseSourceCrossTab(records) {
  const sources = ['自行求助', '自行求助-網絡諮詢', '接受轉介-司法系統', '接受轉介-社會處', '接受轉介-民間團體'];
  const items = ['諮詢服務', '陪同出庭'];
  
  const result = {};
  
  items.forEach(item => {
    result[item] = {};
    sources.forEach(source => {
      result[item][source] = records.filter(r => 
        r['服務項目'] === item && r['個案來源'] === source
      ).length;
    });
    result[item]['小計'] = records.filter(r => r['服務項目'] === item).length;
  });
  
  result['總計'] = {};
  sources.forEach(source => {
    result['總計'][source] = records.filter(r => r['個案來源'] === source).length;
  });
  result['總計']['小計'] = records.length;
  
  return result;
}

/**
 * 依身分/性別統計
 */
function countByIdentityGender(records, isWithdrawal) {
  const identities = ['聲請人', '相對人', '關係人', '民眾'];
  const genders = ['女', '男'];
  
  const filtered = records.filter(r => {
    const topic = r['服務主題'];
    return isWithdrawal ? topic === '撤回' : topic !== '撤回';
  });
  
  const result = {};
  
  identities.forEach(identity => {
    result[identity] = {};
    genders.forEach(gender => {
      const subset = filtered.filter(r => {
        const 身分 = r['個案身分'] || '';
        return 身分.includes(identity) && r['個案性別'] === gender;
      });
      
      result[identity][gender] = {
        服務人數: subset.length,
        陪同出庭: sumMultiSelectCounts(subset, '陪同出庭'),
        人身安全: sumMultiSelectCounts(subset, '人身安全'), // 新增
        法律服務: sumMultiSelectCounts(subset, '法律服務'),
        會談處遇: sumMultiSelectCounts(subset, '會談處遇'), // 2025以前
        會談服務: sumMultiSelectCounts(subset, '會談服務'), // 2026以後
        通報: sumMultiSelectCounts(subset, '通報'),
        轉介: sumMultiSelectCounts(subset, '轉介'),
        聯繫: sumMultiSelectCounts(subset, '聯繫')
      };
      
      // 計算該分類的項次總計
      // 注意：2026年起 '會談服務' 取代 '會談處遇'，但為了相容，我們兩個都加，反正沒值的那個會是 0
      result[identity][gender]['總計'] = 
        result[identity][gender].陪同出庭 +
        result[identity][gender].人身安全 +
        result[identity][gender].法律服務 +
        result[identity][gender].會談處遇 +
        result[identity][gender].會談服務 +
        result[identity][gender].通報 +
        result[identity][gender].轉介 +
        result[identity][gender].聯繫;
    });
  });
  
  // 計算總計 (依性別加總所有身分)
  result['總計'] = { 女: { 服務人數: 0 }, 男: { 服務人數: 0 } };
  genders.forEach(gender => {
    let total = { 
      服務人數: 0, 
      陪同出庭: 0, 
      人身安全: 0, 
      法律服務: 0, 
      會談處遇: 0, 
      會談服務: 0, 
      通報: 0, 
      轉介: 0, 
      聯繫: 0, 
      總計: 0 
    };
    
    identities.forEach(identity => {
      Object.keys(total).forEach(key => {
        total[key] += result[identity][gender][key] || 0;
      });
    });
    result['總計'][gender] = total;
  });
  
  return result;
}

/**
 * 計算多選欄位項目數總和
 */
function sumMultiSelectCounts(records, fieldName) {
  return records.reduce((sum, record) => {
    return sum + countRecordMultiSelectItems(record, fieldName);
  }, 0);
}

/**
 * 服務量統計
 */
function countServiceVolume(records) {
  const result = {};
  const identities = {
    '被害人': ['聲請人', '聲請人-未成年'],
    '相對人': ['相對人', '相對人-未成年'],
    '其他': ['關係人', '關係人-未成年', '民眾', '民眾-未成年']
  };
  
  // 各服務項目
  const items = [
    { name: '社會福利服務', fields: ['會談處遇', '會談服務', '通報', '轉介'], withdrawal: false },
    { name: '陪同出庭', fields: ['陪同出庭', '人身安全'], withdrawal: false },
    { name: '擔任程序監理人', fields: [], withdrawal: false },
    { name: '法律服務', fields: ['法律服務'], withdrawal: false },
    { name: '保護令撤回案件諮詢服務', fields: ['陪同出庭', '法律服務', '會談處遇', '會談服務', '通報', '轉介', '聯繫', '人身安全'], withdrawal: true },
    { name: '網絡聯繫', fields: ['聯繫'], withdrawal: false },
    { name: '其他', fields: [], withdrawal: false }
  ];
  
  items.forEach(item => {
    result[item.name] = { 合計: 0, 被害人: {}, 相對人: {}, 其他: {} };
    
    Object.keys(identities).forEach(groupName => {
      const identityValue = identities[groupName];
      result[item.name][groupName] = { 小計: 0, 女: 0, 男: 0 };
      
      ['女', '男'].forEach(gender => {
        const subset = records.filter(r => {
          const 身分 = r['個案身分'] || '';
          const 主題 = r['服務主題'];
          
          let identityMatch = false;
          if (Array.isArray(identityValue)) {
            identityMatch = identityValue.some(v => 身分.includes(v));
          } else {
            identityMatch = 身分.includes(identityValue);
          }
          
          const withdrawalMatch = item.withdrawal ? 主題 === '撤回' : 主題 !== '撤回';
          
          return identityMatch && r['個案性別'] === gender && withdrawalMatch;
        });
        
        let count = 0;
        if (item.fields.length > 0) {
          item.fields.forEach(field => {
            count += sumMultiSelectCounts(subset, field);
          });
        }
        
        result[item.name][groupName][gender] = count;
        result[item.name][groupName]['小計'] += count;
      });
      
      result[item.name]['合計'] += result[item.name][groupName]['小計'];
    });
  });
  
  // 合計行
  result['合計'] = { 合計: 0, 被害人: { 小計: 0, 女: 0, 男: 0 }, 相對人: { 小計: 0, 女: 0, 男: 0 }, 其他: { 小計: 0, 女: 0, 男: 0 } };
  items.forEach(item => {
    Object.keys(identities).forEach(groupName => {
      result['合計'][groupName]['女'] += result[item.name][groupName]['女'];
      result['合計'][groupName]['男'] += result[item.name][groupName]['男'];
      result['合計'][groupName]['小計'] += result[item.name][groupName]['小計'];
    });
    result['合計']['合計'] += result[item.name]['合計'];
  });
  
  return result;
}

/**
 * 產生報表 HTML (新版格式: 2026/01/15 更新)
 */
function generateReportHtml(stats, year, month) {
  const rocYear = year - 1911;
  const monthText = month === '全年' ? '全年' : month + '月';
  
  let html = `<!DOCTYPE html>
<html>
<head>
<meta charset='UTF-8'>
<style>
@page { 
  size: A4; 
  margin: 10mm 10mm; /* 縮小邊界以容納更多內容 */
}
@media print { 
  body { 
    -webkit-print-color-adjust: exact; 
    print-color-adjust: exact; 
    font-size: 10pt !important; /* 縮小整體字體 */
  } 
  table { 
    page-break-inside: avoid; 
    margin-bottom: 15px !important; /* 縮小表格下方間距 */
  } 
  tr, td, th { page-break-inside: avoid; } /* 避免列印時切斷列 */
  h1 { 
    page-break-after: avoid; 
    font-size: 11pt !important; 
    margin-bottom: 5px !important; 
    margin-top: 15px !important; 
  }
  h1.main-title {
    margin-bottom: 5px !important;
    margin-top: 0 !important;
  }
  h2 { font-size: 11pt !important; margin-bottom: 2px !important; margin-top: 2px !important;}
  th, td { 
    padding: 2px !important; /* 更小表格內距 */
    font-size: 9.5pt !important; /* 字體再縮小一點點 */
  }
  .diagonal-container { height: 50px !important; } /* 縮小斜線表頭 */
  .print-box { 
    min-height: 80px !important; /* 最小方框高度，允許自動長高 */
  }
  /* 隱藏列印時的頁首頁尾 (如網址、日期、頁碼) */
  @page { margin-top: 5mm; margin-bottom: 5mm; }
}

body { 
  font-family: 'Microsoft JhengHei', Arial, sans-serif; 
  margin: 0 auto; /* 置中 */
  padding: 20px; 
  font-size: 11pt; 
  line-height: 1.3; 
  color: #000; 
  max-width: 210mm; /* 限制最大寬度為A4大小 */
  background-color: white; /* 確保背景是白色 */
  box-sizing: border-box; /* 確保 padding 算在寬度內 */
}
h1 { font-size: 13pt; text-align: left; margin-bottom: 5px; font-weight: bold; margin-top: 10px; }
h1.main-title { font-size: 20px; text-align: center; margin-bottom: 15px; margin-top: 0; }
table { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
th, td { border: 1px solid #000; padding: 4px; text-align: center; font-size: 11pt; }

/* 浮動列印按鈕樣式 */
.print-btn {
  display: none; /* 預設隱藏 */
  position: fixed;
  bottom: 20px;
  right: 20px;
  padding: 12px 24px;
  background-color: #28a745; /* 改為綠色，看起來更像完成的動作 */
  color: white;
  border: none;
  border-radius: 5px;
  font-size: 16pt;
  font-weight: bold;
  cursor: pointer;
  box-shadow: 0 4px 6px rgba(0,0,0,0.3);
  z-index: 1000;
  transition: opacity 0.3s ease-in-out;
}
.print-btn:hover { background-color: #218838; }
@media print { .print-btn { display: none !important; } }

/* 讓 div 可打字、且列印時會出現框線與內文（會自動長高） */
div.print-box {
  width: 100%;
  min-height: 100px;
  border: 2px solid #000;
  box-sizing: border-box;
  font-family: 'Microsoft JhengHei', Arial, sans-serif;
  font-size: 11pt;
  padding: 5px;
  outline: none;
  overflow: hidden; /* 防止出現卷軸，印表機會印出全部內文 */
  white-space: pre-wrap; /* 允許換行 */
  word-break: break-all;
}
div.print-box:empty:before {
  content: attr(placeholder);
  color: #888;
}

/* 對角線表頭樣式 */
.diagonal-container {
  position: relative;
  width: 100%;
  height: 60px; /* 變矮以容納內容 */
  overflow: hidden;
}
.diagonal-line {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background: linear-gradient(to top right, transparent 49.5%, #000 49.5%, #000 50.5%, transparent 50.5%);
  z-index: 1;
}
.th-content {
  position: relative;
  width: 100%;
  height: 100%;
  z-index: 2;
}
.th-method {
  position: absolute;
  top: 5px;
  right: 10px;
  font-weight: bold;
}
.th-item {
  position: absolute;
  bottom: 5px;
  left: 10px;
  font-weight: bold;
}

/* 斜線儲存格 */
.slash {
  background: linear-gradient(to top right, transparent 49.5%, #000 49.5%, #000 50.5%, transparent 50.5%);
  background-color: #f0f0f0; /* 選用：加上淡灰色背景 */
}

/* 文字對齊調整 */
.left-align { text-align: left; padding-left: 15px; }

</style>
</head>
<body>

<!-- 新增：大標題 -->
<h1 class="main-title">
  苗栗縣政府 ${stats.rocYear} 年家庭暴力事件服務處 ${stats.monthText} 服務統計表
</h1>

<h1>壹、個案服務量：服務 ${stats.totalDistinctCases} 人(陪同出庭 ${stats.accompCases} 人； 諮詢服務 ${stats.consultCases} 人)； ${stats.totalPersonTimes} 人次； ${stats.totalItemTimes} 項次</h1>

<table>
  <thead>
    <tr>
      <th style="width: 250px; padding: 0;">
        <div class="diagonal-container">
            <div class="diagonal-line"></div>
            <div class="th-content">
                <span style="position: absolute; top: 5px; right: 10px; font-size: 14px;">服務方式</span>
                <span style="position: absolute; bottom: 5px; left: 10px; font-size: 14px;">項目</span>
            </div>
        </div>
      </th>
      <th style="vertical-align: middle;">電話諮詢</th>
      <th style="vertical-align: middle;">臨櫃面談</th>
      <th style="vertical-align: middle;">書面諮詢</th>
      <th style="vertical-align: middle;">小計</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="left-align">陪同出庭服務</td>
      <td>${stats.serviceMethod.陪同出庭.電談 + stats.serviceMethod.陪同出庭.聯繫未果 || ''}</td>
      <td>${stats.serviceMethod.陪同出庭.面談 || ''}</td>
      <td>${stats.serviceMethod.陪同出庭.書面 || ''}</td>
      <td>${stats.serviceMethod.陪同出庭.小計}</td>
    </tr>
    <tr>
      <td class="left-align">諮詢服務</td>
      <td>${stats.serviceMethod.諮詢服務.電談 + stats.serviceMethod.諮詢服務.聯繫未果 || ''}</td>
      <td>${stats.serviceMethod.諮詢服務.面談 || ''}</td>
      <td>${stats.serviceMethod.諮詢服務.書面 || ''}</td>
      <td>${stats.serviceMethod.諮詢服務.小計}</td>
    </tr>
    <tr>
      <td class="left-align">總計(項次)</td>
      <td>${stats.serviceMethod.總計.電談 + stats.serviceMethod.總計.聯繫未果}</td>
      <td>${stats.serviceMethod.陪同出庭.面談 + stats.serviceMethod.諮詢服務.面談}</td>
      <td>${stats.serviceMethod.總計.書面 || ''}</td>
      <td>${stats.serviceMethod.總計.小計}</td>
    </tr>
  </tbody>
</table>


<h1>貳、案件來源(${stats.totalPersonTimes} 人次)</h1>

<table>
  <thead>
    <tr>
      <th colspan="2" style="width: 250px; padding: 0;">
        <div class="diagonal-container">
            <div class="diagonal-line"></div>
            <div class="th-content">
                <span style="position: absolute; top: 5px; right: 10px; font-size: 14px;">服務結果</span>
                <span style="position: absolute; bottom: 5px; left: 10px; font-size: 14px;">服務個案資料</span>
            </div>
        </div>
      </th>
      <th style="vertical-align: middle;">諮詢服務</th>
      <th style="vertical-align: middle;">陪同出庭服務</th>
      <th style="vertical-align: middle;">小計</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td rowspan="5" style="width: 30px; vertical-align: middle;">個案來源</td>
      <td>司法系統</td>
      <td>${stats.caseSource['諮詢服務']['接受轉介-司法系統']}</td>
      <td>${stats.caseSource['陪同出庭']['接受轉介-司法系統']}</td>
      <td>${stats.caseSource['總計']['接受轉介-司法系統']}</td>
    </tr>
    <tr>
      <td>社會處</td>
      <td>${stats.caseSource['諮詢服務']['接受轉介-社會處']}</td>
      <td>${stats.caseSource['陪同出庭']['接受轉介-社會處']}</td>
      <td>${stats.caseSource['總計']['接受轉介-社會處']}</td>
    </tr>
    <tr>
      <td>民間團體</td>
      <td>${stats.caseSource['諮詢服務']['接受轉介-民間團體']}</td>
      <td>${stats.caseSource['陪同出庭']['接受轉介-民間團體']}</td>
      <td>${stats.caseSource['總計']['接受轉介-民間團體']}</td>
    </tr>
    <tr>
      <td>自行求助</td>
      <td>${stats.caseSource['諮詢服務']['自行求助']}</td>
      <td>${stats.caseSource['陪同出庭']['自行求助']}</td>
      <td>${stats.caseSource['總計']['自行求助']}</td>
    </tr>
    <tr>
      <td>自行求助-網絡諮詢</td>
      <td>${stats.caseSource['諮詢服務']['自行求助-網絡諮詢']}</td>
      <td>${stats.caseSource['陪同出庭']['自行求助-網絡諮詢']}</td>
      <td>${stats.caseSource['總計']['自行求助-網絡諮詢']}</td>
    </tr>
    <tr>
      <td colspan="2">合 計(人次)</td>
      <td>${stats.caseSource['諮詢服務']['小計']}</td>
      <td>${stats.caseSource['陪同出庭']['小計']}</td>
      <td>${stats.caseSource['總計']['小計']}</td>
    </tr>
  </tbody>
</table>


${(() => {
  const nw = stats.nonWithdrawal;
  const totalPeople = nw.總計.女.服務人數 + nw.總計.男.服務人數;
  const totalItems = nw.總計.女.總計 + nw.總計.男.總計;
  
  return `<h1>參、保護令非撤回案件服務(服務 ${totalPeople} 人次, ${totalItems} 項次)</h1>
<table>
  <thead>
    <tr>
      <th rowspan="2" colspan="2" style="width: 150px; vertical-align: middle;">項目</th>
      <th colspan="2">陪同服務</th>
      <th rowspan="2" style="vertical-align: middle;">法律<br>服務</th>
      <th colspan="3">社會福利</th>
      <th rowspan="2" style="vertical-align: middle;">網絡<br>聯繫</th>
      <th rowspan="2" style="vertical-align: middle;">合計</th>
    </tr>
    <tr>
      <th>陪同<br>出庭</th>
      <th>人身<br>安全</th>
      <th>會談<br>服務</th>
      <th>通報</th>
      <th>轉介</th>
    </tr>
  </thead>
  <tbody>
    ${generateIdentityGenderRowsV2(stats.nonWithdrawal)}
  </tbody>
</table>`;
})()}


${(() => {
  const wd = stats.withdrawal;
  const totalPeoplewd = wd.總計.女.服務人數 + wd.總計.男.服務人數;
  const totalItemswd = wd.總計.女.總計 + wd.總計.男.總計;
  
  return `<h1>肆、保護令撤回案件服務(服務 ${totalPeoplewd} 人次, ${totalItemswd} 項次)</h1>
<table>
  <thead>
    <tr>
      <th rowspan="2" colspan="2" style="width: 150px; vertical-align: middle;">項目</th>
      <th colspan="2">陪同服務</th>
      <th rowspan="2" style="vertical-align: middle;">法律<br>服務</th>
      <th colspan="3">社會福利</th>
      <th rowspan="2" style="vertical-align: middle;">網絡<br>聯繫</th>
      <th rowspan="2" style="vertical-align: middle;">合計</th>
    </tr>
    <tr>
      <th>陪同<br>出庭</th>
      <th>人身<br>安全</th>
      <th>會談<br>服務</th>
      <th>通報</th>
      <th>轉介</th>
    </tr>
  </thead>
  <tbody>
    ${generateIdentityGenderRowsV2(stats.withdrawal)}
  </tbody>
</table>`;
})()}




${(() => {
  const st = stats.serviceTypeStats;
  return `<h1>伍、服務案件類型(${st.總計} 人次) (其他-抗告、傷害、跟騷)</h1>
<table>
  <thead>
    <tr>
      <th>暫時</th>
      <th>通常</th>
      <th>變更</th>
      <th>延長</th>
      <th>撤銷</th>
      <th>撤回</th>
      <th>追加</th>
      <th>違反</th>
      <th rowspan="2" style="vertical-align: middle;">其他</th>
      <th rowspan="2" style="vertical-align: middle;">總計</th>
    </tr>
    <tr>
      <th>保護令</th>
      <th>保護令</th>
      <th>保護令</th>
      <th>保護令</th>
      <th>保護令</th>
      <th>保護令</th>
      <th>保護令</th>
      <th>保護令</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>${st.暫時}</td>
      <td>${st.通常}</td>
      <td>${st.變更}</td>
      <td>${st.延長}</td>
      <td>${st.撤銷}</td>
      <td>${st.撤回}</td>
      <td>${st.追加}</td>
      <td>${st.違反}</td>
      <td>${st.其他}</td>
      <td>${st.總計}</td>
    </tr>
  </tbody>
</table>`;
})()}



${(() => {
  const sv = stats.serviceVolume;
  
  // 生成服務量統計表格行 HTML
  const generateVolRows = () => {
    const items = ['社會福利服務', '陪同出庭', '擔任程序監理人', '法律服務', '保護令撤回案件諮詢服務', '網絡聯繫', '其他', '合計'];
    let rows = '';
    
    items.forEach((item, idx) => {
      const d = sv[item] || { 合計: 0, 被害人: { 小計: 0, 女: 0, 男: 0 }, 相對人: { 小計: 0, 女: 0, 男: 0 }, 其他: { 小計: 0, 女: 0, 男: 0 } };
      
      rows += `<tr>
        <td style="text-align: left;">${item}</td>
        <td>${d.合計 || 0}</td>
        <td>${d.被害人 ? d.被害人.小計 : 0}</td>
        <td>${d.被害人 ? d.被害人.女 : 0}</td>
        <td>${d.被害人 ? d.被害人.男 : 0}</td>
        <td>${d.相對人 ? d.相對人.小計 : 0}</td>
        <td>${d.相對人 ? d.相對人.女 : 0}</td>
        <td>${d.相對人 ? d.相對人.男 : 0}</td>
        <td>${d.其他 ? d.其他.小計 : 0}</td>
        <td>${d.其他 ? d.其他.女 : 0}</td>
        <td>${d.其他 ? d.其他.男 : 0}</td>
        <td></td>
        <td></td>
        <td></td>
      </tr>`;
    });
    return rows;
  };


  return `<h1>陸、服務量統計表</h1>
<table>
  <thead>
    <tr>
      <th style="border-bottom: none; text-align: right;">身分</th>
      <th style="border-bottom: none;"></th>
      <th colspan="3">被害人（人次）</th>
      <th colspan="3">相對人（人次）</th>
      <th colspan="3">其他（人次）</th>
      <th rowspan="2">聯繫會報<br>（場次）</th>
      <th rowspan="2">督導會議<br>（場次）</th>
      <th rowspan="2">宣導活動<br>（場次）</th>
    </tr>
    <tr>
      <th style="border-top: none; text-align: left;">性別 / 服務項目</th>
      <th style="border-top: none;">合計</th>
      <th>小計</th>
      <th>女</th>
      <th>男</th>
      <th>小計</th>
      <th>女</th>
      <th>男</th>
      <th>小計</th>
      <th>女</th>
      <th>男</th>
    </tr>
  </thead>
  <tbody>
    ${generateVolRows()}
  </tbody>
</table>

<!-- 新增：柒、捌 與簽核欄位 -->
<div style="margin-top: 30px; display: flex; justify-content: space-between;">
  <div style="width: 48%;">
    <h2>柒、本月大事記：</h2>
    <div id="this-month-note" class="print-box" contenteditable="true" placeholder="請輸入本月大事記..."></div>
  </div>
  <div style="width: 48%;">
    <h2>捌、下月大事記：</h2>
    <div id="next-month-note" class="print-box" contenteditable="true" placeholder="請輸入下月大事記..."></div>
  </div>
</div>

<div style="margin-top: 40px; display: flex; justify-content: space-between; font-size: 16px;">
  <div style="width: 33%;">主任：</div>
  <div style="width: 33%;">督導：</div>
  <div style="width: 33%;">填表人：</div>
</div>

<button id="floating-print-btn" class="print-btn" onclick="window.print()">🖨️ 列印報表</button>

<script>
  const thisMonth = document.getElementById('this-month-note');
  const nextMonth = document.getElementById('next-month-note');
  const printBtn = document.getElementById('floating-print-btn');

  function checkFieldsAndTogglePrintBtn() {
    if (thisMonth && nextMonth && printBtn) {
      const thisText = thisMonth.innerText || thisMonth.textContent;
      const nextText = nextMonth.innerText || nextMonth.textContent;
      if (thisText.trim() !== '' && nextText.trim() !== '') {
        printBtn.style.display = 'block';
      } else {
        printBtn.style.display = 'none';
      }
    }
  }

  if (thisMonth && nextMonth) {
    thisMonth.addEventListener('input', checkFieldsAndTogglePrintBtn);
    nextMonth.addEventListener('input', checkFieldsAndTogglePrintBtn);
  }
</script>
`;
})()}

</body>
</html>`;
  
  return html;
}

/**
 * 生成身分/性別統計表格行 (簡潔版)
 */
function generateIdentityGenderRowsSimple(data) {
  const identities = ['聲請人', '相對人', '關係人', '民眾'];
  const genders = ['女', '男'];
  
  let rows = '';
  
  identities.forEach((identity, idxI) => {
    genders.forEach((gender, idxG) => {
      const d = data[identity] && data[identity][gender] ? data[identity][gender] : {};
      const total = (d.陪同出庭 || 0) + (d.法律服務 || 0) + (d.會談處遇 || 0) + (d.通報 || 0) + (d.轉介 || 0) + (d.聯繫 || 0);
      const rowClass = (idxI * 2 + idxG) % 2 === 1 ? " class='alt'" : "";
      
      rows += `<tr${rowClass}><td class='lbl'>${identity}</td><td>${gender}</td><td>${d.服務人數 || 0}</td><td>${d.陪同出庭 || 0}</td><td>${d.法律服務 || 0}</td><td>${d.會談處遇 || 0}</td><td>${d.通報 || 0}</td><td>${d.轉介 || 0}</td><td>${d.聯繫 || 0}</td><td class='b'>${total}</td></tr>`;
    });
  });
  
  const totalF = data['總計'] && data['總計']['女'] ? data['總計']['女'] : {};
  const totalM = data['總計'] && data['總計']['男'] ? data['總計']['男'] : {};
  const grandTotal = (totalF.總計 || 0) + (totalM.總計 || 0);
  
  rows += `<tr class='total'><td class='lbl' colspan='2'>總計</td><td>女${totalF.服務人數 || 0}/男${totalM.服務人數 || 0}</td><td>${(totalF.陪同出庭 || 0) + (totalM.陪同出庭 || 0)}</td><td>${(totalF.法律服務 || 0) + (totalM.法律服務 || 0)}</td><td>${(totalF.會談處遇 || 0) + (totalM.會談處遇 || 0)}</td><td>${(totalF.通報 || 0) + (totalM.通報 || 0)}</td><td>${(totalF.轉介 || 0) + (totalM.轉介 || 0)}</td><td>${(totalF.聯繫 || 0) + (totalM.聯繫 || 0)}</td><td class='grand'>${grandTotal}</td></tr>`;
  
  return rows;
}

/**
 * 生成服務量統計表格行 (簡潔版)
 */
function generateServiceVolumeRowsSimple(data) {
  const items = ['社會福利服務', '陪同出庭', '擔任程序監理人', '法律服務', '保護令撤回案件諮詢服務', '網絡聯繫', '其他', '合計'];
  
  let rows = '';
  
  items.forEach((item, idx) => {
    const d = data[item] || { 合計: 0, 被害人: { 小計: 0, 女: 0, 男: 0 }, 相對人: { 小計: 0, 女: 0, 男: 0 }, 其他: { 小計: 0, 女: 0, 男: 0 } };
    const isTotal = item === '合計';
    const rowClass = isTotal ? " class='total'" : (idx % 2 === 1 ? " class='alt'" : "");
    
    rows += `<tr${rowClass}><td class='lbl'>${item}</td><td class='${isTotal ? "grand" : ""}'>${d.合計 || 0}</td><td>${d.被害人 ? d.被害人.小計 : 0}</td><td>${d.被害人 ? d.被害人.女 : 0}</td><td>${d.被害人 ? d.被害人.男 : 0}</td><td>${d.相對人 ? d.相對人.小計 : 0}</td><td>${d.相對人 ? d.相對人.女 : 0}</td><td>${d.相對人 ? d.相對人.男 : 0}</td><td>${d.其他 ? d.其他.小計 : 0}</td><td>${d.其他 ? d.其他.女 : 0}</td><td>${d.其他 ? d.其他.男 : 0}</td><td></td><td></td><td></td></tr>`;
  });
  
  return rows;
}

/**
 * 匯出統計報表到 Excel (Google Sheets)
 * @param {number} year - 西元年份
 * @param {string} month - 月份
 * @returns {Object} { url: string, name: string }
 */
function exportStatisticsToExcel(year, month) {
  checkUserPermission();
  
  const rocYear = year - 1911;
  const monthText = month === '全年' ? '全年' : month + '月';
  const fileName = `${rocYear}年${monthText}服務統計表`;
  
  // 取得統計資料 (與 generateStatisticsReport 完全相同)
  const allRecords = getAllRecordsForStats(year);
  const records = filterRecordsByPeriod(allRecords, year, month);
  
  const stats = {
    totalCount: records.length,
    rocYear: rocYear,
    monthText: monthText,
    consultCases: countDistinctCases(records, '諮詢服務'),
    accompCases: countDistinctCases(records, '陪同出庭'),
    totalDistinctCases: countTotalDistinctCases(records),
    totalPersonTimes: records.length,
    totalItemTimes: 0,
    serviceTopic: countByServiceTopic(records),
    serviceItems: countServiceItems(records),
    serviceMethod: countServiceMethodCrossTab(records, year),
    caseSource: countCaseSourceCrossTab(records),
    nonWithdrawal: countByIdentityGender(records, false),
    withdrawal: countByIdentityGender(records, true),
    serviceVolume: countServiceVolume(records),
    serviceTypeStats: countServiceTypes(records)
  };
  stats.totalItemTimes = stats.serviceMethod.總計.小計;
  
  // 建立新的試算表
  const ss = SpreadsheetApp.create(fileName);
  
  // 使用與年報相同的格式產生 Excel (與產生報表內容一致)
  const sheet1 = ss.getSheets()[0];
  sheet1.setName('統計報表');
  createAnnualSummaryHTMLStyle(sheet1, stats);
  
  // 回傳試算表 URL
  return {
    url: ss.getUrl(),
    name: fileName
  };
}

/**
 * 格式化表格範圍
 */
function formatTableRange(sheet, startRow, startCol, numRows, numCols) {
  const range = sheet.getRange(startRow, startCol, numRows, numCols);
  range.setBorder(true, true, true, true, true, true);
  
  // 表頭樣式
  const headerRange = sheet.getRange(startRow, startCol, 1, numCols);
  headerRange.setBackground('#1a5276').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  
  // 資料對齊
  const dataRange = sheet.getRange(startRow + 1, startCol, numRows - 1, numCols);
  dataRange.setHorizontalAlignment('center');
  
  // 第一欄靠左
  sheet.getRange(startRow + 1, startCol, numRows - 1, 1).setHorizontalAlignment('left').setBackground('#d4e6f1');
}

/**
 * 建立身分/性別統計工作表
 */
function createIdentityGenderSheet(sheet, data, title) {
  sheet.getRange(1, 1).setValue(title).setFontSize(12).setFontWeight('bold');
  
  const headers = [['個案身分', '性別', '服務人數', '陪同出庭', '法律服務', '會談處遇', '通報', '轉介', '網絡聯繫', '各項總計']];
  sheet.getRange(3, 1, 1, 10).setValues(headers);
  
  const identities = ['聲請人', '相對人', '關係人', '民眾'];
  const genders = ['女', '男'];
  
  let row = 4;
  identities.forEach(identity => {
    genders.forEach(gender => {
      const d = data[identity] && data[identity][gender] ? data[identity][gender] : {};
      const total = (d.陪同出庭 || 0) + (d.法律服務 || 0) + (d.會談處遇 || 0) + (d.通報 || 0) + (d.轉介 || 0) + (d.聯繫 || 0);
      sheet.getRange(row, 1, 1, 10).setValues([[identity, gender, d.服務人數 || 0, d.陪同出庭 || 0, d.法律服務 || 0, d.會談處遇 || 0, d.通報 || 0, d.轉介 || 0, d.聯繫 || 0, total]]);
      row++;
    });
  });
  
  // 總計行
  const totalF = data['總計'] && data['總計']['女'] ? data['總計']['女'] : {};
  const totalM = data['總計'] && data['總計']['男'] ? data['總計']['男'] : {};
  const grandTotal = (totalF.總計 || 0) + (totalM.總計 || 0);
  
  sheet.getRange(row, 1, 1, 10).setValues([['總計', '', 
    `女${totalF.服務人數 || 0}/男${totalM.服務人數 || 0}`,
    (totalF.陪同出庭 || 0) + (totalM.陪同出庭 || 0),
    (totalF.法律服務 || 0) + (totalM.法律服務 || 0),
    (totalF.會談處遇 || 0) + (totalM.會談處遇 || 0),
    (totalF.通報 || 0) + (totalM.通報 || 0),
    (totalF.轉介 || 0) + (totalM.轉介 || 0),
    (totalF.聯繫 || 0) + (totalM.聯繫 || 0),
    grandTotal
  ]]);
  
  formatTableRange(sheet, 3, 1, row - 2, 10);
  sheet.autoResizeColumns(1, 10);
}

/**
 * 建立服務量統計工作表
 */
function createServiceVolumeSheet(sheet, data, title) {
  sheet.getRange(1, 1).setValue(title).setFontSize(12).setFontWeight('bold');
  
  // 合併表頭
  sheet.getRange(3, 1).setValue('服務項目');
  sheet.getRange(3, 2).setValue('合計');
  sheet.getRange(3, 3).setValue('被害人(人次)');
  sheet.getRange(3, 3, 1, 3).merge();
  sheet.getRange(3, 6).setValue('相對人(人次)');
  sheet.getRange(3, 6, 1, 3).merge();
  sheet.getRange(3, 9).setValue('其他(人次)');
  sheet.getRange(3, 9, 1, 3).merge();
  sheet.getRange(3, 12).setValue('聯繫會報(場次)');
  sheet.getRange(3, 13).setValue('督導會議(場次)');
  sheet.getRange(3, 14).setValue('宣導活動(場次)');
  
  sheet.getRange(4, 3, 1, 9).setValues([['小計', '女', '男', '小計', '女', '男', '小計', '女', '男']]);
  
  const items = ['社會福利服務', '陪同出庭', '擔任程序監理人', '法律服務', '保護令撤回案件諮詢服務', '網絡聯繫', '其他', '合計'];
  
  let row = 5;
  items.forEach(item => {
    const d = data[item] || { 合計: 0, 被害人: { 小計: 0, 女: 0, 男: 0 }, 相對人: { 小計: 0, 女: 0, 男: 0 }, 其他: { 小計: 0, 女: 0, 男: 0 } };
    sheet.getRange(row, 1, 1, 14).setValues([[
      item, d.合計 || 0,
      d.被害人 ? d.被害人.小計 : 0, d.被害人 ? d.被害人.女 : 0, d.被害人 ? d.被害人.男 : 0,
      d.相對人 ? d.相對人.小計 : 0, d.相對人 ? d.相對人.女 : 0, d.相對人 ? d.相對人.男 : 0,
      d.其他 ? d.其他.小計 : 0, d.其他 ? d.其他.女 : 0, d.其他 ? d.其他.男 : 0,
      '', '', ''
    ]]);
    row++;
  });
  
  // 格式化
  const headerRange = sheet.getRange(3, 1, 2, 14);
  headerRange.setBorder(true, true, true, true, true, true);
  headerRange.setBackground('#1a5276').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  
  const dataRange = sheet.getRange(5, 1, items.length, 14);
  dataRange.setBorder(true, true, true, true, true, true);
  dataRange.setHorizontalAlignment('center');
  
  sheet.getRange(5, 1, items.length, 1).setHorizontalAlignment('left').setBackground('#d4e6f1');
  sheet.getRange(5 + items.length - 1, 1, 1, 14).setBackground('#aed6f1').setFontWeight('bold');
  
  sheet.autoResizeColumns(1, 14);
}

// ================================================================
// 年報匯出功能
// ================================================================

/**
 * 匯出年度報表 (完整年報 Excel)
 * @param {number} rocYear - 民國年份 (如 114)
 * @returns {Object} { url, name }
 */
function exportAnnualReport(rocYear) {
  checkUserPermission();
  
  const year = rocYear + 1911; // 轉換為西元年
  const fileName = `${rocYear}年家暴服務紀錄年報`;
  
  // 取得該年度的所有資料 (傳入西元年份確保讀取正確的年度資料表)
  const allRecords = getAllRecordsForStats(year);
  const yearRecords = filterRecordsByPeriod(allRecords, year, '全年');
  
  // 建立新的試算表
  const ss = SpreadsheetApp.create(fileName);
  
  // 計算各種統計數據 (同 generateStatisticsReport)
  const stats = {
    totalCount: yearRecords.length,
    rocYear: rocYear,
    monthText: '全年',
    consultCases: countDistinctCases(yearRecords, '諮詢服務'),
    accompCases: countDistinctCases(yearRecords, '陪同出庭'),
    totalDistinctCases: countTotalDistinctCases(yearRecords),
    totalPersonTimes: yearRecords.length,
    serviceTopic: countByServiceTopic(yearRecords),
    serviceItems: countServiceItems(yearRecords),
    serviceMethod: countServiceMethodCrossTab(yearRecords, year),
    caseSource: countCaseSourceCrossTab(yearRecords),
    nonWithdrawal: countByIdentityGender(yearRecords, false),
    withdrawal: countByIdentityGender(yearRecords, true),
    serviceVolume: countServiceVolume(yearRecords),
    serviceTypeStats: countServiceTypes(yearRecords)
  };
  stats.totalItemTimes = stats.serviceMethod.總計.小計; // 總項次
  
  // === 唯一的工作表：年度總表 (比照 HTML 報告) ===
  const sheet1 = ss.getSheets()[0];
  sheet1.setName('年度總表');
  createAnnualSummaryHTMLStyle(sheet1, stats);
  
  // 為了防止有人找不到原始資料，保留 {年份}_data
  const sheet2 = ss.insertSheet(rocYear + '_data');
  createYearDataSheet(sheet2, yearRecords);
  
  return {
    url: ss.getUrl(),
    name: fileName
  };
}

/**
 * 依據 HTML 報告格式建立 Excel 總表
 */
function createAnnualSummaryHTMLStyle(sheet, stats) {
  let row = 1;
  
  // 標題設定 helper
  const setMainTitle = (text) => {
    sheet.getRange(row, 1).setValue(text).setFontWeight('bold').setFontSize(14);
    row += 2;
  };
  
  // 表格框線與對齊歸中 helper
  const formatTable = (startRow, startCol, numRows, numCols) => {
    const range = sheet.getRange(startRow, startCol, numRows, numCols);
    range.setBorder(true, true, true, true, true, true);
    range.setHorizontalAlignment('center').setVerticalAlignment('middle');
    // 第一行為粗體
    sheet.getRange(startRow, startCol, 1, numCols).setFontWeight('bold');
  };

  // ----------------------------------------------------
  // 壹、個案服務量
  // ----------------------------------------------------
  setMainTitle(`壹、個案服務量：服務 ${stats.totalDistinctCases} 人(陪同出庭 ${stats.accompCases} 人； 諮詢服務 ${stats.consultCases} 人)； ${stats.totalPersonTimes} 人次； ${stats.totalItemTimes} 項次`);
  
  // 表單標頭
  sheet.getRange(row, 1, 1, 5).setValues([['服務方式項目', '電話諮詢', '臨櫃面談', '書面諮詢', '小計']]);
  // 資料列
  sheet.getRange(row + 1, 1, 3, 5).setValues([
    ['陪同出庭服務', stats.serviceMethod.陪同出庭.電談 + stats.serviceMethod.陪同出庭.聯繫未果 || '', stats.serviceMethod.陪同出庭.面談 || '', stats.serviceMethod.陪同出庭.書面 || '', stats.serviceMethod.陪同出庭.小計],
    ['諮詢服務', stats.serviceMethod.諮詢服務.電談 + stats.serviceMethod.諮詢服務.聯繫未果 || '', stats.serviceMethod.諮詢服務.面談 || '', stats.serviceMethod.諮詢服務.書面 || '', stats.serviceMethod.諮詢服務.小計],
    ['總計(項次)', stats.serviceMethod.總計.電談 + stats.serviceMethod.總計.聯繫未果, stats.serviceMethod.陪同出庭.面談 + stats.serviceMethod.諮詢服務.面談, stats.serviceMethod.總計.書面 || '', stats.serviceMethod.總計.小計]
  ]);
  formatTable(row, 1, 4, 5);
  row += 6;

  // ----------------------------------------------------
  // 貳、案件來源
  // ----------------------------------------------------
  setMainTitle(`貳、案件來源(${stats.totalPersonTimes} 人次)`);
  
  // 表單標頭
  sheet.getRange(row, 1, 1, 5).setValues([['服務結果服務個案資料', '', '諮詢服務', '陪同出庭服務', '小計']]);
  sheet.getRange(row, 1, 1, 2).merge();
  sheet.getRange(row + 1, 1, 6, 5).setValues([
    ['個案來源', '司法系統', stats.caseSource['諮詢服務']['接受轉介-司法系統'], stats.caseSource['陪同出庭']['接受轉介-司法系統'], stats.caseSource['總計']['接受轉介-司法系統']],
    ['', '社會處', stats.caseSource['諮詢服務']['接受轉介-社會處'], stats.caseSource['陪同出庭']['接受轉介-社會處'], stats.caseSource['總計']['接受轉介-社會處']],
    ['', '民間團體', stats.caseSource['諮詢服務']['接受轉介-民間團體'], stats.caseSource['陪同出庭']['接受轉介-民間團體'], stats.caseSource['總計']['接受轉介-民間團體']],
    ['', '自行求助', stats.caseSource['諮詢服務']['自行求助'], stats.caseSource['陪同出庭']['自行求助'], stats.caseSource['總計']['自行求助']],
    ['', '自行求助-網絡諮詢', stats.caseSource['諮詢服務']['自行求助-網絡諮詢'], stats.caseSource['陪同出庭']['自行求助-網絡諮詢'], stats.caseSource['總計']['自行求助-網絡諮詢']],
    ['合 計(人次)', '', stats.caseSource['諮詢服務']['小計'], stats.caseSource['陪同出庭']['小計'], stats.caseSource['總計']['小計']]
  ]);
  sheet.getRange(row + 1, 1, 5, 1).merge(); // 合併個案來源
  sheet.getRange(row + 6, 1, 1, 2).merge(); // 合計(人次)
  formatTable(row, 1, 7, 5);
  row += 9;

  // 共用的六大服務產生器 (給參、肆使用)
  const generateIdentityRowsToExcel = (statObj, titlePrefix) => {
    const totalPeople = statObj.總計.女.服務人數 + statObj.總計.男.服務人數;
    const totalItems = statObj.總計.女.總計 + statObj.總計.男.總計;
    setMainTitle(`${titlePrefix}(服務 ${totalPeople} 人次, ${totalItems} 項次)`);
    
    sheet.getRange(row, 1, 2, 2).merge().setValue('項目');
    sheet.getRange(row, 3, 1, 2).merge().setValue('陪同服務');
    sheet.getRange(row, 5, 2, 1).merge().setValue('法律\n服務');
    sheet.getRange(row, 6, 1, 3).merge().setValue('社會福利');
    sheet.getRange(row, 9, 2, 1).merge().setValue('網絡\n聯繫');
    sheet.getRange(row, 10, 2, 1).merge().setValue('合計');
    sheet.getRange(row + 1, 3, 1, 6).setValues([['陪同\n出庭', '人身\n安全', '會談\n服務', '通報', '轉介', '']]);
    
    // 設定表頭合併與對齊
    formatTable(row, 1, 2, 10);
    
    // 寫入資料
    let startRow = row + 2;
    
    // 身分列舉 (不含總計)
    ['聲請人', '相對人', '關係人', '民眾'].forEach(identity => {
      const gObj = statObj[identity];
      if (!gObj) return;
      
      const maleObj = gObj['男'];
      const femaleObj = gObj['女'];
      
      // 女
      sheet.getRange(startRow, 1).setValue(identity);
      sheet.getRange(startRow, 1, 2, 1).merge(); // 每個身分合併兩列
      sheet.getRange(startRow, 2).setValue(`女(${femaleObj['服務人數']} 人次)`);
      sheet.getRange(startRow, 3, 1, 8).setValues([[
        femaleObj['陪同出庭'], femaleObj['人身安全'], femaleObj['法律服務'], 
        femaleObj['會談服務'], femaleObj['通報'], femaleObj['轉介'], 
        femaleObj['聯繫'] || 0, femaleObj['總計']
      ]]);
      startRow++;
      
      // 男
      sheet.getRange(startRow, 2).setValue(`男(${maleObj['服務人數']} 人次)`);
      sheet.getRange(startRow, 3, 1, 8).setValues([[
        maleObj['陪同出庭'], maleObj['人身安全'], maleObj['法律服務'], 
        maleObj['會談服務'], maleObj['通報'], maleObj['轉介'], 
        maleObj['聯繫'] || 0, maleObj['總計']
      ]]);
      startRow++;
    });
    
    // 總計列 (合併女男為一行)
    const totalFem = statObj['總計']['女'];
    const totalMale = statObj['總計']['男'];
    sheet.getRange(startRow, 1).setValue('總 計');
    sheet.getRange(startRow, 2).setValue(`女(${totalFem['服務人數']} 人次)\n男(${totalMale['服務人數']} 人次)`);
    sheet.getRange(startRow, 3, 1, 8).setValues([[
      totalFem['陪同出庭'] + totalMale['陪同出庭'],
      totalFem['人身安全'] + totalMale['人身安全'],
      totalFem['法律服務'] + totalMale['法律服務'],
      totalFem['會談服務'] + totalMale['會談服務'],
      totalFem['通報'] + totalMale['通報'],
      totalFem['轉介'] + totalMale['轉介'],
      (totalFem['聯繫'] || 0) + (totalMale['聯繫'] || 0),
      totalFem['總計'] + totalMale['總計']
    ]]);
    startRow++;
    
    const dataRows = 4 * 2 + 1; // 4個身分 * 2列 + 總計1列 = 9行
    formatTable(row + 2, 1, dataRows, 10);
    sheet.getRange(row + 2, 1, dataRows, 1).setVerticalAlignment('middle');
    row += dataRows + 4;
  };

  // ----------------------------------------------------
  // 參、保護令非撤回案件服務
  // ----------------------------------------------------
  generateIdentityRowsToExcel(stats.nonWithdrawal, '參、保護令非撤回案件服務');

  // ----------------------------------------------------
  // 肆、保護令撤回案件服務
  // ----------------------------------------------------
  generateIdentityRowsToExcel(stats.withdrawal, '肆、保護令撤回案件服務');

  // ----------------------------------------------------
  // 伍、服務案件類型
  // ----------------------------------------------------
  const st = stats.serviceTypeStats;
  setMainTitle(`伍、服務案件類型(${st.總計} 人次) (其他-抗告、傷害、跟騷)`);
  
  sheet.getRange(row, 1, 2, 10).setValues([
    ['暫時', '通常', '變更', '延長', '撤銷', '撤回', '追加', '違反', '其他', '總計'],
    ['保護令', '保護令', '保護令', '保護令', '保護令', '保護令', '保護令', '保護令', '', '']
  ]);
  sheet.getRange(row, 9, 2, 1).merge().setValue('其他');
  sheet.getRange(row, 10, 2, 1).merge().setValue('總計');
  
  sheet.getRange(row + 2, 1, 1, 10).setValues([[st.暫時, st.通常, st.變更, st.延長, st.撤銷, st.撤回, st.追加, st.違反, st.其他, st.總計]]);
  formatTable(row, 1, 3, 10);
  row += 5;

  // ----------------------------------------------------
  // 陸、服務量統計表
  // ----------------------------------------------------
  setMainTitle(`陸、服務量統計表`);
  const sv = stats.serviceVolume;
  
  sheet.getRange(row, 1, 2, 14).setValues([
    ['身分', '', '被害人（人次）', '', '', '相對人（人次）', '', '', '其他（人次）', '', '', '聯繫會報\n（場次）', '督導會議\n（場次）', '宣導活動\n（場次）'],
    ['性別\n服務項目', '合計', '小計', '女', '男', '小計', '女', '男', '小計', '女', '男', '', '', '']
  ]);
  sheet.getRange(row, 3, 1, 3).merge().setValue('被害人（人次）');
  sheet.getRange(row, 6, 1, 3).merge().setValue('相對人（人次）');
  sheet.getRange(row, 9, 1, 3).merge().setValue('其他（人次）');
  sheet.getRange(row, 12, 2, 1).merge().setValue('聯繫會報\n（場次）');
  sheet.getRange(row, 13, 2, 1).merge().setValue('督導會議\n（場次）');
  sheet.getRange(row, 14, 2, 1).merge().setValue('宣導活動\n（場次）');
  
  let currentVolRow = row + 2;
  const volItems = ['社會福利服務', '陪同出庭', '擔任程序監理人', '法律服務', '保護令撤回案件諮詢服務', '網絡聯繫', '其他', '合計'];
  
  volItems.forEach(item => {
    const d = sv[item] || { 合計: 0, 被害人: { 小計: 0, 女: 0, 男: 0 }, 相對人: { 小計: 0, 女: 0, 男: 0 }, 其他: { 小計: 0, 女: 0, 男: 0 } };
    sheet.getRange(currentVolRow, 1, 1, 14).setValues([[
      item,
      d.合計, d.被害人.小計, d.被害人.女, d.被害人.男,
      d.相對人.小計, d.相對人.女, d.相對人.男,
      d.其他.小計, d.其他.女, d.其他.男,
      '', '', ''
    ]]);
    currentVolRow++;
  });
  formatTable(row, 1, volItems.length + 2, 14);
  sheet.getRange(row + 2, 1, volItems.length, 1).setHorizontalAlignment('left'); // 項目內容置左
  
  // 自動調整欄寬
  sheet.autoResizeColumns(1, 14);
}



/**
 * 依篩選條件計算多選欄位數
 */
function sumMultiSelectCountsByFilter(records, fieldName, identity, gender) {
  const filtered = records.filter(r => {
    const 身分 = r['個案身分'] || '';
    return 身分.includes(identity) && r['個案性別'] === gender;
  });
  return sumMultiSelectCounts(filtered, fieldName);
}

/**
 * 建立工作表 4: {年份}_data
 */
function createYearDataSheet(sheet, records) {
  // 欄位標題
  const headers = ['服務日期', '個案姓名', '個案性別', '個案國籍', '個案來源', '個案身分', 
                   '暴力類型', '有無在案', '在案社工', '接受轉介項目', '轉案股別/單位', 
                   '服務社工', '服務方式', '服務主題', '陪同出庭', '法律服務', 
                   '會談處遇', '通報', '轉介', '聯繫', '服務項目', '建立時間'];
  
  // 確保有足夠的欄數
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sheet, 1, 1, 1, headers.length);
  
  // 資料
  if (records.length > 0) {
    const data = records.map(r => headers.map(h => {
      const val = r[h];
      if (val instanceof Date) {
        return Utilities.formatDate(val, 'Asia/Taipei', 'yyyy-MM-dd');
      }
      return val || '';
    }));
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
  
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 格式化表頭
 */
function formatHeader(sheet, startRow, startCol, numRows, numCols) {
  const range = sheet.getRange(startRow, startCol, numRows, numCols);
  range.setBackground('#1a5276')
       .setFontColor('white')
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setBorder(true, true, true, true, true, true);
}

/**
 * 建立轉介來源統計表
 */
function buildReferralStatsRows(sheet, startRow, records, year) {
  let row = startRow;
  
  // === 表頭 ===
  // 第一列
  sheet.getRange(row, 1).setValue('個案來源');
  sheet.getRange(row, 2).setValue('自行求助');
  sheet.getRange(row, 3).setValue('接受轉介');
  sheet.getRange(row, 6).setValue('自行求助-網絡諮詢');
  sheet.getRange(row, 7).setValue('小計');
  
  // 合併儲存格
  sheet.getRange(row, 3, 1, 3).merge().setHorizontalAlignment('center'); // 接受轉介合併 3 欄
  
  row++;
  
  // 第二列
  sheet.getRange(row, 1).setValue('項目');
  sheet.getRange(row, 3).setValue('縣府轉介');
  sheet.getRange(row, 4).setValue('民間單位轉介');
  sheet.getRange(row, 5).setValue('司法轉介');
  
  row++;
  
  // === 內容 ===
  const services = ['陪同出庭', '諮詢服務'];
  // 定義來源與關鍵字/值對應
  const sources = [
    { key: '自行求助', label: '自行求助' },
    { key: '接受轉介-社會處', label: '縣府轉介' },
    { key: '接受轉介-民間單位', label: '民間單位轉介' }, // 假設值
    { key: '接受轉介-司法系統', label: '司法轉介' },
    { key: '自行求助-網絡諮詢', label: '自行求助-網絡諮詢' }
  ];
  
  let grandTotals = [0, 0, 0, 0, 0, 0]; // 各欄總計
  
  services.forEach(service => {
    sheet.getRange(row, 1).setValue(service);
    
    let rowTotal = 0;
    
    sources.forEach((source, index) => {
      // 篩選符合條件的不重複人數
      const uniqueNames = new Set();
      records.forEach(r => {
        const date = r['服務日期'] instanceof Date ? r['服務日期'] : new Date(r['服務日期']);
        
        // 判斷來源
        let sourceMatch = false;
        const val = (r['個案來源'] || '').trim(); // 使用正確的鍵名 '個案來源' 並去空白
        
        if (source.key === '自行求助') {
          sourceMatch = val === '自行求助';
        } else if (source.key === '自行求助-網絡諮詢') {
          sourceMatch = val === '自行求助-網絡諮詢';
        } else if (source.key === '接受轉介-社會處') {
          sourceMatch = val === '接受轉介-社會處';
        } else if (source.key === '接受轉介-司法系統') {
          sourceMatch = val === '接受轉介-司法系統';
        } else if (source.key === '接受轉介-民間單位') {
           // 寬鬆判斷：只要是「接受轉介」開頭，且不是社會處或司法系統
           if (val.startsWith('接受轉介') && 
               !val.includes('社會處') && 
               !val.includes('司法系統')) {
             sourceMatch = true;
           }
        }
        
        if (date.getFullYear() === year && 
            r['服務項目'] === service &&
            sourceMatch &&
            r['個案姓名']) {
          uniqueNames.add(r['個案姓名'].trim());
        }
      });
      
      const count = uniqueNames.size;
      sheet.getRange(row, index + 2).setValue(count);
      
      rowTotal += count;
      grandTotals[index] += count;
    });
    
    sheet.getRange(row, 7).setValue(rowTotal);
    grandTotals[5] += rowTotal;
    
    row++;
  });
  
  // === 總計列 ===
  sheet.getRange(row, 1).setValue('總計');
  grandTotals.forEach((total, index) => {
    sheet.getRange(row, index + 2).setValue(total);
  });
  
  // === 設定特定樣式 ===
  // 框線與置中
  const tableRange = sheet.getRange(startRow, 1, row - startRow + 1, 7);
  tableRange.setBorder(true, true, true, true, true, true);
  tableRange.setHorizontalAlignment('center').setVerticalAlignment('middle');
  // 項目靠左
  sheet.getRange(startRow + 1, 1, row - startRow, 1).setHorizontalAlignment('left');
  
  // 表頭背景色 (第一行和第二行)
  sheet.getRange(startRow, 1, 2, 7).setBackground('#f0f0f0').setFontWeight('bold'); // 淺灰背景
  
  return row + 2; // 留白
}

/**
 * 生成「參、保護令非撤回案件服務」的資料列
 */
function generateIdentityGenderRowsV2(data) {
  const identities = ['聲請人', '相對人', '關係人', '民眾'];
  const genders = ['女', '男'];
  
  let rows = '';
  
  // 1. 各身分資料列
  identities.forEach((identity) => {
    // 每個身分有兩列 (女, 男)
    // 第一列: 包含身分 rowspan
    const dateFem = data[identity]['女'];
    const dateMale = data[identity]['男'];
    
    // 會談服務 = 新版(會談服務) + 舊版(會談處遇)
    const meetingFem = (dateFem.會談服務 || 0) + (dateFem.會談處遇 || 0);
    const meetingMale = (dateMale.會談服務 || 0) + (dateMale.會談處遇 || 0);

    // 女:
    rows += `<tr>
      <td rowspan="2" style="vertical-align: middle;">${identity}</td>
      <td>女(${dateFem.服務人數} 人次)</td>
      <td>${dateFem.陪同出庭}</td>
      <td>${dateFem.人身安全}</td>
      <td>${dateFem.法律服務}</td>
      <td>${meetingFem}</td>
      <td>${dateFem.通報}</td>
      <td>${dateFem.轉介}</td>
      <td>${dateFem.聯繫}</td>
      <td>${dateFem.總計}</td>
    </tr>`;
    
    // 男:
    rows += `<tr>
      <td>男(${dateMale.服務人數} 人次)</td>
      <td>${dateMale.陪同出庭}</td>
      <td>${dateMale.人身安全}</td>
      <td>${dateMale.法律服務}</td>
      <td>${meetingMale}</td>
      <td>${dateMale.通報}</td>
      <td>${dateMale.轉介}</td>
      <td>${dateMale.聯繫}</td>
      <td>${dateMale.總計}</td>
    </tr>`;
  });
  
  // 2. 總計列
  const totalFem = data['總計']['女'];
  const totalMale = data['總計']['男'];
  const totalMeetingFem = (totalFem.會談服務 || 0) + (totalFem.會談處遇 || 0);
  const totalMeetingMale = (totalMale.會談服務 || 0) + (totalMale.會談處遇 || 0);
  
  rows += `<tr>
      <td style="vertical-align: middle;">總 計</td>
      <td>女(${totalFem.服務人數} 人次)<br>男(${totalMale.服務人數} 人次)</td>
      <td>${totalFem.陪同出庭 + totalMale.陪同出庭}</td>
      <td>${totalFem.人身安全 + totalMale.人身安全}</td>
      <td>${totalFem.法律服務 + totalMale.法律服務}</td>
      <td>${totalMeetingFem + totalMeetingMale}</td>
      <td>${totalFem.通報 + totalMale.通報}</td>
      <td>${totalFem.轉介 + totalMale.轉介}</td>
      <td>${totalFem.聯繫 + totalMale.聯繫}</td>
      <td>${totalFem.總計 + totalMale.總計}</td>
    </tr>`;
    
  return rows;
}

/**
 * 計算服務案件類型 (用於第五張報表)
 */
function countServiceTypes(records) {
  const result = {
    暫時: 0, 通常: 0, 變更: 0, 延長: 0, 撤銷: 0, 撤回: 0, 追加: 0, 違反: 0, 其他: 0, 總計: 0
  };
  
  // 明確定義關鍵字與對應欄位
  const typeKeywords = [
    { key: '暫時', words: ['暫時'] },
    { key: '通常', words: ['通常'] },
    { key: '變更', words: ['變更'] },
    { key: '延長', words: ['延長'] },
    { key: '撤銷', words: ['撤銷'] },
    { key: '撤回', words: ['撤回'] },
    { key: '追加', words: ['追加'] },
    { key: '違反', words: ['違反'] }
  ];

  records.forEach(record => {
    const topic = record['服務主題'] || '';
    // 支援多種分隔符
    const items = topic.split(/[;#；]+/).map(s => s.trim()).filter(s => s);
    
    items.forEach(item => {
      let matched = false;
      for (const type of typeKeywords) {
        if (type.words.some(w => item.includes(w))) {
          result[type.key]++;
          matched = true;
          break;
        }
      }
      
      // 未匹配到的都算「其他」（包含 抗告, 傷害, 跟騷...）
      if (!matched && item) {
        result.其他++;
      }
    });
  });
  
  // 計算總計
  result.總計 = 
    result.暫時 + result.通常 + result.變更 + result.延長 + 
    result.撤銷 + result.撤回 + result.追加 + result.違反 + result.其他;
    
  return result;
}
