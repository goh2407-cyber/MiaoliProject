// ================================================================
// MeetingService.gs - 會面紀錄 CRUD（家事會面服務）
// ================================================================

// ==================== 欄位定義 ====================

const MEETING_RECORD_FIELDS = [
  'RecordId', '服務日期', '會面分案案號', '個案姓名', '指定性別', '服務方式',
  '個案身分', '司法案號', '在案與否', '在案單位社工', '國籍別', '轉介單位',
  '服務社工', '服務主題',
  '交往交付', '專業諮詢-法律', '專業諮詢-社福', '會談服務',
  '聯繫', '轉介', '通報',
  '建立時間', '修改時間', '建立者', '修改者'
];

// ==================== 年度分表輔助函數 ====================

/**
 * 根據年份取得或建立對應的會面紀錄 Sheet
 */
function getMeetingSheetByYear(year) {
  if (!spreadsheet) initializeSpreadsheet();
  
  const sheetName = '會面紀錄_' + year;
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, MEETING_RECORD_FIELDS.length).setValues([MEETING_RECORD_FIELDS]);
    
    const headerRange = sheet.getRange(1, 1, 1, MEETING_RECORD_FIELDS.length);
    headerRange.setBackground('#1A1A1A');
    headerRange.setFontColor('#FAFAF8');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    console.log('已建立會面年度 Sheet:', sheetName);
  }
  
  return sheet;
}

// ==================== 新增紀錄 ====================

/**
 * 新增會面紀錄
 */
function createMeetingRecord(recordData) {
  return executeWithRetry(function() {
    checkUserPermission();
    const user = getCurrentUser();
    
    validateMeetingRecordData(recordData);
    
    const recordId = Utilities.getUuid();
    
    const serviceDate = recordData['服務日期'] || '';
    const year = getYearFromDate(serviceDate);
    const sheet = getMeetingSheetByYear(year);
    
    const now = new Date();
    const rowData = [
      recordId,
      recordData['服務日期'] || '',
      recordData['會面分案案號'] || '',
      recordData['個案姓名'] || '',
      recordData['指定性別'] || '',
      recordData['服務方式'] || '',
      recordData['個案身分'] || '',
      recordData['司法案號'] || '',
      recordData['在案與否'] ? '是' : '否',
      recordData['在案單位社工'] || '',
      recordData['國籍別'] || '',
      recordData['轉介單位'] || '',
      recordData['服務社工'] || '',
      recordData['服務主題'] || '',
      formatMultiChoice(recordData['交往交付']),
      formatMultiChoice(recordData['專業諮詢-法律']),
      formatMultiChoice(recordData['專業諮詢-社福']),
      formatMultiChoice(recordData['會談服務']),
      formatMultiChoice(recordData['聯繫']),
      formatMultiChoice(recordData['轉介']),
      formatMultiChoice(recordData['通報']),
      Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),
      Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),
      user.email,
      user.email
    ];
    
    sheet.appendRow(rowData);
    updateLastModifiedTime('會面紀錄_' + year);
    clearMeetingRecordsCache();
    
    console.log('✅ 新增會面紀錄成功，RecordId:', recordId);
    return { success: true, recordId: recordId, message: '新增成功' };
  });
}

// ==================== 更新紀錄 ====================

/**
 * 更新會面紀錄
 */
function updateMeetingRecord(recordId, updateData) {
  return executeWithRetry(function() {
    checkUserPermission();
    const user = getCurrentUser();
    
    const found = findMeetingRecordInAllYears(recordId);
    if (found.rowIndex === -1) {
      throw new Error('找不到該筆紀錄');
    }
    
    const sheet = found.sheet;
    const rowIndex = found.rowIndex;
    
    const existingRow = sheet.getRange(rowIndex, 1, 1, MEETING_RECORD_FIELDS.length).getValues()[0];
    const now = new Date();
    
    const updatedRow = [
      recordId,
      updateData['服務日期'] !== undefined ? updateData['服務日期'] : existingRow[1],
      updateData['會面分案案號'] !== undefined ? updateData['會面分案案號'] : existingRow[2],
      updateData['個案姓名'] !== undefined ? updateData['個案姓名'] : existingRow[3],
      updateData['指定性別'] !== undefined ? updateData['指定性別'] : existingRow[4],
      updateData['服務方式'] !== undefined ? updateData['服務方式'] : existingRow[5],
      updateData['個案身分'] !== undefined ? updateData['個案身分'] : existingRow[6],
      updateData['司法案號'] !== undefined ? updateData['司法案號'] : existingRow[7],
      updateData['在案與否'] !== undefined ? (updateData['在案與否'] ? '是' : '否') : existingRow[8],
      updateData['在案單位社工'] !== undefined ? updateData['在案單位社工'] : existingRow[9],
      updateData['國籍別'] !== undefined ? updateData['國籍別'] : existingRow[10],
      updateData['轉介單位'] !== undefined ? updateData['轉介單位'] : existingRow[11],
      updateData['服務社工'] !== undefined ? updateData['服務社工'] : existingRow[12],
      updateData['服務主題'] !== undefined ? updateData['服務主題'] : existingRow[13],
      updateData['交往交付'] !== undefined ? formatMultiChoice(updateData['交往交付']) : existingRow[14],
      updateData['專業諮詢-法律'] !== undefined ? formatMultiChoice(updateData['專業諮詢-法律']) : existingRow[15],
      updateData['專業諮詢-社福'] !== undefined ? formatMultiChoice(updateData['專業諮詢-社福']) : existingRow[16],
      updateData['會談服務'] !== undefined ? formatMultiChoice(updateData['會談服務']) : existingRow[17],
      updateData['聯繫'] !== undefined ? formatMultiChoice(updateData['聯繫']) : existingRow[18],
      updateData['轉介'] !== undefined ? formatMultiChoice(updateData['轉介']) : existingRow[19],
      updateData['通報'] !== undefined ? formatMultiChoice(updateData['通報']) : existingRow[20],
      existingRow[21],
      Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),
      existingRow[23],
      user.email
    ];
    
    sheet.getRange(rowIndex, 1, 1, MEETING_RECORD_FIELDS.length).setValues([updatedRow]);
    updateLastModifiedTime(sheet.getName());
    clearMeetingRecordsCache();
    
    console.log('✅ 更新會面紀錄成功，RecordId:', recordId);
    return { success: true, message: '更新成功' };
  });
}

// ==================== 刪除紀錄 ====================

function deleteMeetingRecord(recordId) {
  return executeWithRetry(function() {
    checkUserPermission();
    
    const found = findMeetingRecordInAllYears(recordId);
    if (found.rowIndex === -1) {
      throw new Error('找不到該筆紀錄');
    }
    
    const sheet = found.sheet;
    sheet.deleteRow(found.rowIndex);
    
    updateLastModifiedTime(sheet.getName());
    clearMeetingRecordsCache();
    
    console.log('✅ 刪除會面紀錄成功');
    return { success: true, message: '刪除成功' };
  });
}

// ==================== 查詢與快取 ====================

function getMeetingRecords(filters, page, pageSize) {
  checkUserPermission();
  
  page = page || 1;
  pageSize = pageSize || 50;
  
  const cache = CacheService.getScriptCache();
  const filterKey = JSON.stringify(filters || {});
  const version = getMeetingRecordsCacheVersion();
  const metaCacheKey = `meeting_records_meta_v${version}_${filterKey}`;
  const cachedMeta = cache.get(metaCacheKey);
  
  let allRecords;
  
  if (cachedMeta) {
    const meta = JSON.parse(cachedMeta);
    allRecords = [];
    for (let i = 0; i < meta.totalChunks; i++) {
        const chunkData = cache.get(`meeting_records_chunk_v${version}_${i}_${filterKey}`);
        if (chunkData) allRecords = allRecords.concat(JSON.parse(chunkData));
        else { allRecords = null; break; }
    }
  }
  
  if (!allRecords) {
    if (!spreadsheet) initializeSpreadsheet();
    
    const targetSheets = [];
    if (filters && filters.year) {
      const sheet = spreadsheet.getSheetByName('會面紀錄_' + filters.year);
      if (sheet) targetSheets.push(sheet);
    } else {
      for (const sheet of spreadsheet.getSheets()) {
        const sheetName = sheet.getName();
        if (sheetName.startsWith('會面紀錄_') && !sheetName.includes('備份')) {
          targetSheets.push(sheet);
        }
      }
    }
    
    allRecords = [];
    for (const sheet of targetSheets) {
      if (sheet.getLastRow() <= 1) continue;
      
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEETING_RECORD_FIELDS.length).getDisplayValues();
      const sheetRecords = data.map(row => {
        const record = {};
        MEETING_RECORD_FIELDS.forEach((field, index) => {
          record[field] = row[index] !== null && row[index] !== undefined ? row[index] : '';
        });
        return record;
      });
      allRecords = allRecords.concat(sheetRecords);
    }
    
    const multiChoiceFields = ['交往交付', '專業諮詢-法律', '專業諮詢-社福', '會談服務', '聯繫', '轉介', '通報'];
    allRecords = allRecords.map(r => {
      multiChoiceFields.forEach(f => {
        const val = r[f];
        r[f] = (val && typeof val === 'string' && val.trim() !== '') ? val.split(';').map(s => s.trim()).filter(Boolean) : [];
      });
      r['在案與否'] = r['在案與否'] === '是' || r['在案與否'] === true;
      return r;
    });
    
    if (filters) {
      allRecords = applyMeetingFilters(allRecords, filters);
    }
    
    allRecords.sort((a, b) => {
        const d1 = String(a['服務日期'] || '');
        const d2 = String(b['服務日期'] || '');
        return d2.localeCompare(d1);
    });
    
    try {
      const totalChunks = Math.ceil(allRecords.length / CACHE_CHUNK_SIZE);
      for (let i = 0; i < totalChunks; i++) {
        const chunk = allRecords.slice(i * CACHE_CHUNK_SIZE, (i + 1) * CACHE_CHUNK_SIZE);
        cache.put(`meeting_records_chunk_v${version}_${i}_${filterKey}`, JSON.stringify(chunk), CACHE_EXPIRATION);
      }
      cache.put(metaCacheKey, JSON.stringify({ totalChunks, totalCount: allRecords.length, version }), CACHE_EXPIRATION);
    } catch (e) {
      console.log('分塊快取儲存失敗:', e.message);
    }
  }
  
  const startIndex = (page - 1) * pageSize;
  const pageRecords = allRecords.slice(startIndex, startIndex + pageSize);
  
  return { records: pageRecords, totalCount: allRecords.length, page, pageSize };
}

function getMeetingRecordsCacheVersion() {
  const cache = CacheService.getScriptCache();
  let version = cache.get('meeting_records_cache_version');
  if (!version) {
    version = PropertiesService.getScriptProperties().getProperty('meeting_records_cache_version') || '1';
    cache.put('meeting_records_cache_version', version, 21600);
  }
  return version;
}

function clearMeetingRecordsCache() {
  const version = new Date().getTime().toString();
  PropertiesService.getScriptProperties().setProperty('meeting_records_cache_version', version);
  CacheService.getScriptCache().put('meeting_records_cache_version', version, 21600);
}

// ==================== 輔助函數 ====================

function findMeetingRecordInAllYears(recordId) {
  if (!spreadsheet) initializeSpreadsheet();
  const searchId = String(recordId).trim();
  for (const sheet of spreadsheet.getSheets()) {
    const name = sheet.getName();
    if (!name.startsWith('會面紀錄_') || name.includes('備份') || sheet.getLastRow() <= 1) continue;
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchId) return { sheet, rowIndex: i + 2 };
    }
  }
  return { sheet: null, rowIndex: -1 };
}

function applyMeetingFilters(records, filters) {
  return records.filter(record => {
    if (filters.RecordId && record.RecordId !== filters.RecordId) return false;
    if (filters.year && getYearFromDate(record['服務日期']) !== filters.year) return false;
    if (filters.month) {
      const m = String(record['服務日期']).length >= 7 ? parseInt(String(record['服務日期']).substring(5, 7)) : 0;
      if (m !== filters.month) return false;
    }
    if (filters.startDate && record['服務日期'] < filters.startDate) return false;
    if (filters.endDate && record['服務日期'] > filters.endDate) return false;
    if (filters['個案姓名'] && !record['個案姓名'].includes(filters['個案姓名'])) return false;
    
    const exact = ['個案身分', '國籍別', '指定性別', '服務社工', '服務方式', '服務主題'];
    for (const field of exact) {
      if (filters[field] && record[field] !== filters[field]) return false;
    }
    if (filters['在案與否'] !== undefined && record['在案與否'] !== filters['在案與否']) return false;
    return true;
  });
}

function validateMeetingRecordData(recordData) {
  const required = ['服務日期', '個案姓名', '個案身分', '國籍別', '指定性別', '服務社工', '服務方式', '服務主題'];
  const missing = required.filter(f => !recordData[f]);
  if (missing.length > 0) throw new Error('缺少必填欄位：' + missing.join('、'));
}
