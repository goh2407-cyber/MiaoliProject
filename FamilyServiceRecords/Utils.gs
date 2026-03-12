// ================================================================
// Utils.gs - 工具函數
// ================================================================

/**
 * 格式化日期為 YYYY-MM-DD
 * @param {Date} date - 日期
 * @returns {string} 格式化後的日期字串
 */
function formatDate(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/**
 * 格式化日期時間為 YYYY-MM-DD HH:mm
 * @param {Date} date - 日期
 * @returns {string} 格式化後的日期時間字串
 */
function formatDateTime(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  const h = String(date.getHours()).padStart(2, '0');
  const min = String(date.getMinutes()).padStart(2, '0');
  return `${y}-${m}-${d} ${h}:${min}`;
}

/**
 * 解析日期字串為 Date 物件
 * @param {string} dateStr - 日期字串
 * @returns {Date|null} Date 物件或 null
 */
function parseDate(dateStr) {
  if (!dateStr) return null;
  const date = new Date(dateStr);
  return isNaN(date.getTime()) ? null : date;
}

/**
 * 取得今天日期（不含時間）
 * @returns {Date} 今天 00:00:00
 */
function getToday() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return today;
}

/**
 * 取得本月第一天
 * @returns {Date} 本月第一天
 */
function getFirstDayOfMonth() {
  const date = new Date();
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

/**
 * 取得本月最後一天
 * @returns {Date} 本月最後一天
 */
function getLastDayOfMonth() {
  const date = new Date();
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

/**
 * 深度複製物件
 * @param {Object} obj - 要複製的物件
 * @returns {Object} 複製後的物件
 */
function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

/**
 * 檢查是否為空值
 * @param {*} value - 值
 * @returns {boolean} 是否為空
 */
function isEmpty(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === 'string') return value.trim() === '';
  if (Array.isArray(value)) return value.length === 0;
  if (typeof value === 'object') return Object.keys(value).length === 0;
  return false;
}

/**
 * 移除字串前後空白
 * @param {string} str - 字串
 * @returns {string} 處理後的字串
 */
function trim(str) {
  return str ? String(str).trim() : '';
}

/**
 * 記錄錯誤訊息
 * @param {string} functionName - 函數名稱
 * @param {Error} error - 錯誤物件
 */
function logError(functionName, error) {
  console.error(`[${functionName}] 錯誤:`, error.message);
  console.error('堆疊:', error.stack);
}
