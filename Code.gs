/* ===== 倉庫掃碼操作人員與指定日期歷史（貼到 Code.gs 最後方） ===== */

// Google Sheets「操作人員」A2 以下維護名單。
function getWarehouseOperators() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('操作人員');
  if (!sheet) return { ok:false, message:'找不到「操作人員」分頁' };
  const last = sheet.getLastRow();
  const operators = last < 2 ? [] : sheet.getRange(2, 1, last - 1, 1)
    .getDisplayValues()
    .map(r => String(r[0] || '').trim())
    .filter(Boolean);
  return {ok:true, operators:[...new Set(operators)]};
}

function getTodayScanHistory() {
  return getScanHistoryByDate();
}

function getScanHistoryByDate(requestedDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) return {ok:false, message:'找不到「掃碼紀錄」分頁'};

  const timeZone = ss.getSpreadsheetTimeZone() || 'Asia/Taipei';
  const today = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  const targetDate = requestedDate == null || requestedDate === ''
    ? today : String(requestedDate).trim();
  const match = targetDate.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return {ok:false, message:'日期格式錯誤，請使用 YYYY-MM-DD'};
  const test = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  if (!Number.isFinite(test.getTime()) || test.toISOString().slice(0,10) !== targetDate) {
    return {ok:false, message:'查詢日期無效'};
  }
  const last = sheet.getLastRow();
  // 先只讀 A 欄定位指定日期，減少歷史資料量大時的讀取時間。
  const dates = last >= 2 ? sheet.getRange(2, 1, last - 1, 1).getValues() : [];
  let firstRow = 0, lastRow = 0;
  dates.forEach((record, i) => {
    const at = record[0];
    if (at instanceof Date && Number.isFinite(at.getTime()) &&
        Utilities.formatDate(at, timeZone, 'yyyy-MM-dd') === targetDate) {
      const rowNum = i + 2;
      if (!firstRow) firstRow = rowNum;
      lastRow = rowNum;
    }
  });
  if (!firstRow) return {ok:true, sourceDate:targetDate, orders:[]};
  const rows = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, 9).getValues();
  const batches = new Map();

  function normalizeName(name) {
    return String(name || '').trim().replace(/（/g,'(').replace(/）/g,')');
  }
  function parseRequiredRaw(raw) {
    const result = new Map();
    String(raw || '').split(/\s+\/\s+/).forEach(group => {
      group.split(/[+＋]/).forEach(part => {
        part = part.trim();
        if (!part || part.includes('箱')) return;
        let name = part, qty = 1;
        const prefix = part.match(/^(\d+)\s*[\*＊]\s*(.+)$/);
        const suffix = part.match(/^(.+?)\s*[xX×]\s*(\d+)$/);
        if (prefix) { qty = Number(prefix[1]); name = prefix[2].trim(); }
        else if (suffix) { qty = Number(suffix[2]); name = suffix[1].trim(); }
        name = normalizeName(name);
        if (name && qty > 0) result.set(name, (result.get(name) || 0) + qty);
      });
    });
    return result;
  }

  rows.forEach(row => {
    const [at, tracking, rawText, productName, productCode, req, scanned, status, operator] = row;
    if (!(at instanceof Date) || !Number.isFinite(at.getTime())) return;
    if (Utilities.formatDate(at, timeZone, 'yyyy-MM-dd') !== targetDate) return;
    const number = String(tracking || '').trim().replace(/^'/, '');
    if (!number) return;
    const operatorName = String(operator || '').trim() || '未記錄';
    const groupKey = number + '\u0001' + operatorName;
    if (!batches.has(groupKey)) {
      batches.set(groupKey, {
        trackingNumber:number, operatorName, items:new Map(),
        completedAt:null, lastAt:at.getTime(), rawProductText:''
      });
    }
    const group = batches.get(groupKey);
    group.lastAt = Math.max(group.lastAt, at.getTime());
    group.rawProductText = group.rawProductText || String(rawText || '');
    if (String(status) === '完成') group.completedAt = at.toISOString();
    const name = normalizeName(productName);
    if (name) {
      if (!group.items.has(name)) group.items.set(name, {name, required:0, scanned:0});
      const item = group.items.get(name);
      item.required = Math.max(item.required, Number(req) || 0);
      item.scanned = Math.max(item.scanned, Number(scanned) || 0);
    }
  });

  const orders = [...batches.values()].map(group => {
    const needed = parseRequiredRaw(group.rawProductText);
    for (const [name, qty] of needed) {
      if (!group.items.has(name)) group.items.set(name, {name, required:qty, scanned:0});
      else {
        const item = group.items.get(name);
        item.required = Math.max(item.required, qty);
      }
    }
    return {
      trackingNumber:group.trackingNumber,
      operatorName:group.operatorName,
      completedAt:group.completedAt,
      items:[...group.items.values()],
      lastAt:group.lastAt
    };
  }).sort((a,b) => b.lastAt - a.lastAt);

  return {ok:true, sourceDate:targetDate, orders};
}
