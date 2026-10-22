const SPREADSHEET_ID = '1wtGfufEO5UFS8hv9NdzP3t3YkmGVr12IINryQY2TMjM';

const ORDER_SHEET_NAME = '出貨原始來源';
const BARCODE_SHEET_NAME = '商品條碼對照';
const LOG_SHEET_NAME = '掃碼紀錄';


/* =========================================================
   Web App
========================================================= */

function doGet() {
  return jsonResponse_({
    ok: true,
    service: 'warehouse-scanner-api',
    message: 'API 已啟動'
  });
}


/*
  Vercel → Apps Script API
  Web App 必須部署為「以我執行」。
*/
function doPost(e) {

  try {

    const payload =
      parseApiPayload_(e);

    const expectedToken =
      PropertiesService
        .getScriptProperties()
        .getProperty('API_TOKEN');

    if (
      expectedToken &&
      payload.apiToken !== expectedToken
    ) {
      return jsonResponse_({
        ok: false,
        message: 'API Token 錯誤'
      });
    }

    const action =
      String(payload.action || '').trim();

    const result =
      dispatchApiAction_(
        action,
        payload
      );

    return jsonResponse_(
      result || {
        ok: false,
        message: '沒有回傳結果'
      }
    );

  }
  catch (error) {

    return jsonResponse_({
      ok: false,
      message:
        error.message || String(error)
    });

  }
}


function dispatchApiAction_(action, payload) {

  switch (action) {

    case 'lookupTracking':
    case 'lookupTrackingNumber':
      return lookupTrackingNumber(
        payload.trackingNumber
      );

    case 'getBarcodeMap':
      return getBarcodeMap();

    case 'exportWarehouseSyncData':
      return exportWarehouseSyncData_();

    case 'learnProduct':
      return learnProduct(
        payload.shippingName,
        payload.barcode
      );

    case 'addBarcodeToExistingProduct':
      return addBarcodeToExistingProduct(
        payload.shippingName,
        payload.barcode
      );

    case 'logScan':
      return logScan(
        payload.data || payload
      );

    case 'logBatchCompletion':
      return logBatchCompletion(
        payload.trackingNumbers || [],
        payload.rawProductTexts || [],
        payload.items || []
      );


case 'getWarehouseOperators':
  return getWarehouseOperators();

case 'getTodayScanHistory':
  return getTodayScanHistory();

case 'getScanHistoryByDate':
  return getScanHistoryByDate(payload.date);


    default:
      return {
        ok: false,
        message:
          '未知 API action：' + action
      };
  }
}


function parseApiPayload_(e) {

  if (
    !e ||
    !e.postData ||
    !e.postData.contents
  ) {
    return {};
  }

  const raw =
    e.postData.contents;

  try {
    return JSON.parse(raw);
  }
  catch (error) {
    throw new Error(
      'POST 內容不是有效 JSON'
    );
  }
}


function jsonResponse_(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}


/* =========================================================
   查詢託運單
   D欄 = 託運單號
   F欄 = 物品名稱
   A欄 = 當日資料日期

   本版優化：
   1. 先只讀 A 欄，找出今天的資料範圍
   2. 再只讀今天範圍內的 D:F
   3. 不再每次把歷史全部 D:F 讀進來
========================================================= */

function lookupTrackingNumber(trackingNumber) {

  const startedAt = Date.now();

  trackingNumber = normalizeTracking_(trackingNumber);

  if (!trackingNumber) {
    return {
      ok: false,
      message: '託運單號為空白'
    };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ORDER_SHEET_NAME);

  if (!sheet) {
    return {
      ok: false,
      message: '找不到「出貨原始來源」工作表'
    };
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      ok: false,
      message: '目前沒有出貨資料'
    };
  }

  /*
    A欄先取得當日範圍。

    使用 Spreadsheet 的時區判斷「今天」，
    避免伺服器時區與試算表日期不同造成誤判。
  */
  const timeZone =
    ss.getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    'Asia/Taipei';

  const todayKey =
    Utilities.formatDate(
      new Date(),
      timeZone,
      'yyyy-MM-dd'
    );

  const dateValues =
    sheet
      .getRange(2, 1, lastRow - 1, 1)
      .getValues();

  let firstTodayIndex = -1;
  let lastTodayIndex = -1;

  for (let i = 0; i < dateValues.length; i++) {

    const value = dateValues[i][0];

    if (!(value instanceof Date)) {
      continue;
    }

    const key =
      Utilities.formatDate(
        value,
        timeZone,
        'yyyy-MM-dd'
      );

    if (key === todayKey) {

      if (firstTodayIndex === -1) {
        firstTodayIndex = i;
      }

      lastTodayIndex = i;
    }
  }

  if (firstTodayIndex === -1) {
    return {
      ok: false,
      message: '今天沒有出貨資料',
      lookupMs: Date.now() - startedAt
    };
  }

  /*
    A欄第2列對應 index 0。
    所以實際 Sheet row：
      firstTodayIndex + 2
      lastTodayIndex + 2

    這裡讀取的只有「今天資料範圍」的 D:F。
  */
  const startRow =
    firstTodayIndex + 2;

  const numRows =
    lastTodayIndex - firstTodayIndex + 1;

  const values =
    sheet
      .getRange(
        startRow,
        4,
        numRows,
        3
      )
      .getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {

    const rowTracking =
      normalizeTracking_(values[i][0]);

    if (rowTracking !== trackingNumber) {
      continue;
    }

    const rawProductText =
      normalizeText_(values[i][2]);

    if (!rawProductText) {
      return {
        ok: false,
        message: '找到託運單，但 F 欄沒有物品名稱',
        lookupMs: Date.now() - startedAt
      };
    }

    const parsed =
      parseOrderItems_(rawProductText);

    return {
      ok: true,
      trackingNumber: rowTracking,
      rawProductText: rawProductText,
      row: startRow + i,
      items: parsed.items,
      ignoredItems: parsed.ignoredItems,
      lookupMs: Date.now() - startedAt
    };
  }

  return {
    ok: false,
    message: '今天找不到這張託運單',
    lookupMs: Date.now() - startedAt
  };
}


/* =========================================================
   F欄商品解析
========================================================= */

function parseOrderItems_(rawText) {

  const text = normalizeText_(rawText);

  const result = {};
  const ignoredItems = [];

  if (!text) {
    return {
      items: [],
      ignoredItems: []
    };
  }

  const parts = text
    .split('+')
    .map(v => v.trim())
    .filter(Boolean);


  parts.forEach(originalPart => {

    let part = originalPart.trim();

    if (containsBox_(part)) {

      ignoredItems.push({
        raw: originalPart,
        reason: '箱裝'
      });

      return;
    }

    let qty = 1;
    let productName = part;

    const starMatch =
      part.match(/^(\d+)\s*[\*＊]\s*(.+)$/);


    if (starMatch) {

      qty = parseInt(
        starMatch[1],
        10
      );

      productName =
        starMatch[2].trim();
    }

    else {

      const xMatch =
        part.match(/^(.+?)\s*[xX×]\s*(\d+)$/);


      if (xMatch) {

        productName =
          xMatch[1].trim();

        qty =
          parseInt(
            xMatch[2],
            10
          );
      }
    }


    if (
      !productName ||
      !Number.isFinite(qty) ||
      qty <= 0
    ) {
      return;
    }

    productName =
      normalizeProductName_(productName);


    if (!result[productName]) {

      result[productName] = {
        productName: productName,
        qty: 0
      };
    }


    result[productName].qty += qty;

  });


  return {

    items:
      Object.keys(result)
        .map(key => result[key]),

    ignoredItems:
      ignoredItems
  };
}


/* =========================================================
   判斷箱裝
========================================================= */

function containsBox_(text) {

  const value =
    normalizeText_(text);

  return value.includes('箱');
}


/* =========================================================
   取得商品條碼對照
========================================================= */

function getBarcodeMap() {

  const ss =
    SpreadsheetApp
      .openById(SPREADSHEET_ID);

  const sheet =
    ss.getSheetByName(
      BARCODE_SHEET_NAME
    );


  if (!sheet) {

    return {
      ok: false,
      message:
        '找不到「商品條碼對照」工作表'
    };
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return {
      ok: true,
      products: []
    };
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        7
      )
      .getDisplayValues();


  const products = [];


  values.forEach(
    (row, index) => {

      const productCode =
        normalizeText_(row[0]);

      const shippingName =
        normalizeProductName_(row[1]);

      const barcode =
        normalizeBarcode_(row[2]);

      const fullName =
        normalizeText_(row[3]);

      const note =
        normalizeText_(row[4]);

      const learningSource =
        normalizeText_(row[5]);

      const createdAt =
        normalizeText_(row[6]);


      if (!shippingName) {
        return;
      }


      products.push({

        row: index + 2,

        productCode:
          productCode,

        shippingName:
          shippingName,

        barcode:
          barcode,

        fullName:
          fullName,

        note:
          note,

        learningSource:
          learningSource,

        createdAt:
          createdAt

      });

    }
  );


  return {
    ok: true,
    products: products
  };
}


/* =========================================================
   找出某個出貨名稱
========================================================= */

function findProductByShippingName(
  shippingName
) {

  shippingName =
    normalizeProductName_(
      shippingName
    );


  const data =
    getBarcodeMap();


  if (!data.ok) {
    return data;
  }


  const matches =
    data.products.filter(
      product =>
        normalizeProductName_(
          product.shippingName
        ) === shippingName
    );


  return {
    ok: true,
    found:
      matches.length > 0,
    products:
      matches
  };
}


/* =========================================================
   找商品條碼
========================================================= */

function findProductByBarcode(
  barcode
) {

  barcode =
    normalizeBarcode_(
      barcode
    );


  if (!barcode) {

    return {
      ok: false,
      message:
        '商品條碼為空白'
    };
  }


  const data =
    getBarcodeMap();


  if (!data.ok) {
    return data;
  }


  const matches =
    data.products.filter(
      product =>
        normalizeBarcode_(
          product.barcode
        ) === barcode
    );


  return {

    ok: true,

    found:
      matches.length > 0,

    products:
      matches
  };
}


/* =========================================================
   自我學習商品
========================================================= */

function learnProduct(
  shippingName,
  barcode
) {

  shippingName =
    normalizeProductName_(
      shippingName
    );

  barcode =
    normalizeBarcode_(
      barcode
    );


  if (!shippingName) {

    return {
      ok: false,
      message:
        '缺少出貨名稱'
    };
  }


  if (!barcode) {

    return {
      ok: false,
      message:
        '缺少商品條碼'
    };
  }


  const ss =
    SpreadsheetApp
      .openById(
        SPREADSHEET_ID
      );


  const sheet =
    ss.getSheetByName(
      BARCODE_SHEET_NAME
    );


  if (!sheet) {

    return {
      ok: false,
      message:
        '找不到商品條碼對照表'
    };
  }


  const barcodeResult =
    findProductByBarcode(
      barcode
    );


  if (
    barcodeResult.ok &&
    barcodeResult.found
  ) {

    const sameName =
      barcodeResult.products
        .some(
          product =>
            normalizeProductName_(
              product.shippingName
            ) === shippingName
        );


    if (sameName) {

      return {
        ok: true,
        alreadyExists: true,
        message:
          '此商品與條碼已存在'
      };
    }


    return {

      ok: false,

      conflict: true,

      message:
        '此條碼已綁定其他商品',

      existingProducts:
        barcodeResult.products
    };
  }


  const nameResult =
    findProductByShippingName(
      shippingName
    );


  if (
    nameResult.ok &&
    nameResult.found
  ) {

    const emptyBarcodeProduct =
      nameResult.products.find(
        product =>
          !normalizeBarcode_(
            product.barcode
          )
      );


    if (emptyBarcodeProduct) {

      const row =
        emptyBarcodeProduct.row;


      const lock =
        LockService.getScriptLock();

      lock.waitLock(10000);

      try {

        sheet
          .getRange(
            row,
            3
          )
          .setNumberFormat('@');

        sheet
          .getRange(
            row,
            3
          )
          .setValue(
            barcode
          );

        sheet
          .getRange(
            row,
            6
          )
          .setValue(
            '自動學習'
          );

        sheet
          .getRange(
            row,
            7
          )
          .setValue(
            new Date()
          );

      } finally {
        lock.releaseLock();
      }

      return {

        ok: true,

        updatedExistingRow: true,

        row: row,

        shippingName:
          shippingName,

        barcode:
          barcode,

        message:
          '已補上商品條碼'
      };
    }
  }


  const newRow =
    sheet.getLastRow() + 1;


  const lock2 =
    LockService.getScriptLock();

  lock2.waitLock(10000);

  try {

    sheet
      .getRange(
        newRow,
        1,
        1,
        7
      )
      .setValues([
        [
          '',
          shippingName,
          barcode,
          '',
          '',
          '自動學習',
          new Date()
        ]
      ]);

  } finally {
    lock2.releaseLock();
  }


  sheet
    .getRange(
      newRow,
      3
    )
    .setNumberFormat('@');


  return {

    ok: true,

    created: true,

    row:
      newRow,

    shippingName:
      shippingName,

    barcode:
      barcode,

    message:
      '已新增商品條碼對照'
  };
}


/* =========================================================
   增加第二個條碼
========================================================= */

function addBarcodeToExistingProduct(
  shippingName,
  barcode
) {

  shippingName =
    normalizeProductName_(
      shippingName
    );

  barcode =
    normalizeBarcode_(
      barcode
    );


  if (
    !shippingName ||
    !barcode
  ) {

    return {
      ok: false,
      message:
        '商品名稱或條碼缺失'
    };
  }


  const barcodeCheck =
    findProductByBarcode(
      barcode
    );


  if (
    barcodeCheck.ok &&
    barcodeCheck.found
  ) {

    const same =
      barcodeCheck.products.some(
        item =>
          normalizeProductName_(
            item.shippingName
          ) === shippingName
      );


    if (same) {

      return {
        ok: true,
        alreadyExists: true,
        message:
          '此條碼已存在'
      };
    }


    return {

      ok: false,

      conflict: true,

      message:
        '此條碼已屬於其他商品',

      existingProducts:
        barcodeCheck.products
    };
  }


  const ss =
    SpreadsheetApp
      .openById(
        SPREADSHEET_ID
      );


  const sheet =
    ss.getSheetByName(
      BARCODE_SHEET_NAME
    );


  const productCheck =
    findProductByShippingName(
      shippingName
    );


  let productCode = '';
  let fullName = '';


  if (
    productCheck.ok &&
    productCheck.found
  ) {

    productCode =
      productCheck.products[0]
        .productCode || '';

    fullName =
      productCheck.products[0]
        .fullName || '';
  }


  const newRow =
    sheet.getLastRow() + 1;


  sheet
    .getRange(
      newRow,
      1,
      1,
      7
    )
    .setValues([
      [
        productCode,
        shippingName,
        barcode,
        fullName,
        '同商品新增條碼',
        '自動學習',
        new Date()
      ]
    ]);


  sheet
    .getRange(
      newRow,
      3
    )
    .setNumberFormat('@');


  return {

    ok: true,

    created: true,

    row:
      newRow,

    message:
      '已新增第二組商品條碼'
  };
}


/* =========================================================
   寫入掃碼紀錄
========================================================= */

function logScan(data) {

  try {

    const ss =
      SpreadsheetApp
        .openById(
          SPREADSHEET_ID
        );


    const sheet =
      ss.getSheetByName(
        LOG_SHEET_NAME
      );


    if (!sheet) {

      return {
        ok: false,
        message:
          '找不到掃碼紀錄工作表'
      };
    }


    const row = [

      new Date(),

      normalizeTracking_(
        data.trackingNumber || ''
      ),

      normalizeText_(
        data.rawProductText || ''
      ),

      normalizeProductName_(
        data.shippingName || ''
      ),

      normalizeText_(
        data.productCode || ''
      ),

      Number(
        data.requiredQty || 0
      ),

      Number(
        data.scannedQty || 0
      ),

      normalizeText_(
        data.status || ''
      )

    ];


    sheet.appendRow(row);


    return {
      ok: true
    };

  }

  catch (error) {

    return {

      ok: false,

      message:
        error.message
    };
  }
}


/* =========================================================
   批次寫入完成紀錄
========================================================= */

function logBatchCompletion(
  trackingNumbers,
  rawProductTexts,
  items
) {

  try {

    const ss =
      SpreadsheetApp
        .openById(
          SPREADSHEET_ID
        );


    const sheet =
      ss.getSheetByName(
        LOG_SHEET_NAME
      );


    if (!sheet) {

      return {
        ok: false,
        message:
          '找不到掃碼紀錄工作表'
      };
    }


    const now =
      new Date();


    const trackingText =
      Array.isArray(
        trackingNumbers
      )
        ? trackingNumbers.join(',')
        : normalizeTracking_(
            trackingNumbers
          );


    const rawText =
      Array.isArray(
        rawProductTexts
      )
        ? rawProductTexts.join(' / ')
        : normalizeText_(
            rawProductTexts
          );


    const rows = [];


    items.forEach(
      item => {

        rows.push([

          now,

          trackingText,

          rawText,

          normalizeProductName_(
            item.shippingName ||
            item.productName
          ),

          normalizeText_(
            item.productCode || ''
          ),

          Number(
            item.requiredQty ||
            item.qty ||
            0
          ),

          Number(
            item.scannedQty ||
            item.requiredQty ||
            item.qty ||
            0
          ),

          '完成'

        ]);

      }
    );


    if (rows.length > 0) {

      const startRow =
        sheet.getLastRow() + 1;


      sheet
        .getRange(
          startRow,
          1,
          rows.length,
          8
        )
        .setValues(
          rows
        );
    }


    return {
      ok: true
    };

  }

  catch (error) {

    return {

      ok: false,

      message:
        error.message
    };
  }
}


/* =========================================================
   商品名稱標準化
========================================================= */

function normalizeProductName_(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }


  let text =
    String(value)
      .replace(/\r/g, '')
      .replace(/\n/g, '')
      .trim();


  text =
    text
      .replace(/（/g, '(')
      .replace(/）/g, ')')
      .replace(/\s+/g, ' ')
      .trim();


  text =
    text.replace(
      '(★單件)',
      '(★單品)'
    );


  return text;
}


/* =========================================================
   託運單號標準化
========================================================= */

function normalizeTracking_(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }


  return String(value)
    .replace(/^'/, '')
    .replace(/\s+/g, '')
    .trim();
}


/* =========================================================
   條碼標準化
========================================================= */

function normalizeBarcode_(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }


  return String(value)
    .replace(/\r/g, '')
    .replace(/\n/g, '')
    .replace(/\s+/g, '')
    .trim();
}


/* =========================================================
   一般文字標準化
========================================================= */

function normalizeText_(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }


  return String(value)
    .replace(/\r/g, '')
    .replace(/\n/g, '')
    .trim();
}


/* =========================================================
   測試解析器
========================================================= */

function testParser() {

  const tests = [

    '10*福06',

    '1(■箱)*道達13+3*道達13',

    '2*(4L)森02+1*森02+2*(4L)森03',

    '1【■箱】*福14+6*福17',

    '1【■箱】*殼14+2*殼14',

    '森A091(★單品)x6',

    '1*(4L)森04+森A091(★單品)',

    '1(■箱)*福06+2*福24'

  ];


  tests.forEach(
    text => {

      Logger.log(
        text
      );

      Logger.log(
        JSON.stringify(
          parseOrderItems_(text)
        )
      );

    }
  );
}


/* =========================================================
   匯出今日同步資料給 Vercel / Supabase
   Google Sheets 仍是原始資料來源
========================================================= */
function exportWarehouseSyncData_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const timeZone = ss.getSpreadsheetTimeZone() || 'Asia/Taipei';
  const today = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');

  function dateKey(value) {
    if (value instanceof Date) return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
    const text = String(value || '').trim();
    if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)) {
      const parts = text.split('-');
      return parts[0] + '-' + parts[1].padStart(2, '0') + '-' + parts[2].padStart(2, '0');
    }
    return '';
  }

  function sheetRows(name, startCol, width) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error('找不到工作表：' + name);
    const last = sheet.getLastRow();
    if (last < 2) return [];
    return sheet.getRange(2, startCol, last - 1, width).getDisplayValues();
  }

  const barcodeSheet = ss.getSheetByName(BARCODE_SHEET_NAME);
  if (!barcodeSheet) throw new Error('找不到商品條碼對照工作表');
  const barcodeLast = barcodeSheet.getLastRow();
  const barcodeProducts = [];
  if (barcodeLast >= 2) {
    const vals = barcodeSheet.getRange(2, 1, barcodeLast - 1, 7).getValues();
    vals.forEach((row, i) => {
      const shippingName = String(row[1] || '').trim();
      if (!shippingName) return;
      const createdAt = row[6] instanceof Date ? row[6].toISOString() : null;
      barcodeProducts.push({
        product_code: String(row[0] || '').trim(),
        shipping_name: shippingName,
        barcode: String(row[2] || '').trim(),
        full_name: String(row[3] || '').trim(),
        note: String(row[4] || '').trim(),
        learning_source: String(row[5] || '').trim(),
        source_created_at: createdAt,
        source_row: i + 2
      });
    });
  }

  const qiaowenSheet = ss.getSheetByName('樵玟出貨CSV檔(自動帶入)');
  if (!qiaowenSheet) throw new Error('找不到樵玟出貨CSV檔(自動帶入)工作表');
  const qiaowenRows = [];
  const qLast = qiaowenSheet.getLastRow();
  if (qLast >= 2) {
    const dates = qiaowenSheet.getRange(2, 1, qLast - 1, 1).getValues();
    const vals = qiaowenSheet.getRange(2, 4, qLast - 1, 3).getDisplayValues(); // D:F
    for (let i = 0; i < vals.length; i++) {
      if (dateKey(dates[i][0]) !== today) continue;
      const tracking = String(vals[i][0] || '').replace(/\s+/g, '').trim();
      const rawText = String(vals[i][2] || '').trim();
      if (!tracking || !rawText) continue;
      qiaowenRows.push({
        source: 'qiaowen',
        source_date: today,
        source_row: i + 2,
        tracking_number: tracking,
        raw_product_text: rawText,
        product_code: null,
        required_qty: 1,
        shipping_name: null
      });
    }
  }

  const shopeeSheet = ss.getSheetByName('蝦皮出貨檔(自動帶入)');
  if (!shopeeSheet) throw new Error('找不到蝦皮出貨檔(自動帶入)工作表');
  const shopeeRows = [];
  const sLast = shopeeSheet.getLastRow();
  if (sLast >= 2) {
    const dates = shopeeSheet.getRange(2, 1, sLast - 1, 1).getValues();
    const vals = shopeeSheet.getRange(2, 4, sLast - 1, 4).getDisplayValues(); // D:G
    const productMap = {};
    barcodeProducts.forEach(p => { if (p.product_code) productMap[p.product_code] = p.shipping_name; });
    for (let i = 0; i < vals.length; i++) {
      if (dateKey(dates[i][0]) !== today) continue;
      const code = String(vals[i][0] || '').trim();
      const qtyText = String(vals[i][2] || '').trim();
      const tracking = String(vals[i][3] || '').replace(/\s+/g, '').trim();
      if (!tracking || !code) continue;
      const qty = Number(qtyText || 1);
      if (!Number.isInteger(qty) || qty <= 0) {
        throw new Error('蝦皮第 ' + (i + 2) + ' 列商品數量無效：' + qtyText);
      }
      const shippingName = productMap[code];
      if (!shippingName) {
        throw new Error('蝦皮第 ' + (i + 2) + ' 列商品代碼無法對照：' + code);
      }
      shopeeRows.push({
        source: 'shopee',
        source_date: today,
        source_row: i + 2,
        tracking_number: tracking,
        raw_product_text: null,
        product_code: code,
        required_qty: qty,
        shipping_name: shippingName
      });
    }
  }

  return {
    ok: true,
    source_date: today,
    barcodeProducts: barcodeProducts,
    shippingRows: qiaowenRows.concat(shopeeRows),
    counts: {
      barcodeProducts: barcodeProducts.length,
      qiaowenRows: qiaowenRows.length,
      shopeeRows: shopeeRows.length
    }
  };
}


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
