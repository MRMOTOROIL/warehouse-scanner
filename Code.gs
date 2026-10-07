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
