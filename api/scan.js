
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({
      ok: false,
      message: 'Method Not Allowed'
    });
  }

  try {
    const body = typeof req.body === 'string'
      ? JSON.parse(req.body || '{}')
      : (req.body || {});

    switch (body.action) {
      case 'syncTodayData':
        return await syncTodayData(res);

      case 'getBarcodeMap':
        return await getBarcodeMapFromSupabase(res);

      case 'lookupTracking':
      case 'lookupTrackingNumber':
        return await lookupTrackingFromSupabase(res, body);

      case 'learnProduct':
      case 'addBarcodeToExistingProduct':
        return await learnAndSyncBarcode(res, body);

      default:
        return await forwardToAppsScript(res, body);
    }
  } catch (error) {
    return res.status(502).json({
      ok: false,
      message: error.message || String(error)
    });
  }
}

// ==========================================
// 共用設定
// ==========================================

function supabaseConfig() {
  return {
    url: (
      process.env.SUPABASE_URL ||
      'https://kaiduxzwmqsfmjxsfwpp.supabase.co'
    ).replace(/\/$/, ''),
    key: process.env.SUPABASE_SERVICE_ROLE_KEY
  };
}

function getToday() {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: 'Asia/Taipei',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(new Date());

  const part = type =>
    parts.find(p => p.type === type)?.value || '';

  return (
    part('year') + '-' +
    part('month') + '-' +
    part('day')
  );
}

function normalizeTracking(value) {
  return String(value ?? '')
    .replace(/^'/, '')
    .replace(/\s+/g, '')
    .toUpperCase();
}

function sbHeaders(key, write = false) {
  return {
    apikey: key,
    Authorization: 'Bearer ' + key,
    ...(write ? {
      'Content-Type': 'application/json',
      Prefer: 'resolution=merge-duplicates,return=minimal'
    } : {})
  };
}

async function callAppsScript(payload) {
  const url = process.env.APPS_SCRIPT_URL;

  if (!url) {
    throw new Error('尚未設定 APPS_SCRIPT_URL');
  }

  const token = process.env.APPS_SCRIPT_API_TOKEN;

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(
      token
        ? { ...payload, apiToken: token }
        : payload
    )
  });

  const raw = await response.text();

  let result;

  try {
    result = JSON.parse(raw);
  } catch {
    throw new Error(
      'Apps Script 回應不是有效 JSON，HTTP ' +
      response.status +
      '。請檢查 Web App 部署與存取權限'
    );
  }

  if (!response.ok) {
    throw new Error(
      result.message ||
      'Apps Script HTTP ' + response.status
    );
  }

  return result;
}

async function forwardToAppsScript(res, body) {
  const result = await callAppsScript(body);
  return res.status(200).json(result);
}

// ==========================================
// 查詢 Supabase 託運單
// ==========================================

async function lookupTrackingFromSupabase(res, body) {
  const startedAt = Date.now();

  const trackingNumber = normalizeTracking(
    body.trackingNumber
  );

  // 黑貓12碼、蝦皮15碼數字、
  // TW開頭15碼英數混合
  const valid =
    /^(?:\d{12}|\d{15}|TW[A-Z0-9]{13})$/.test(
      trackingNumber
    );

  if (!valid) {
    return res.status(200).json({
      ok: false,
      message: '託運單格式不正確'
    });
  }

  const { url, key } = supabaseConfig();

  if (!key) {
    return res.status(500).json({
      ok: false,
      message: '尚未設定 SUPABASE_SERVICE_ROLE_KEY'
    });
  }

  const today = getToday();

  // 同時相容有、沒有前置單引號的資料
  // 使用單次 Supabase 查詢
  const candidates = [
    trackingNumber,
    "'" + trackingNumber
  ];

  const inValue = 'in.(' +
    candidates.map(v =>
      '"' + v.replace(/"/g, '\\"') + '"'
    ).join(',') +
    ')';

  const params = new URLSearchParams({
    select:
      'source,source_date,source_row,' +
      'tracking_number,raw_product_text,' +
      'product_code,shipping_name,required_qty',
    source_date: 'eq.' + today,
    tracking_number: inValue,
    order: 'source_row.asc',
    limit: '1000'
  });

  const response = await fetch(
    url +
    '/rest/v1/warehouse_shipping_source_rows?' +
    params,
    {
      headers: sbHeaders(key)
    }
  );

  if (!response.ok) {
    const detail = (await response.text()).slice(0, 300);

    return res.status(502).json({
      ok: false,
      message:
        'Supabase 查詢失敗 HTTP ' +
        response.status + '：' + detail
    });
  }

  const rows = await response.json();

  if (!rows.length) {
    return res.status(200).json({
      ok: false,
      message: '今天找不到這張託運單，請先更新今日託運單',
      lookupMs: Date.now() - startedAt
    });
  }

  const sources = [
    ...new Set(rows.map(row => row.source))
  ];

  if (sources.length !== 1) {
    return res.status(200).json({
      ok: false,
      message: '相同託運單存在不同出貨來源，請檢查資料'
    });
  }

  let rawProductText = '';
  let items = [];

  // 樵玟出貨
  if (sources[0] === 'qiaowen') {
    const uniqueTexts = [
      ...new Set(
        rows
          .map(row =>
            String(row.raw_product_text || '').trim()
          )
          .filter(Boolean)
      )
    ];

    rawProductText = uniqueTexts.join('+');

    if (!rawProductText) {
      return res.status(200).json({
        ok: false,
        message: '找到託運單，但商品內容為空白'
      });
    }
  }

  // 蝦皮出貨
  else if (sources[0] === 'shopee') {
    const quantities = new Map();

    for (const row of rows) {
      const name = String(
        row.shipping_name || ''
      ).trim();

      const qty = Number(row.required_qty);

      if (
        !name ||
        !Number.isInteger(qty) ||
        qty <= 0
      ) {
        return res.status(200).json({
          ok: false,
          message:
            '蝦皮商品名稱或數量無效，來源列：' +
            row.source_row
        });
      }

      quantities.set(
        name,
        (quantities.get(name) || 0) + qty
      );
    }

    items = [...quantities].map(
      ([productName, qty]) => ({
        productName,
        qty
      })
    );

    rawProductText = items
      .map(item =>
        item.qty + '*' + item.productName
      )
      .join('+');
  }

  else {
    return res.status(200).json({
      ok: false,
      message: '不支援的出貨來源：' + sources[0]
    });
  }

  return res.status(200).json({
    ok: true,
    trackingNumber,
    rawProductText,
    items,
    ignoredItems: [],
    source: sources[0],
    row: rows[0].source_row,
    lookupMs: Date.now() - startedAt
  });
}

// ==========================================
// 直接從 Supabase 讀取商品條碼
// ==========================================

async function getBarcodeMapFromSupabase(res) {
  const { url, key } = supabaseConfig();

  if (!key) {
    return res.status(500).json({
      ok: false,
      message: '尚未設定 SUPABASE_SERVICE_ROLE_KEY'
    });
  }

  const rows = [];

  // 分頁讀取，避免 Supabase 預設1000筆限制
  for (let offset = 0; ; offset += 1000) {
    const params = new URLSearchParams({
      select:
        'source_row,product_code,shipping_name,' +
        'barcode,full_name,note,learning_source,' +
        'source_created_at',
      order: 'source_row.asc',
      limit: '1000',
      offset: String(offset)
    });

    const response = await fetch(
      url +
      '/rest/v1/warehouse_product_barcodes?' +
      params,
      {
        headers: sbHeaders(key)
      }
    );

    if (!response.ok) {
      throw new Error(
        '讀取商品條碼失敗 HTTP ' + response.status
      );
    }

    const page = await response.json();
    rows.push(...page);

    if (page.length < 1000) break;
  }

  const products = rows
    .filter(row => row.shipping_name)
    .map(row => ({
      row: row.source_row,
      productCode: row.product_code || '',
      shippingName: row.shipping_name,
      barcode: row.barcode || '',
      fullName: row.full_name || '',
      note: row.note || '',
      learningSource: row.learning_source || '',
      createdAt: row.source_created_at || ''
    }));

  return res.status(200).json({
    ok: true,
    products
  });
}

// ==========================================
// 商品學習：先 Google Sheets，再 Supabase
// ==========================================

async function learnAndSyncBarcode(res, body) {
  const { url, key } = supabaseConfig();

  if (!key) {
    return res.status(500).json({
      ok: false,
      message: '尚未設定 SUPABASE_SERVICE_ROLE_KEY'
    });
  }

  // Google Sheets 是商品建檔的主要來源
  const result = await callAppsScript(body);

  if (!result.ok) {
    return res.status(200).json(result);
  }

  // 取得 Google Sheets 更新後的條碼資料
  const barcodeMap = await callAppsScript({
    action: 'getBarcodeMap'
  });

  if (!barcodeMap.ok) {
    return res.status(502).json({
      ok: false,
      message: barcodeMap.message || '商品對照讀取失敗',
      sheetsUpdated: true
    });
  }

  const barcode = String(
    body.barcode || ''
  ).replace(/\s+/g, '');

  const name = String(
    body.shippingName || ''
  ).trim();

  const matches = (barcodeMap.products || [])
    .filter(p =>
      String(p.barcode || '').replace(/\s+/g, '') === barcode &&
      String(p.shippingName || '').trim() === name
    );

  if (!matches.length) {
    return res.status(502).json({
      ok: false,
      message:
        '商品已寫入 Google Sheets，但無法確認新條碼，請重新同步',
      sheetsUpdated: true
    });
  }

  const records = matches.map(p => ({
    source_row: p.row,
    product_code: p.productCode || null,
    shipping_name: p.shippingName,
    barcode: p.barcode || null,
    full_name: p.fullName || null,
    note: p.note || null,
    learning_source: p.learningSource || null,
    source_created_at: p.createdAt || null,
    synced_at: new Date().toISOString()
  }));

  const response = await fetch(
    url +
    '/rest/v1/warehouse_product_barcodes' +
    '?on_conflict=source_row',
    {
      method: 'POST',
      headers: sbHeaders(key, true),
      body: JSON.stringify(records)
    }
  );

  if (!response.ok) {
    return res.status(502).json({
      ok: false,
      message:
        'Google Sheets 已寫入，但 Supabase 同步失敗，HTTP ' +
        response.status,
      sheetsUpdated: true
    });
  }

  return res.status(200).json({
    ...result,
    supabaseSynced: true
  });
}

// ==========================================
// Google Sheets 同步至 Supabase
// ==========================================

async function syncTodayData(res) {
  const { url, key } = supabaseConfig();

  if (!key) {
    return res.status(500).json({
      ok: false,
      message: '尚未設定 SUPABASE_SERVICE_ROLE_KEY'
    });
  }

  const sourceData = await callAppsScript({
    action: 'exportWarehouseSyncData'
  });

  if (!sourceData.ok) {
    return res.status(502).json({
      ok: false,
      message:
        sourceData.message ||
        '讀取 Google Sheets 失敗'
    });
  }

  const headers = sbHeaders(key, true);

  async function upsert(table, rows, conflict) {
    if (!rows.length) return 0;

    let count = 0;

    for (let i = 0; i < rows.length; i += 300) {
      const batch = rows.slice(i, i + 300)
        .map(row => ({
          ...row,
          synced_at: new Date().toISOString()
        }));

      const requestUrl =
        url +
        '/rest/v1/' +
        table +
        '?on_conflict=' +
        encodeURIComponent(conflict);

      const response = await fetch(requestUrl, {
        method: 'POST',
        headers,
        body: JSON.stringify(batch)
      });

      if (!response.ok) {
        throw new Error(
          'Supabase 寫入 ' + table +
          ' 失敗 HTTP ' + response.status +
          '：' + (await response.text()).slice(0, 300)
        );
      }

      count += batch.length;
    }

    return count;
  }

  const barcodeRows = (
    sourceData.barcodeProducts || []
  ).map(p => ({
    product_code: p.product_code || null,
    shipping_name: p.shipping_name,
    barcode: p.barcode || null,
    full_name: p.full_name || null,
    note: p.note || null,
    learning_source: p.learning_source || null,
    source_created_at: p.source_created_at || null,
    source_row: p.source_row
  }));

  const shippingRows = (
    sourceData.shippingRows || []
  ).map(row => ({
    ...row,
    raw_product_text: row.raw_product_text || null,
    product_code: row.product_code || null,
    shipping_name: row.shipping_name || null
  }));

  const barcodeCount = await upsert(
    'warehouse_product_barcodes',
    barcodeRows,
    'source_row'
  );

  const shippingCount = await upsert(
    'warehouse_shipping_source_rows',
    shippingRows,
    'source,source_date,source_row'
  );

  return res.status(200).json({
    ok: true,
    message: '今日出貨與商品資料同步完成',
    source_date: sourceData.source_date,
    counts: {
      barcodeProducts: barcodeCount,
      shippingRows: shippingCount,
      qiaowenRows:
        sourceData.counts?.qiaowenRows || 0,
      shopeeRows:
        sourceData.counts?.shopeeRows || 0
    }
  });
}
