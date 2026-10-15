
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({
      ok: false,
      message: 'Method Not Allowed'
    });
  }

  try {
    const body =
      typeof req.body === 'string'
        ? JSON.parse(req.body || '{}')
        : (req.body || {});

    if (body.action === 'syncTodayData') {
      return syncTodayData(req, res);
    }

    if (
      body.action === 'lookupTracking' ||
      body.action === 'lookupTrackingNumber'
    ) {
      return lookupTrackingFromSupabase(res, body);
    }

    const url = process.env.APPS_SCRIPT_URL;

    if (!url) {
      return res.status(500).json({
        ok: false,
        message: '尚未設定 APPS_SCRIPT_URL'
      });
    }

    const payload = { ...body };

    if (process.env.APPS_SCRIPT_API_TOKEN) {
      payload.apiToken =
        process.env.APPS_SCRIPT_API_TOKEN;
    }

    const upstream = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    const responseText = await upstream.text();

    let data;

    try {
      data = JSON.parse(responseText);
    } catch {
      data = {
        ok: false,
        message: 'Apps Script 回應不是有效 JSON'
      };
    }

    return res
      .status(upstream.ok ? 200 : upstream.status)
      .json(data);

  } catch (error) {
    return res.status(502).json({
      ok: false,
      message: error?.message || String(error)
    });
  }
}

function supabaseConfig() {
  return {
    url: (
      process.env.SUPABASE_URL ||
      'https://kaiduxzwmqsfmjxsfwpp.supabase.co'
    ).replace(/\/$/, ''),
    key: process.env.SUPABASE_SERVICE_ROLE_KEY
  };
}

/* ==========================================
   Supabase 查詢託運單
========================================== */

async function lookupTrackingFromSupabase(res, body) {
  const startedAt = Date.now();

  const trackingNumber = String(
    body.trackingNumber || ''
  )
    .replace(/^'/, '')
    .replace(/\s+/g, '')
    .toUpperCase();

  // 黑貓12碼、蝦皮15碼與TW英數格式
  if (
    !/^(?:\d{12}|\d{15}|TW[A-Z0-9]{13})$/.test(
      trackingNumber
    )
  ) {
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

  const today = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Taipei',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).format(new Date());

  const headers = {
    apikey: key,
    Authorization: 'Bearer ' + key
  };

  const rows = [];

  // 相容舊資料前面多一個單引號
  for (const candidate of [
    trackingNumber,
    "'" + trackingNumber
  ]) {
    const qs = new URLSearchParams({
      select:
        'source,source_date,source_row,tracking_number,' +
        'raw_product_text,product_code,shipping_name,required_qty',
      source_date: 'eq.' + today,
      tracking_number: 'eq.' + candidate,
      order: 'source_row.asc',
      limit: '1000'
    });

    const response = await fetch(
      url + '/rest/v1/warehouse_shipping_source_rows?' + qs,
      { headers }
    );

    if (!response.ok) {
      return res.status(502).json({
        ok: false,
        message:
          'Supabase 查詢失敗（HTTP ' +
          response.status +
          '）'
      });
    }

    rows.push(...await response.json());
  }

  if (!rows.length) {
    return res.status(200).json({
      ok: false,
      message:
        '今天找不到這張託運單（請先更新今日託運單）',
      lookupMs: Date.now() - startedAt
    });
  }

  const sources = [
    ...new Set(rows.map(row => row.source))
  ];

  if (sources.length !== 1) {
    return res.status(200).json({
      ok: false,
      message:
        '相同託運單同時出現在不同出貨來源，請檢查資料'
    });
  }

  let rawProductText = '';
  let items = [];

  // 樵玟出貨
  if (sources[0] === 'qiaowen') {
    const unique = [
      ...new Set(
        rows
          .map(row =>
            String(row.raw_product_text || '').trim()
          )
          .filter(Boolean)
      )
    ];

    rawProductText = unique.join('+');

    if (!rawProductText) {
      return res.status(200).json({
        ok: false,
        message:
          '找到託運單，但樵玟商品內容為空白'
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
            '蝦皮商品名稱或數量無效（來源列 ' +
            row.source_row +
            '）'
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
      message:
        '不支援的出貨來源：' + sources[0]
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

/* ==========================================
   Google Sheets 同步至 Supabase
========================================== */

async function syncTodayData(req, res) {
  const appsScriptUrl = process.env.APPS_SCRIPT_URL;
  const apiToken =
    process.env.APPS_SCRIPT_API_TOKEN || '';

  const {
    url: supabaseUrl,
    key: serviceKey
  } = supabaseConfig();

  if (!appsScriptUrl) {
    return res.status(500).json({
      ok: false,
      message: '尚未設定 APPS_SCRIPT_URL'
    });
  }

  if (!serviceKey) {
    return res.status(500).json({
      ok: false,
      message:
        '尚未設定 SUPABASE_SERVICE_ROLE_KEY'
    });
  }

  try {
    const sourcePayload = {
      action: 'exportWarehouseSyncData'
    };

    if (apiToken) {
      sourcePayload.apiToken = apiToken;
    }

    const sourceResponse = await fetch(
      appsScriptUrl,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(sourcePayload)
      }
    );

    const sourceText =
      await sourceResponse.text();

    let sourceData;

    try {
      sourceData = JSON.parse(sourceText);
    } catch {
      return res.status(502).json({
        ok: false,
        message: 'Apps Script 回應不是有效 JSON',
        detail: sourceText.slice(0, 300)
      });
    }

    if (!sourceResponse.ok || !sourceData.ok) {
      return res.status(502).json({
        ok: false,
        message:
          sourceData.message ||
          '讀取 Google Sheets 同步資料失敗'
      });
    }

    const headers = {
      apikey: serviceKey,
      Authorization: 'Bearer ' + serviceKey,
      'Content-Type': 'application/json',
      Prefer:
        'resolution=merge-duplicates,return=minimal'
    };

    async function upsert(table, rows, conflict) {
      if (!rows || !rows.length) return 0;

      let count = 0;

      for (
        let i = 0;
        i < rows.length;
        i += 300
      ) {
        const batch = rows
          .slice(i, i + 300)
          .map(row => ({
            ...row,
            synced_at: new Date().toISOString()
          }));

        const requestUrl =
          supabaseUrl +
          '/rest/v1/' +
          table +
          '?on_conflict=' +
          encodeURIComponent(conflict);

        const response = await fetch(
          requestUrl,
          {
            method: 'POST',
            headers,
            body: JSON.stringify(batch)
          }
        );

        if (!response.ok) {
          const detail = (
            await response.text()
          ).slice(0, 500);

          throw new Error(
            'Supabase 寫入 ' +
            table +
            ' 失敗（HTTP ' +
            response.status +
            '）：' +
            detail
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
      learning_source:
        p.learning_source || null,
      source_created_at:
        p.source_created_at || null,
      source_row: p.source_row
    }));

    const shippingRows = (
      sourceData.shippingRows || []
    ).map(row => ({
      ...row,
      raw_product_text:
        row.raw_product_text || null,
      product_code:
        row.product_code || null,
      shipping_name:
        row.shipping_name || null
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
      message:
        '今日託運單與商品資料同步完成',
      source_date: sourceData.source_date,
      counts: {
        barcodeProducts: barcodeCount,
        shippingRows: shippingCount,
        qiaowenRows:
          (sourceData.counts || {}).qiaowenRows || 0,
        shopeeRows:
          (sourceData.counts || {}).shopeeRows || 0
      },
      note:
        '掃描託運單時改從 Supabase 讀取已同步資料。'
    });

  } catch (error) {
    return res.status(502).json({
      ok: false,
      message: error.message || String(error)
    });
  }
}
