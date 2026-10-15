export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, message: 'Method Not Allowed' });
  }

  const appsScriptUrl = process.env.APPS_SCRIPT_URL;
  const apiToken = process.env.APPS_SCRIPT_API_TOKEN || '';

  if (!appsScriptUrl) {
    return res.status(500).json({ ok: false, message: '尚未設定 APPS_SCRIPT_URL' });
  }

  try {
    const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});

    if (body.action === 'syncTodayData') {
      return await syncTodayData(req, res, body);
    }
    const payload = { ...body };

    if (apiToken) {
      payload.apiToken = apiToken;
    }

    const response = await fetch(appsScriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    const text = await response.text();

    let data;
    try {
      data = JSON.parse(text);
    } catch {
      data = {
        ok: false,
        message: 'Apps Script 回應不是有效 JSON',
        raw: text.slice(0, 500)
      };
    }

    return res.status(response.ok ? 200 : response.status).json(data);
  } catch (error) {
    return res.status(502).json({
      ok: false,
      message: error?.message || String(error)
    });
  }
}


async function syncTodayData(req, res, body) {
  const appsScriptUrl = process.env.APPS_SCRIPT_URL;
  const apiToken = process.env.APPS_SCRIPT_API_TOKEN || '';
  const supabaseUrl = (process.env.SUPABASE_URL || 'https://kaiduxzwmqsfmjxsfwpp.supabase.co').replace(/\/$/, '');
  const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

  if (!appsScriptUrl) return res.status(500).json({ ok: false, message: '尚未設定 APPS_SCRIPT_URL' });
  if (!serviceKey) return res.status(500).json({ ok: false, message: '尚未設定 SUPABASE_SERVICE_ROLE_KEY' });

  const sourcePayload = { action: 'exportWarehouseSyncData' };
  if (apiToken) sourcePayload.apiToken = apiToken;

  const sourceResponse = await fetch(appsScriptUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(sourcePayload)
  });
  const sourceText = await sourceResponse.text();
  let sourceData;
  try { sourceData = JSON.parse(sourceText); }
  catch { return res.status(502).json({ ok: false, message: 'Apps Script 回應不是有效 JSON', detail: sourceText.slice(0, 300) }); }
  if (!sourceResponse.ok || !sourceData.ok) {
    return res.status(502).json({ ok: false, message: sourceData.message || '讀取 Google Sheets 同步資料失敗' });
  }

  const headers = {
    apikey: serviceKey,
    Authorization: 'Bearer ' + serviceKey,
    'Content-Type': 'application/json',
    Prefer: 'resolution=merge-duplicates,return=minimal'
  };

  async function upsert(table, rows, conflict) {
    if (!rows || !rows.length) return 0;
    let count = 0;
    for (let i = 0; i < rows.length; i += 300) {
      const batch = rows.slice(i, i + 300).map(row => ({ ...row, synced_at: new Date().toISOString() }));
      const url = supabaseUrl + '/rest/v1/' + table + '?on_conflict=' + encodeURIComponent(conflict);
      const response = await fetch(url, { method: 'POST', headers, body: JSON.stringify(batch) });
      if (!response.ok) {
        const detail = (await response.text()).slice(0, 500);
        throw new Error('Supabase 寫入 ' + table + ' 失敗（HTTP ' + response.status + '）：' + detail);
      }
      count += batch.length;
    }
    return count;
  }

  try {
    const barcodeRows = (sourceData.barcodeProducts || []).map(p => ({
      product_code: p.product_code || null,
      shipping_name: p.shipping_name,
      barcode: p.barcode || null,
      full_name: p.full_name || null,
      note: p.note || null,
      learning_source: p.learning_source || null,
      source_created_at: p.source_created_at || null,
      source_row: p.source_row
    }));
    const shippingRows = (sourceData.shippingRows || []).map(row => ({
      ...row,
      raw_product_text: row.raw_product_text || null,
      product_code: row.product_code || null,
      shipping_name: row.shipping_name || null
    }));

    const barcodeCount = await upsert('warehouse_product_barcodes', barcodeRows, 'source_row');
    const shippingCount = await upsert('warehouse_shipping_source_rows', shippingRows, 'source,source_date,source_row');

    return res.status(200).json({
      ok: true,
      message: '今日託運單與商品資料同步完成',
      source_date: sourceData.source_date,
      counts: {
        barcodeProducts: barcodeCount,
        shippingRows: shippingCount,
        qiaowenRows: (sourceData.counts || {}).qiaowenRows || 0,
        shopeeRows: (sourceData.counts || {}).shopeeRows || 0
      },
      note: '此階段已同步至 Supabase；現有託運單查詢仍由 Apps Script 處理。'
    });
  } catch (error) {
    return res.status(502).json({ ok: false, message: error.message || String(error) });
  }
}
