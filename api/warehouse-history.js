/* 倉庫掃碼日期歷史：Google Sheets 為來源，Supabase 為快取。 */
export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'private, no-store');
  if (req.method !== 'POST') return res.status(405).json({ok:false,message:'僅支援 POST'});
  try {
    const body = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {});
    const mode = String(body.mode || 'read');
    const date = String(body.date || '').trim();
    const parsedDate = /^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(date) ? new Date(date + 'T00:00:00Z') : null;
    if (!parsedDate || !Number.isFinite(parsedDate.getTime()) || parsedDate.toISOString().slice(0,10) !== date) {
      return res.status(400).json({ok:false,message:'日期必須是有效的 YYYY-MM-DD'});
    }
    if (mode !== 'read' && mode !== 'sync') return res.status(400).json({ok:false,message:'不支援的查詢模式'});
    const url = String(process.env.SUPABASE_URL || 'https://kaiduxzwmqsfmjxsfwpp.supabase.co').replace(/\/$/,'');
    const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
    if (!key) return res.status(500).json({ok:false,message:'Vercel 尚未設定 SUPABASE_SERVICE_ROLE_KEY'});
    const endpoint = url + '/rest/v1/warehouse_scan_history_cache';
    const headers = {apikey:key,Authorization:'Bearer ' + key,'Content-Type':'application/json'};
    if (mode === 'read') {
      const qs = new URLSearchParams({select:'source_date,orders,synced_at',source_date:'eq.' + date,limit:'1'});
      const response = await fetch(endpoint + '?' + qs, {headers});
      const raw = await response.text();
      if (!response.ok) throw new Error(dbError(response.status, raw));
      const rows = JSON.parse(raw);
      if (!rows.length) return res.status(200).json({ok:true,found:false,sourceDate:date,orders:[]});
      return res.status(200).json({ok:true,found:true,sourceDate:date,orders:rows[0].orders,syncedAt:rows[0].synced_at});
    }
    // 同步單一日期，不每次全表搬運；只由伺服器持有 API Token 和 service-role key。
    const appsUrl = process.env.APPS_SCRIPT_URL;
    if (!appsUrl) return res.status(500).json({ok:false,message:'Vercel 尚未設定 APPS_SCRIPT_URL'});
    const payload = {action:'getScanHistoryByDate',date,force:true};
    if (process.env.APPS_SCRIPT_API_TOKEN) payload.apiToken = process.env.APPS_SCRIPT_API_TOKEN;
    const source = await fetch(appsUrl, {
      method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify(payload),signal:AbortSignal.timeout(60000)
    });
    const raw = await source.text();
    let data;
    try { data = JSON.parse(raw); }
    catch (_) { throw new Error('Apps Script 回傳不是 JSON，請檢查部署網址與存取權限'); }
    if (!source.ok || !data.ok || !Array.isArray(data.orders)) {
      throw new Error(data.message || 'Google Sheets 歷史讀取失敗');
    }
    const syncedAt = new Date().toISOString();
    const record = {source_date:date,orders:data.orders,synced_at:syncedAt};
    const query = new URLSearchParams({on_conflict:'source_date'});
    const saved = await fetch(endpoint + '?' + query, {
      method:'POST',headers:{...headers,Prefer:'resolution=merge-duplicates,return=minimal'},
      body:JSON.stringify([record])
    });
    if (!saved.ok) throw new Error(dbError(saved.status, await saved.text()));
    return res.status(200).json({ok:true,found:true,sourceDate:date,orders:data.orders,syncedAt});
  } catch (error) {
    return res.status(502).json({ok:false,message:error?.message || String(error)});
  }
}
function dbError(status, raw) {
  const detail = String(raw || '').slice(0,240);
  if (status === 404 && /PGRST205|Could not find the table|does not exist/i.test(detail)) {
    return '尚未建立 Supabase 資料表 warehouse_scan_history_cache，請先執行安裝 SQL';
  }
  return 'Supabase HTTP ' + status + '：' + detail;
}
