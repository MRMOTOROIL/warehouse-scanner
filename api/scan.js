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
