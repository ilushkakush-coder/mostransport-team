const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwRe5CmVRT-AsHEB-d2IFlAOz8arWwBxQCrg8JQGlE79Af0_81D8OdqdAXYTSXX6tOD/exec';

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Метод не поддерживается' });
  }

  try {
    const payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    const upstream = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
      body: JSON.stringify(payload || {})
    });
    const text = await upstream.text();
    let result;
    try {
      result = JSON.parse(text);
    } catch (_) {
      result = { ok: false, error: `Apps Script вернул неожиданный ответ (${upstream.status})` };
    }

    return res.status(upstream.ok ? 200 : 502).json(result);
  } catch (error) {
    console.error('Apps Script proxy failed', error);
    return res.status(502).json({ ok: false, error: 'Сервер не смог связаться с Apps Script' });
  }
};
