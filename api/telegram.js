const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwRe5CmVRT-AsHEB-d2IFlAOz8arWwBxQCrg8JQGlE79Af0_81D8OdqdAXYTSXX6tOD/exec';

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Метод не поддерживается' });
  }

  const hookSecret = String(req.query && req.query.hook || '');
  if (!hookSecret) {
    return res.status(401).json({ ok: false, error: 'Webhook secret is missing' });
  }

  try {
    const payload = typeof req.body === 'string' ? req.body : JSON.stringify(req.body || {});
    const upstream = await fetch(`${APPS_SCRIPT_URL}?hook=${encodeURIComponent(hookSecret)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
      body: payload
    });

    if (!upstream.ok) {
      console.error('Telegram relay upstream failed with status', upstream.status);
      return res.status(502).json({ ok: false, error: 'Apps Script не принял обновление Telegram' });
    }

    return res.status(200).json({ ok: true });
  } catch (error) {
    console.error('Telegram webhook relay failed', error);
    return res.status(502).json({ ok: false, error: 'Не удалось передать обновление Apps Script' });
  }
};
