export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, error: "Método não permitido." });
  }

  const appsScriptUrl = process.env.APPS_SCRIPT_URL;
  if (!appsScriptUrl) {
    return res.status(500).json({
      ok: false,
      error: "O servidor ainda não está configurado. Falta APPS_SCRIPT_URL no Vercel."
    });
  }

  try {
    const response = await fetch(appsScriptUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(req.body || {})
    });

    const text = await response.text();
    let data;
    try { data = JSON.parse(text); }
    catch { data = { ok: response.ok, message: text }; }

    if (!response.ok || data.ok === false) {
      return res.status(502).json({
        ok: false,
        error: data.error || "O serviço de processamento recusou o pedido."
      });
    }

    return res.status(200).json({ ok: true, ...data });
  } catch (err) {
    return res.status(500).json({
      ok: false,
      error: "Falha de comunicação com o serviço de email."
    });
  }
}
