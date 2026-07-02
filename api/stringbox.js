const PLANT_ID_DEFAULT = "ad493847-3dd7-4526-9122-123e35d1374a";
const PLANT_TZ = "America/Fortaleza";

function todayInPlantTz() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: PLANT_TZ, year: "numeric", month: "2-digit", day: "2-digit",
  }).formatToParts(new Date());
  const map = Object.fromEntries(parts.map(p => [p.type, p.value]));
  return `${map.year}${map.month}${map.day}`;
}

// Proxy para GET /stringbox/samplesv2/plant/{plantId}/date/{date} do INGECON SUN Monitor.
// A API key fica só aqui no servidor — nunca é exposta ao browser.
export default async function handler(req, res) {
  if (req.method !== "GET") {
    res.setHeader("Allow", "GET");
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { date } = req.query;
  if (!date || !/^\d{8}$/.test(date)) {
    return res.status(400).json({ error: "Parâmetro 'date' inválido — use o formato YYYYMMDD" });
  }

  const apiKey = process.env.INGECON_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "INGECON_API_KEY não configurada no servidor" });
  }
  const plantId = process.env.INGECON_PLANT_ID || PLANT_ID_DEFAULT;

  const upstreamUrl = `https://www.ingeconsunmonitor.com/api/stringbox/samplesv2/plant/${plantId}/date/${date}`;

  let upstreamRes;
  try {
    upstreamRes = await fetch(upstreamUrl, {
      headers: { "X-API-KEY": apiKey, "Accept-Encoding": "gzip" },
    });
  } catch {
    return res.status(502).json({ error: "Falha ao contatar a API do INGECON SUN Monitor" });
  }

  if (!upstreamRes.ok) {
    const text = await upstreamRes.text().catch(() => "");
    return res.status(upstreamRes.status).json({
      error: `INGECON API retornou ${upstreamRes.status}`,
      detail: text.slice(0, 500),
    });
  }

  const data = await upstreamRes.json();

  const isToday = date === todayInPlantTz();
  res.setHeader(
    "Cache-Control",
    isToday
      ? "public, s-maxage=60, stale-while-revalidate=300"
      : "public, s-maxage=86400, stale-while-revalidate=604800"
  );

  return res.status(200).json(data);
}
