import { plantId, ingeconHeaders, todayInPlantTz } from "./_lib/ingecon.js";

// "SM3/INV3.2.1ST" → "3.2.1"
function posFromGid(gid) {
  const m = String(gid).match(/INV(\d+)\.(\d+\.\d+)ST$/i);
  return m ? `${m[1]}.${m[2]}` : null;
}

function yesterdayYmd() {
  const today = todayInPlantTz();
  const y = parseInt(today.slice(0, 4), 10), mo = parseInt(today.slice(4, 6), 10) - 1, d = parseInt(today.slice(6, 8), 10);
  const dt = new Date(Date.UTC(y, mo, d));
  dt.setUTCDate(dt.getUTCDate() - 1);
  return `${dt.getUTCFullYear()}${String(dt.getUTCMonth() + 1).padStart(2, "0")}${String(dt.getUTCDate()).padStart(2, "0")}`;
}

// Mapeia BoardId → posição real do combinador (ex.: "3.2.1"). BoardId é o identificador
// verdadeiramente único por inversor — o serviço de stringbox já traz GId+BoardId juntos,
// então não precisa cruzar com o serviço de inversores (que tem SNs duplicados/reaproveitados).
// Usa o dia de ontem (já fechado, cobertura completa).
export default async function handler(req, res) {
  if (req.method !== "GET") {
    res.setHeader("Allow", "GET");
    return res.status(405).json({ error: "Method not allowed" });
  }

  let headers;
  try {
    headers = ingeconHeaders();
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }

  const date = yesterdayYmd();
  const url = `https://www.ingeconsunmonitor.com/api/stringbox/samplesv2/plant/${plantId()}/date/${date}`;

  let sbRes;
  try {
    sbRes = await fetch(url, { headers });
  } catch {
    return res.status(502).json({ error: "Falha ao contatar a API do INGECON SUN Monitor" });
  }
  if (!sbRes.ok) {
    const text = await sbRes.text().catch(() => "");
    return res.status(sbRes.status).json({ error: `INGECON API retornou ${sbRes.status}`, detail: text.slice(0, 500) });
  }

  const sbRecords = await sbRes.json();
  const map = {};
  sbRecords.forEach(r => {
    if (!r.BoardId || !r.GId) return;
    const pos = posFromGid(r.GId);
    if (pos) map[r.BoardId] = pos;
  });

  res.setHeader("Cache-Control", "public, s-maxage=86400, stale-while-revalidate=604800");
  return res.status(200).json(map);
}
