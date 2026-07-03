import { getSupabase, lazyCleanup } from "./_lib/supabase.js";
import { plantId, ingeconHeaders, todayInPlantTz, toISODate } from "./_lib/ingecon.js";

function eachIsoDate(fromIso, toIso) {
  const out = [];
  const d = new Date(fromIso + "T00:00:00Z");
  const end = new Date(toIso + "T00:00:00Z");
  while (d <= end) {
    out.push(d.toISOString().slice(0, 10));
    d.setUTCDate(d.getUTCDate() + 1);
  }
  return out;
}

// Proxy + cache (Supabase) para geração diária por inversor.
// Usa /ingecon/samplesv2/groupbyday, que cobre um intervalo inteiro de dias numa única chamada.
export default async function handler(req, res) {
  if (req.method !== "GET") {
    res.setHeader("Allow", "GET");
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { from, to } = req.query;
  if (!from || !/^\d{8}$/.test(from) || !to || !/^\d{8}$/.test(to)) {
    return res.status(400).json({ error: "Parâmetros 'from' e 'to' inválidos — use o formato YYYYMMDD" });
  }

  let supabase;
  try {
    supabase = getSupabase();
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }

  const fromIso = toISODate(from);
  const toIso = toISODate(to);
  const today = todayInPlantTz();
  const includesToday = from <= today && today <= to;

  const { data: cached, error: readErr } = await supabase
    .from("ingecon_generation_daily")
    .select("sn,date,e_injection,e_absorption")
    .gte("date", fromIso)
    .lte("date", toIso);
  if (readErr) return res.status(500).json({ error: "Erro ao ler cache: " + readErr.message });

  const cachedDates = new Set((cached || []).map(r => r.date));
  const allDates = eachIsoDate(fromIso, toIso);
  const hasGap = allDates.some(d => !cachedDates.has(d));
  const needsFetch = hasGap || includesToday;

  let merged = cached || [];

  if (needsFetch) {
    let headers;
    try {
      headers = ingeconHeaders();
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
    const url = `https://www.ingeconsunmonitor.com/api/ingecon/samplesv2/groupbyday/plant/${plantId()}/from/${from}/to/${to}`;
    let upstreamRes;
    try {
      upstreamRes = await fetch(url, { headers });
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
    const records = await upstreamRes.json();

    // Alguns inversores reportam por mais de um "board" (ex.: 2 registros pro mesmo SN+dia) —
    // soma as duplicatas em vez de deixar o upsert falhar (ON CONFLICT não aceita a mesma
    // chave duas vezes no mesmo lote).
    const bySnDate = new Map();
    records.filter(r => r.SN && r.DateTime).forEach(r => {
      const key = `${r.SN}|${String(r.DateTime).slice(0, 10)}`;
      const prev = bySnDate.get(key);
      if (prev) {
        prev.e_injection = (prev.e_injection ?? 0) + (r.EInjection ?? 0);
        prev.e_absorption = (prev.e_absorption ?? 0) + (r.EAbsorption ?? 0);
      } else {
        bySnDate.set(key, {
          sn: r.SN,
          date: String(r.DateTime).slice(0, 10),
          e_injection: r.EInjection ?? null,
          e_absorption: r.EAbsorption ?? null,
          updated_at: new Date().toISOString(),
        });
      }
    });
    const rows = [...bySnDate.values()];

    if (rows.length) {
      const { error: upsertErr } = await supabase
        .from("ingecon_generation_daily")
        .upsert(rows, { onConflict: "sn,date" });
      if (upsertErr) return res.status(500).json({ error: "Erro ao gravar cache: " + upsertErr.message });
    }

    const byKey = new Map();
    (cached || []).forEach(r => byKey.set(`${r.sn}|${r.date}`, r));
    rows.forEach(r => byKey.set(`${r.sn}|${r.date}`, r));
    merged = [...byKey.values()];
  }

  await lazyCleanup(supabase).catch(() => {});

  res.setHeader("Cache-Control", includesToday
    ? "public, s-maxage=60, stale-while-revalidate=300"
    : "public, s-maxage=3600, stale-while-revalidate=86400");

  return res.status(200).json(
    merged.map(r => ({ sn: r.sn, date: r.date, eInjection: r.e_injection, eAbsorption: r.e_absorption }))
  );
}
