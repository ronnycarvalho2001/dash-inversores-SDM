export const PLANT_ID_DEFAULT = "ad493847-3dd7-4526-9122-123e35d1374a";
export const PLANT_TZ = "America/Fortaleza";

export function todayInPlantTz() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: PLANT_TZ, year: "numeric", month: "2-digit", day: "2-digit",
  }).formatToParts(new Date());
  const map = Object.fromEntries(parts.map(p => [p.type, p.value]));
  return `${map.year}${map.month}${map.day}`;
}

export function plantId() {
  return process.env.INGECON_PLANT_ID || PLANT_ID_DEFAULT;
}

export function ingeconHeaders() {
  const apiKey = process.env.INGECON_API_KEY;
  if (!apiKey) throw new Error("INGECON_API_KEY não configurada no servidor");
  return { "X-API-KEY": apiKey, "Accept-Encoding": "gzip" };
}

export function toISODate(yyyymmdd) {
  return `${yyyymmdd.slice(0,4)}-${yyyymmdd.slice(4,6)}-${yyyymmdd.slice(6,8)}`;
}
