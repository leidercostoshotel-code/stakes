/**
 * BetControl Pro — Google Apps Script backend
 *
 * Instrucciones:
 * 1. Abre script.google.com y crea un proyecto nuevo.
 * 2. Reemplaza todo el contenido con este archivo.
 * 3. Implementar → Nueva implementación → Aplicación web
 *    - Ejecutar como: Yo
 *    - Acceso: Cualquier persona
 * 4. Copia la URL y pégala en la app, sección "Google Sheets".
 */

const SHEET_NAME = "Apuestas";
const META_SHEET = "Meta";

function doGet(e) {
  const action   = e.parameter.action;
  const callback = e.parameter.callback;
  const result   = (action === "load") ? loadData() : { status: "error", message: "Acción no reconocida" };
  const json     = JSON.stringify(result);
  // JSONP: si viene un parámetro callback, envolver la respuesta para evitar CORS
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === "save") return jsonResponse(saveData(payload));
    return jsonResponse({ status: "error", message: "Acción no reconocida" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

function saveData(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Hoja de apuestas ────────────────────────────────────────────────────
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  sheet.clearContents();

  const headers = [
    "ID", "Fecha", "Tipo", "Apuesta formulada", "Mercado", "Casa",
    "Cuota", "Stake", "Costo compra", "Resultado", "Confianza",
    "Ganancia neta", "ROI %", "Nota", "Creada el", "Cerrada el"
  ];
  sheet.appendRow(headers);

  const bets = payload.bets || [];
  bets.forEach(b => {
    const stake = Number(b.stake) || 0;
    const cuota = Number(b.cuota) || 0;
    const costo = Number(b.costoCompra) || 0;
    let profit = -costo;
    if (b.resultado === "Ganada")  profit = stake * cuota - stake - costo;
    if (b.resultado === "Perdida") profit = -stake - costo;
    const inversion = stake + costo;
    const roi = (inversion && b.resultado !== "Pendiente") ? (profit / inversion * 100) : 0;

    sheet.appendRow([
      b.id, b.fecha, b.tipo, b.formulada || "", b.mercado, b.casa || "",
      b.cuota, b.stake, costo, b.resultado, b.confianza || "Media",
      profit.toFixed(2), roi.toFixed(2), b.nota || "",
      b.createdAt || "", b.closedAt || ""
    ]);
  });

  // Formato cabecera
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1e293b").setFontColor("#f8fafc").setFontWeight("bold");
  sheet.setFrozenRows(1);

  // ── Hoja de metadatos / configuración ───────────────────────────────────
  let meta = ss.getSheetByName(META_SHEET);
  if (!meta) meta = ss.insertSheet(META_SHEET);
  meta.clearContents();
  meta.appendRow(["Clave", "Valor"]);
  const cfg = payload.settings || {};
  meta.appendRow(["bancaInicial", cfg.bancaInicial || 0]);
  meta.appendRow(["unidadPct",    cfg.unidadPct    || 2]);
  meta.appendRow(["exportadoEl",  new Date().toISOString()]);
  meta.appendRow(["totalApuestas", bets.length]);

  return { status: "ok", saved: bets.length };
}

function loadData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { status: "error", message: "Hoja 'Apuestas' no encontrada. Guarda primero desde la app." };

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { status: "ok", bets: [], settings: {} };

  const headers = rows[0];
  const idx = k => headers.indexOf(k);

  const bets = rows.slice(1).map(row => ({
    id:          row[idx("ID")],
    fecha:       row[idx("Fecha")],
    tipo:        row[idx("Tipo")],
    formulada:   row[idx("Apuesta formulada")],
    mercado:     row[idx("Mercado")],
    casa:        row[idx("Casa")],
    cuota:       Number(row[idx("Cuota")]) || 0,
    stake:       Number(row[idx("Stake")]) || 0,
    costoCompra: Number(row[idx("Costo compra")]) || 0,
    resultado:   row[idx("Resultado")],
    confianza:   row[idx("Confianza")] || "Media",
    nota:        row[idx("Nota")] || "",
    createdAt:   row[idx("Creada el")] || "",
    closedAt:    row[idx("Cerrada el")] || null,
  })).filter(b => b.id);

  // Leer configuración
  const meta = ss.getSheetByName(META_SHEET);
  const settings = { bancaInicial: 0, unidadPct: 2 };
  if (meta) {
    meta.getDataRange().getValues().slice(1).forEach(([k, v]) => {
      if (k === "bancaInicial") settings.bancaInicial = Number(v) || 0;
      if (k === "unidadPct")    settings.unidadPct    = Number(v) || 2;
    });
  }

  return { status: "ok", bets, settings };
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
