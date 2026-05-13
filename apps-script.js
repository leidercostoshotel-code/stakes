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

const USERS_SHEET = "Usuarios";
const META_SHEET  = "Meta";

// ── Helpers ──────────────────────────────────────────────────────────────────

function sanitizeEmail(email) {
  return email.toLowerCase().replace(/[@.]/g, "_");
}

function dataSheetName(email) {
  return "Data_" + sanitizeEmail(email);
}

function getUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET);
    sheet.appendRow(["ID", "Nombre", "Email", "PasswordHash", "CreadoEl"]);
    const hdr = sheet.getRange(1, 1, 1, 5);
    hdr.setBackground("#1e293b").setFontColor("#f8fafc").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findUser(email) {
  const sheet = getUsersSheet();
  const rows  = sheet.getDataRange().getValues();
  const headers = rows[0];
  const emailIdx = headers.indexOf("Email");
  const hashIdx  = headers.indexOf("PasswordHash");
  const nameIdx  = headers.indexOf("Nombre");
  const idIdx    = headers.indexOf("ID");

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][emailIdx]).toLowerCase() === email.toLowerCase()) {
      return {
        row: i + 1,
        id:   rows[i][idIdx],
        name: rows[i][nameIdx],
        email: rows[i][emailIdx],
        passwordHash: rows[i][hashIdx]
      };
    }
  }
  return null;
}

function verifyUser(email, passwordHash) {
  const user = findUser(email);
  if (!user) return false;
  return user.passwordHash === passwordHash;
}

// ── HTTP handlers ─────────────────────────────────────────────────────────────

function doGet(e) {
  const action       = e.parameter.action;
  const callback     = e.parameter.callback;
  const userId       = e.parameter.userId       || "";
  const passwordHash = e.parameter.passwordHash || "";
  let result;

  if      (action === "load")         result = loadData(userId, passwordHash);
  else if (action === "login")        result = loginUser({ email: userId, passwordHash });
  else if (action === "register")     result = registerUser({ name: e.parameter.name || "", email: userId, passwordHash });
  else if (action === "requestReset") result = requestReset(userId);
  else if (action === "confirmReset") result = confirmReset({ email: userId, code: e.parameter.code || "", newPasswordHash: e.parameter.newPasswordHash || "" });
  else result = { status: "error", message: "Acción no reconocida" };

  const json = JSON.stringify(result);
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
    if (payload.action === "register") return jsonResponse(registerUser(payload));
    if (payload.action === "login")    return jsonResponse(loginUser(payload));
    if (payload.action === "save")     return jsonResponse(saveData(payload));
    return jsonResponse({ status: "error", message: "Acción no reconocida" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

// ── Auth actions ──────────────────────────────────────────────────────────────

function registerUser(payload) {
  const { name, email, passwordHash } = payload;
  if (!name || !email || !passwordHash) {
    return { status: "error", message: "Datos incompletos" };
  }

  const existing = findUser(email);
  if (existing) {
    return { status: "error", message: "Este email ya está registrado" };
  }

  const sheet = getUsersSheet();
  const id = Utilities.getUuid();
  sheet.appendRow([id, name, email.toLowerCase(), passwordHash, new Date().toISOString()]);

  return { status: "ok", name };
}

function loginUser(payload) {
  const { email, passwordHash } = payload;
  if (!email || !passwordHash) {
    return { status: "error", message: "Datos incompletos" };
  }

  const user = findUser(email);
  if (!user || user.passwordHash !== passwordHash) {
    return { status: "error", message: "Credenciales incorrectas" };
  }

  return { status: "ok", name: user.name };
}

// ── Password reset ────────────────────────────────────────────────────────────

function requestReset(email) {
  if (!email) return { status: "error", message: "Email requerido" };

  const user = findUser(email);
  if (!user) return { status: "error", message: "No existe una cuenta con ese email" };

  const code    = String(Math.floor(100000 + Math.random() * 900000)); // 6 dígitos
  const expiry  = new Date(Date.now() + 15 * 60 * 1000).toISOString(); // 15 min

  // Guardar código en hoja ResetTokens
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let rtSheet = ss.getSheetByName("ResetTokens");
  if (!rtSheet) {
    rtSheet = ss.insertSheet("ResetTokens");
    rtSheet.appendRow(["Email", "Code", "Expiry"]);
  }

  // Eliminar tokens previos del mismo email
  const data = rtSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
      rtSheet.deleteRow(i + 1);
    }
  }

  rtSheet.appendRow([email.toLowerCase(), code, expiry]);

  // Enviar email con el código
  GmailApp.sendEmail(
    email,
    "🔑 Código de recuperación — BetControl Pro",
    "",
    {
      htmlBody: `
        <div style="font-family:Inter,sans-serif;max-width:480px;margin:0 auto;background:#070b18;color:#f8fafc;border-radius:16px;overflow:hidden">
          <div style="background:linear-gradient(135deg,#0ea5e9,#7c3aed);padding:28px 32px;text-align:center">
            <h1 style="margin:0;font-size:24px;color:#fff">⚽ BetControl Pro</h1>
          </div>
          <div style="padding:32px">
            <h2 style="font-size:20px;margin-bottom:12px;color:#f8fafc">Recuperación de contraseña</h2>
            <p style="color:#94a3b8;margin-bottom:24px">Usa el siguiente código para restablecer tu contraseña. Expira en <strong style="color:#f8fafc">15 minutos</strong>.</p>
            <div style="background:#0f172a;border:1px solid #1e3a5f;border-radius:12px;padding:24px;text-align:center;margin-bottom:24px">
              <span style="font-size:40px;font-weight:900;letter-spacing:10px;color:#38bdf8">${code}</span>
            </div>
            <p style="color:#64748b;font-size:12px">Si no solicitaste este código, ignora este mensaje.</p>
          </div>
        </div>
      `
    }
  );

  return { status: "ok" };
}

function confirmReset({ email, code, newPasswordHash }) {
  if (!email || !code || !newPasswordHash) return { status: "error", message: "Datos incompletos" };

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const rtSheet = ss.getSheetByName("ResetTokens");
  if (!rtSheet) return { status: "error", message: "Código inválido o expirado" };

  const data = rtSheet.getDataRange().getValues();
  let found  = null;
  let rowIdx = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email.toLowerCase() && String(data[i][1]) === String(code)) {
      found  = data[i];
      rowIdx = i + 1;
      break;
    }
  }

  if (!found) return { status: "error", message: "Código incorrecto" };
  if (new Date(found[2]) < new Date()) {
    rtSheet.deleteRow(rowIdx);
    return { status: "error", message: "El código ha expirado. Solicita uno nuevo" };
  }

  // Actualizar contraseña en hoja Usuarios
  const usersSheet = getUsersSheet();
  const rows = usersSheet.getDataRange().getValues();
  const headers = rows[0];
  const emailIdx = headers.indexOf("Email");
  const hashIdx  = headers.indexOf("PasswordHash");

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][emailIdx]).toLowerCase() === email.toLowerCase()) {
      usersSheet.getRange(i + 1, hashIdx + 1).setValue(newPasswordHash);
      break;
    }
  }

  rtSheet.deleteRow(rowIdx);
  return { status: "ok" };
}

// ── Data actions ──────────────────────────────────────────────────────────────

function saveData(payload) {
  const { userId, passwordHash } = payload;

  // If userId provided, verify; anonymous saves still allowed for backwards compat
  if (userId && passwordHash) {
    if (!verifyUser(userId, passwordHash)) {
      return { status: "error", message: "No autorizado" };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = userId ? dataSheetName(userId) : "Apuestas";

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

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

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1e293b").setFontColor("#f8fafc").setFontWeight("bold");
  sheet.setFrozenRows(1);

  // Meta sheet per user
  const metaName = userId ? "Meta_" + sanitizeEmail(userId) : META_SHEET;
  let meta = ss.getSheetByName(metaName);
  if (!meta) meta = ss.insertSheet(metaName);
  meta.clearContents();
  meta.appendRow(["Clave", "Valor"]);
  const cfg = payload.settings || {};
  meta.appendRow(["bancaInicial", cfg.bancaInicial || 0]);
  meta.appendRow(["unidadPct",    cfg.unidadPct    || 2]);
  meta.appendRow(["exportadoEl",  new Date().toISOString()]);
  meta.appendRow(["totalApuestas", bets.length]);

  return { status: "ok", saved: bets.length };
}

function loadData(userId, passwordHash) {
  if (userId && passwordHash) {
    if (!verifyUser(userId, passwordHash)) {
      return { status: "error", message: "No autorizado" };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = userId ? dataSheetName(userId) : "Apuestas";

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: "ok", bets: [], settings: {} };

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

  const metaName = userId ? "Meta_" + sanitizeEmail(userId) : META_SHEET;
  const meta = ss.getSheetByName(metaName);
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
