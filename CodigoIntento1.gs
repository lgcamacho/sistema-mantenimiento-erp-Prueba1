const SHEET_NAME = "Datos";
const LOG_SHEET_NAME = "Log Cerrados";
const EHS_SHEET_NAME = "Validacion por EHS";
const OPEN_SHEET_NAME = "MTTOs Abiertos";
const REF_SHEET_NAME = "Log Refacciones";
const OEE_SHEET_NAME = "OEE_Registro";
const PERSONAL_SHEET_NAME = "Log_Personal";
const FOLDER_NAME = "Evidencias MTTO Nidec";
const ADMIN_PASS = "1451MR";
const CORREO_COMPRAS = "jose.h.ramirez@nidec-ga.com, luis.barrita@nidec-ga.com, genaro.deleon@nidec-ga.com, sara.o.rodriguez@nidec-ga.com, margarita.b.barboza@nidec-ga.com, jose.s.sanchez@nidec-ga.com, wenceslao.pacheco@nidec-ga.com, noe.martinez@nidec-ga.com";

function getDB() {
  try { return SpreadsheetApp.openById('1--AsNXslAhbu3PxAldhAtAMIA2a5LZRBvrtbmFLOZc0'); } 
  catch(e) { return SpreadsheetApp.getActiveSpreadsheet(); }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ERP MAINTENANCE SYSTEM')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function parseAnyDate(val) {
  if (!val) return null;
  let d = null;
  let now = new Date();

  // 1. Si ya viene como objeto Date desde Google Sheets
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    d = val;
  } else if (typeof val === 'number') {
    d = new Date(val);
  } else {
    let str = String(val).trim();
    if (!str || str === "N/A" || str === "Sin fecha" || str === "Pendiente") return null;
    if (/^\d{10,13}$/.test(str)) return new Date(parseInt(str, 10));

    // Limpieza de comas y espacios
    str = str.replace(/,/g, " ").replace(/\s+/g, " ").trim();

    // Detección de formato DD/MM/YYYY o MM/DD/YYYY
    let match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (match) {
      let p1 = parseInt(match[1], 10);
      let p2 = parseInt(match[2], 10);
      let year = parseInt(match[3], 10);
      
      let timeMatch = str.match(/(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
      let hh = timeMatch ? parseInt(timeMatch[1], 10) : 0;
      let mm = timeMatch ? parseInt(timeMatch[2], 10) : 0;

      // Si p2 > 12, p1 es el mes y p2 es el día
      if (p2 > 12) {
        d = new Date(year, p1 - 1, p2, hh, mm);
      } else {
        // Intenta asumir p1 = día, p2 = mes (Formato Mexicano/Español)
        d = new Date(year, p2 - 1, p1, hh, mm);
      }
    } else {
      d = new Date(str);
    }
  }

  if (!d || isNaN(d.getTime())) return null;

  // ESCUDO INTELIGENTE: Si la fecha es futura (ej. Octubre/Noviembre/Diciembre por confusión de mes/día)
  // Invierte automáticamente Mes y Día para corregir a Agosto
  if (d.getTime() > now.getTime() + 86400000) { 
    let correctedDay = d.getMonth() + 1;  // El mes equivocado era en realidad el día
    let correctedMonth = d.getDate() - 1; // El día equivocado era en realidad el mes
    d = new Date(d.getFullYear(), correctedMonth, correctedDay, d.getHours(), d.getMinutes());
  }

  return d;
}

function checkAndInitAllSheets() {
  try {
    const ss = getDB();
    if (!ss.getSheetByName(OPEN_SHEET_NAME)) { let sh = ss.insertSheet(OPEN_SHEET_NAME); sh.appendRow(["Folio", "Fecha", "Qué", "Por qué", "Dónde", "Cuándo", "Quién", "Cómo", "Refacciones", "Evidencia Inicial", "Timestamp"]); sh.getRange("A1:K1").setFontWeight("bold").setBackground("#3b82f6").setFontColor("white"); }
    if (!ss.getSheetByName(LOG_SHEET_NAME)) { let sh = ss.insertSheet(LOG_SHEET_NAME); sh.appendRow(["Folio", "Fecha Cierre", "Tipo", "Supervisor", "Técnico", "Máquina", "Descripción", "Evidencia Final", "Qué Inicial", "Dónde", "Quién Abrió", "Evidencia Inicial", "Apertura TS", "Cierre TS"]); sh.getRange("A1:N1").setFontWeight("bold").setBackground("#009B4A").setFontColor("white"); }
    if (!ss.getSheetByName(EHS_SHEET_NAME)) { let sh = ss.insertSheet(EHS_SHEET_NAME); sh.appendRow(["ID EHS", "Folio Validado", "Fecha", "EHS Inspector", "No. Emp", "Foto Verificación EHS", "Evidencia Validación"]); sh.getRange("A1:G1").setFontWeight("bold").setBackground("#dc2626").setFontColor("white"); }
    getOrCreateRefSheet(); 
    getOrCreateOEESheet(); 
    getOrCreatePersonalSheet(); // 👈 INICIALIZA HOJA Y PRECARGA LOS 19 EMPLEADOS AUTOMÁTICAMENTE
    getOrCreateEquiposSheet(); // 👈 AGREGAR ESTA LÍNEA (Inicializa Log_Equipos con los 438 equipos)
    getOrCreateProdSheet();
    if (!ss.getSheetByName(SHEET_NAME)) { let sh = ss.insertSheet(SHEET_NAME); sh.appendRow(["Métrica", "Valor"]); sh.appendRow(["Abiertos", 0]); sh.appendRow(["Cerrados", 0]); sh.appendRow(["Última Act", "Sin datos"]); sh.appendRow(["EHS Pendientes", 0]); sh.appendRow(["EHS Validados", 0]); sh.appendRow(["Ref Solic.", 0]); sh.appendRow(["Ref Entregadas", 0]); }
  } catch(e) {}
}


function getOrCreateRefSheet() {
  const ss = getDB(); 
  let refSheet = ss.getSheetByName(REF_SHEET_NAME);
  if (!refSheet) {
    refSheet = ss.insertSheet(REF_SHEET_NAME);
    refSheet.appendRow([
      "FolioRef", "FolioMttoVinc", "Area", "Linea", "Maquina", "IdMaquina", 
      "FotoFalla", "ParteSolicitada", "EstatusGeneral", "Cotizacion", "NumReq", 
      "Autorizacion", "OrdenCompra", "RecepcionMaterial", "FotoEvidenciaMaterial", 
      "FechaRegistro", "UrlCotizacion", "UrlAprobacion", "FechaEntrega", 
      "UrlReqEvidencia", "UrlOCEvidencia", "DirAprobacion", "UrlDirAprobacion", 
      "TsCotizacion", "TsRequisicion", "TsAutorizacion", "TsDirAprobacion", 
      "TsOrdenCompra", "TsEntregaMaterial"
    ]);
    refSheet.getRange("A1:AC1").setFontWeight("bold").setBackground("#f59e0b").setFontColor("black");
  }
  return refSheet;
}

function getOrCreateOEESheet() {
  const ss = getDB(); let oeeSheet = ss.getSheetByName(OEE_SHEET_NAME);
  if (!oeeSheet) {
    oeeSheet = ss.insertSheet(OEE_SHEET_NAME);
    oeeSheet.appendRow(["FechaRegistro", "TiempoPlanificadoHoras", "PiezasTeoricas", "PiezasReales", "PiezasBuenas", "DisponibilidadPct", "PerformancePct", "CalidadPct", "OEEPct"]);
    oeeSheet.getRange("A1:I1").setFontWeight("bold").setBackground("#eab308").setFontColor("black");
    oeeSheet.appendRow([new Date().toLocaleString('es-MX', {hour12:false}), 168, 10000, 9500, 9300, 90.0, 95.0, 97.9, 83.7]);
  }
  return oeeSheet;
}

function obtenerSiguienteFolioRef() {
  let refSheet = getOrCreateRefSheet(); let maxNum = 0;
  if (refSheet.getLastRow() > 1) { let vals = refSheet.getRange(2, 1, refSheet.getLastRow() - 1, 1).getValues(); vals.forEach(r => { let match = String(r[0]).match(/\d+/); if (match) { let num = parseInt(match[0], 10); if (num > maxNum) maxNum = num; } }); }
  return "REF-" + String(maxNum + 1).padStart(4, '0');
}

function obtenerSiguienteFolio() {
  const ss = getDB(); let maxNum = 0; let sheetsToCheck = [OPEN_SHEET_NAME, LOG_SHEET_NAME];
  sheetsToCheck.forEach(sheetName => { let sh = ss.getSheetByName(sheetName); if (sh && sh.getLastRow() > 1) { let vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues(); vals.forEach(r => { let match = String(r[0]).match(/\d+/); if (match) { let num = parseInt(match[0], 10); if (num > maxNum) maxNum = num; } }); } });
  return "FOLIO-" + String(maxNum + 1).padStart(4, '0');
}

function crearBlobSeguro(base64Data, contentType, filename) {
  try { var cleanBase64 = base64Data || ""; if (cleanBase64.indexOf(',') !== -1) { cleanBase64 = cleanBase64.split(',')[1]; } cleanBase64 = cleanBase64.replace(/[^A-Za-z0-9\+\/\=]/g, ""); if (!cleanBase64) return null; var decoded = Utilities.base64Decode(cleanBase64); return Utilities.newBlob(decoded, contentType || "application/octet-stream", filename || "documento"); } catch(e) { return null; }
}

function enviarCorreoCorporativo(subject, titulo, mensajesHtml, folioRef, area, linea, maq, parte, folioMtto, urlAdjunto1, urlAdjunto2) {
  try {
    let tablaInfo = folioRef !== "N/A" || parte !== "N/A" ? `<table style="width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px;"><tr><td style="padding: 10px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: bold; width: 35%;">Folio Refacción:</td><td style="padding: 10px; border: 1px solid #e2e8f0; font-weight:bold; color:#009B4A;">${folioRef}</td></tr><tr><td style="padding: 10px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: bold;">Folio Mantenimiento:</td><td style="padding: 10px; border: 1px solid #e2e8f0;">${folioMtto || "Pendiente"}</td></tr><tr><td style="padding: 10px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: bold;">Máquina:</td><td style="padding: 10px; border: 1px solid #e2e8f0; font-weight:bold;">${maq}</td></tr><tr><td style="padding: 10px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: bold;">Parte Solicitada:</td><td style="padding: 10px; border: 1px solid #e2e8f0; color: #dc2626; font-weight: bold;">${parte}</td></tr></table>` : `<table style="width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px;"><tr><td style="padding: 10px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: bold; width: 35%;">Folio Mantenimiento:</td><td style="padding: 10px; border: 1px solid #e2e8f0; font-weight:bold; color:#009B4A;">${folioMtto}</td></tr><tr><td style="padding: 10px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: bold;">Máquina:</td><td style="padding: 10px; border: 1px solid #e2e8f0; font-weight:bold;">${maq}</td></tr></table>`;
    let htmlBody = `<div style="font-family: 'Segoe UI', Arial, sans-serif; padding: 30px; background-color: #f4f7f6; color: #333;"><div style="max-width: 600px; margin: 0 auto; background: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.05); border-top: 6px solid #009B4A;"><div style="padding: 25px; text-align: center; border-bottom: 1px solid #e2e8f0;"><img src="https://www.nidec.com/files/user/www-nidec-com/corporate/about/slogan/img/img_slogan_01.png" style="height: 50px;"></div><div style="padding: 30px;"><h2 style="color: #009B4A; margin-top: 0; text-align: center;">${titulo}</h2><div style="font-size: 15px; color: #475569;">${mensajesHtml}</div>${tablaInfo}</div></div></div>`;
    let adjuntos = []; let urls = [urlAdjunto1, urlAdjunto2]; urls.forEach(url => { if (url && url.includes("drive.google.com")) { let match = url.match(/[-\w]{25,}/); if (match && match[0]) { try { let file = DriveApp.getFileById(match[0]); adjuntos.push(file.getBlob()); } catch(err) {} } } });
    if (adjuntos.length > 0) { MailApp.sendEmail({ to: CORREO_COMPRAS, subject: subject, htmlBody: htmlBody, attachments: adjuntos }); } else { MailApp.sendEmail({ to: CORREO_COMPRAS, subject: subject, htmlBody: htmlBody }); }
  } catch(e) {}
}

function consultarIA(preguntaUsuario) {
  try {
    if (!preguntaUsuario || String(preguntaUsuario).trim() === "") return "🤖 <strong>SekaIA:</strong> ¿En qué te puedo ayudar hoy?";
    let rawQ = String(preguntaUsuario).trim(); let q = rawQ.toLowerCase(); const ss = getDB();
    let openSh = ss.getSheetByName(OPEN_SHEET_NAME); let logSh = ss.getSheetByName(LOG_SHEET_NAME); let ehsSh = ss.getSheetByName(EHS_SHEET_NAME); let refSh = getOrCreateRefSheet();
    let openData = openSh && openSh.getLastRow() > 1 ? openSh.getDataRange().getValues() : []; let logData = logSh && logSh.getLastRow() > 1 ? logSh.getDataRange().getValues() : []; let ehsData = ehsSh && ehsSh.getLastRow() > 1 ? ehsSh.getDataRange().getValues() : []; let refData = refSh && refSh.getLastRow() > 1 ? refSh.getDataRange().getValues() : [];

    if (q.includes("abrieron") || q.includes("cerraron") || q.includes("abierto") || q.includes("cerrado") || q.includes("cuanto") || q.includes("cuánto")) {
      let hoy = new Date(); let abHoyList = []; let ceHoyList = [];
      for(let i = 1; i < openData.length; i++) { let d = parseAnyDate(openData[i][10] || openData[i][1]); if (d && d.getDate() === hoy.getDate() && d.getMonth() === hoy.getMonth() && d.getFullYear() === hoy.getFullYear()) { let p = parseDonde(openData[i][4]); abHoyList.push({ folio: openData[i][0], maquina: p.maquina, idMaq: p.idMaquina, que: openData[i][2], quien: openData[i][6] }); } }
      for(let i = 1; i < logData.length; i++) { let d = parseAnyDate(logData[i][13] || logData[i][1]); if (d && d.getDate() === hoy.getDate() && d.getMonth() === hoy.getMonth() && d.getFullYear() === hoy.getFullYear()) { let p = parseDonde(logData[i][9] || logData[i][5]); ceHoyList.push({ folio: logData[i][0], maquina: p.maquina || logData[i][5], desc: logData[i][6], tec: logData[i][4] }); } }
      if (q.includes("abrieron") || q.includes("abierto")) { let html = `🤖 <strong>SekaIA - Mantenimientos Abiertos Hoy (${hoy.toLocaleDateString('es-MX')}):</strong><br><br>Se han abierto <b style="color:#3b82f6;">${abHoyList.length} mantenimientos</b> hoy.<br>`; if (abHoyList.length > 0) { html += `<br>📋 <strong>Desglose de Aperturas Hoy:</strong><br>`; abHoyList.forEach(item => { html += `&nbsp;&nbsp;• <b>${item.folio}</b>: "${item.que}" | Equipo: <b>${item.maquina} (${item.idMaq})</b> (Solicitó: ${item.quien})<br>`; }); } return html; }
      if (q.includes("cerraron") || q.includes("cerrado")) { let html = `🤖 <strong>SekaIA - Mantenimientos Cerrados Hoy (${hoy.toLocaleDateString('es-MX')}):</strong><br><br>Se han cerrado <b style="color:#009B4A;">${ceHoyList.length} mantenimientos</b> hoy.<br>`; if (ceHoyList.length > 0) { html += `<br>📋 <strong>Desglose de Cierres Hoy:</strong><br>`; ceHoyList.forEach(item => { html += `&nbsp;&nbsp;• <b>${item.folio}</b>: "${item.desc}" | Equipo: <b>${item.maquina}</b> (Técnico: ${item.tec})<br>`; }); } return html; }
    }

    const stopWords = ["cuantas", "cuantos", "cuanta", "cuanto", "estan", "pendientes", "y", "de", "que", "maquinas", "maquina", "dime", "ver", "los", "las", "un", "una", "por", "favor", "hay", "tengo", "cuales", "cual", "en", "el", "la", "del"];
    let words = q.replace(/[^a-z0-9áéíóúñ\-\s]/gi, " ").split(/\s+/).filter(w => w.length > 2); let keywordsCore = words.filter(w => !stopWords.includes(w));
    if (keywordsCore.length > 0 && !q.match(/(mre[-\s]?\d+|cleater|dryer|transferencia)/i)) {
      let partMatchesRef = []; let partMatchesOpen = [];
      for (let i = 1; i < refData.length; i++) { let r = refData[i]; let folioR = String(r[0] || "").trim(); if (!folioR || folioR === "N/A") continue; let parte = String(r[7] || "").toLowerCase(); let estatusGen = String(r[8] || "").toUpperCase(); let recepcionMat = String(r[13] || "").trim(); let finalizado = (recepcionMat.includes("Completo") || estatusGen.includes("PROGRAMAR") || estatusGen.includes("UTILIZADAS")); let matchCore = keywordsCore.some(k => parte.includes(k) || k.includes(parte) || (k.length > 4 && parte.includes(k.substring(0, k.length - 1)))); if (matchCore) { partMatchesRef.push({ folio: folioR, parte: r[7], maquina: `${r[4]} (${r[5] || 'Sin ID'})`, areaLinea: `${r[2]} / Línea ${r[3]}`, estatus: estatusGen, recepcion: recepcionMat, fecha: r[15], finalizado: finalizado }); } }
      for (let i = 1; i < openData.length; i++) { let r = openData[i]; let queText = String(r[2] || "").toLowerCase(); let refInfo = String(r[8] || "").toLowerCase(); let matchCore = keywordsCore.some(k => queText.includes(k) || refInfo.includes(k)); if (matchCore) { let dInfo = parseDonde(r[4]); partMatchesOpen.push({ folio: r[0], falla: r[2], maquina: `${dInfo.maquina} (${dInfo.idMaquina})`, areaLinea: `${dInfo.area} / Línea ${dInfo.linea}`, quien: r[6], fecha: r[1] }); } }
      if (partMatchesRef.length > 0 || partMatchesOpen.length > 0) {
        let termLabel = keywordsCore.join(", ").toUpperCase(); let pendRef = partMatchesRef.filter(x => !x.finalizado); let compRef = partMatchesRef.filter(x => x.finalizado);
        let html = `🤖 <strong>SekaIA - Reporte de Componentes para "${termLabel}":</strong><br><br>`;
        if (pendRef.length > 0) { html += `🚨 <strong style="color:#d97706;">Refacciones Pendientes (${pendRef.length}):</strong><br>`; pendRef.forEach(p => { html += `&nbsp;&nbsp;📦 <b>${p.folio}</b>: ${p.parte}<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;📍 <b>Máquina:</b> ${p.maquina} [${p.areaLinea}]<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;⏳ <b>Estatus:</b> <span style="color:#8b5cf6; font-weight:bold;">${p.estatus}</span> (${p.recepcion})<br><br>`; }); }
        if (partMatchesOpen.length > 0) { html += `🛠️ <strong style="color:#dc2626;">Fallas Abiertas Relacionadas (${partMatchesOpen.length}):</strong><br>`; partMatchesOpen.forEach(o => { html += `&nbsp;&nbsp;🔧 <b>${o.folio}</b>: "${o.falla}"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;📍 <b>Máquina:</b> ${o.maquina} [${o.areaLinea}] (Reportó: ${o.quien})<br><br>`; }); }
        if (compRef.length > 0) { html += `✅ <strong style="color:#009B4A;">Refacciones Entregadas / Listas (${compRef.length}):</strong><br>`; compRef.forEach(c => { html += `&nbsp;&nbsp;✔ <b>${c.folio}</b>: ${c.parte} para <b>${c.maquina}</b><br>`; }); }
        return html;
      }
    }

    let matchMaq = q.match(/(mre[-\s]?\d+|cleater|dryer|transferencia|[a-z0-9]{3,}-\d+)/i);
    if (matchMaq) {
      let target = matchMaq[0].toUpperCase().replace(/\s+/g, ""); let targetClean = target.replace("-", ""); let logMatches = []; let totalDowntimeMs = 0;
      for(let i = 1; i < logData.length; i++) { let mName = String(logData[i][5] || "").toUpperCase(); let dInfo = String(logData[i][9] || "").toUpperCase(); if (mName.replace("-","").includes(targetClean) || dInfo.replace("-","").includes(targetClean)) { let tsA = parseInt(logData[i][12]); let tsC = parseInt(logData[i][13]); let diff = (tsA && tsC && tsC > tsA) ? (tsC - tsA) : 0; totalDowntimeMs += diff; logMatches.push({ folio: logData[i][0], fecha: logData[i][1], tec: logData[i][4], sup: logData[i][3], desc: logData[i][6], downtimeMins: Math.round(diff / 60000) }); } }
      let openMatches = []; for(let i = 1; i < openData.length; i++) { let dInfo = String(openData[i][4] || "").toUpperCase(); if (dInfo.replace("-","").includes(targetClean)) { openMatches.push({ folio: openData[i][0], falla: openData[i][2], quien: openData[i][6], fecha: openData[i][1] }); } }
      let refMatches = []; for(let i = 1; i < refData.length; i++) { let mName = String(refData[i][4] || "").toUpperCase(); let mId = String(refData[i][5] || "").toUpperCase(); if (mName.replace("-","").includes(targetClean) || mId.replace("-","").includes(targetClean)) { refMatches.push({ folioRef: refData[i][0], parte: refData[i][7], estatus: refData[i][8], fecha: refData[i][15], recepcion: refData[i][13] }); } }
      if (logMatches.length === 0 && openMatches.length === 0 && refMatches.length === 0) { return `🤖 <strong>SekaIA:</strong> No encontré registros para el equipo <strong>${target}</strong>.`; }
      let downtimeHrs = (totalDowntimeMs / 3600000).toFixed(1); let html = `🤖 <strong>SekaIA - Expediente Técnico de ${target}:</strong><br><br>📊 <strong>Métricas Clave:</strong><br>• Mantenimientos cerrados: <b>${logMatches.length}</b><br>• Fallas activas/abiertas: <b style="color:${openMatches.length > 0 ? '#dc2626' : '#009B4A'};">${openMatches.length}</b><br>• Tiempo total fuera de servicio: <b>${downtimeHrs} hrs</b> (${Math.round(totalDowntimeMs/60000)} mins)<br>• Refacciones requeridas: <b>${refMatches.length}</b><br><br>`;
      if (openMatches.length > 0) { html += `🚨 <strong style="color:#dc2626;">Reportes Abiertos Actuales:</strong><br>`; openMatches.forEach(o => { html += `&nbsp;&nbsp;• <b>${o.folio}</b>: "${o.falla}" (Reportó: ${o.quien} | ${o.fecha})<br>`; }); html += `<br>`; }
      if (refMatches.length > 0) { html += `⚙️ <strong>Historial de Refacciones:</strong><br>`; refMatches.slice(-5).reverse().forEach(r => { html += `&nbsp;&nbsp;• <b>${r.folioRef}</b>: ${r.parte} | Estatus: <b>${r.estatus}</b> (${r.recepcion})<br>`; }); html += `<br>`; }
      if (logMatches.length > 0) { html += `🛠️ <strong>Últimas Reparaciones Exitosas:</strong><br>`; logMatches.slice(-3).reverse().forEach(l => { html += `&nbsp;&nbsp;• <b>${l.folio}</b> (${l.fecha}): ${l.desc} | Técnico: <b>${l.tec}</b> (${l.downtimeMins} min)<br>`; }); }
      return html;
    }

    return `🤖 <strong>SekaIA:</strong> Hola. Soy tu asistente de IA de Nidec Motor y me llamo SekaIA. ¿En qué te puedo ayudar hoy?`;
  } catch(e) { return "🤖 <strong>SekaIA Error:</strong> " + e.toString(); }
}

function getKpisWCMData() {
  try {
    checkAndInitAllSheets();
    const ss = getDB();
    let openSheet = ss.getSheetByName(OPEN_SHEET_NAME);
    let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    let nowTs = Date.now();
    let machineStats = {};
    
    let totalAbiertosReal = (openSheet && openSheet.getLastRow() > 1) ? openSheet.getLastRow() - 1 : 0;
    let totalCerradosReal = (logSheet && logSheet.getLastRow() > 1) ? logSheet.getLastRow() - 1 : 0;

    if (logSheet && logSheet.getLastRow() > 1) {
      let data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();
      data.forEach(r => {
        // LECTURA INTELIGENTE DE FECHAS Y TIMESTAMPS
        let dCierre = parseAnyDate(r[13]) || parseAnyDate(r[1]);
        let dApertura = parseAnyDate(r[12]) || dCierre;

        let tsApertura = dApertura ? dApertura.getTime() : null;
        let tsCierre = dCierre ? dCierre.getTime() : null;
        
        let maquinaName = String(r[5] || "MÁQUINA GENERAL").trim();
        let dInfoStr = String(r[9] || ""); let lineaExtr = "N/A"; let idMaqExtr = "N/A";
        if(dInfoStr.includes("||")) { let p = dInfoStr.split("||").map(s => s.trim()); lineaExtr = p[1] || "N/A"; maquinaName = p[2] || maquinaName; idMaqExtr = p[3] || "N/A"; }

        let cleanMaq = maquinaName.replace(/['"]/g, ""); let cleanLin = lineaExtr.replace(/['"]/g, ""); let cleanId = idMaqExtr.replace(/['"]/g, "");
        let key = cleanMaq + "||" + cleanLin + "||" + cleanId;

        if (!machineStats[key]) {
            machineStats[key] = { maquina: cleanMaq, linea: cleanLin, idMaquina: cleanId, historicalDowntimeMins: 0, historicalFailures: 0, hasActiveOpen: false, openStartTs: null, lastCloseTs: null, firstEventTs: nowTs };
        }

        if (tsApertura && tsCierre && !isNaN(tsApertura) && !isNaN(tsCierre)) {
          if (tsApertura < machineStats[key].firstEventTs) machineStats[key].firstEventTs = tsApertura;
          let validCierre = tsCierre >= tsApertura ? tsCierre : tsApertura + 60000; 
          let diffMins = (validCierre - tsApertura) / 60000; 
          machineStats[key].historicalDowntimeMins += diffMins; 
          machineStats[key].historicalFailures += 1;
          if (!machineStats[key].lastCloseTs || validCierre > machineStats[key].lastCloseTs) {
            machineStats[key].lastCloseTs = validCierre;
          }
        }
      });
    }

    if (openSheet && openSheet.getLastRow() > 1) {
      let data = openSheet.getRange(2, 1, openSheet.getLastRow() - 1, openSheet.getLastColumn()).getValues();
      data.forEach(r => {
        let dApertura = parseAnyDate(r[10]) || parseAnyDate(r[1]);
        let tsApertura = dApertura ? dApertura.getTime() : nowTs;
        
        let dInfoStr = String(r[4] || "");
        let maquinaName = "MÁQUINA GENERAL"; let lineaExtr = "N/A"; let idMaqExtr = "N/A";
        if (dInfoStr.includes("||")) { let p = dInfoStr.split("||").map(s => s.trim()); lineaExtr = p[1] || "N/A"; maquinaName = p[2] || "MÁQUINA GENERAL"; idMaqExtr = p[3] || "N/A"; }

        let cleanMaq = maquinaName.replace(/['"]/g, ""); let cleanLin = lineaExtr.replace(/['"]/g, ""); let cleanId = idMaqExtr.replace(/['"]/g, "");
        let key = cleanMaq + "||" + cleanLin + "||" + cleanId;

        if (!machineStats[key]) {
            machineStats[key] = { maquina: cleanMaq, linea: cleanLin, idMaquina: cleanId, historicalDowntimeMins: 0, historicalFailures: 0, hasActiveOpen: false, openStartTs: null, lastCloseTs: null, firstEventTs: tsApertura };
        }

        if (tsApertura < machineStats[key].firstEventTs) machineStats[key].firstEventTs = tsApertura;
        machineStats[key].hasActiveOpen = true;
        if (!machineStats[key].openStartTs || tsApertura < machineStats[key].openStartTs) {
          machineStats[key].openStartTs = tsApertura;
        }
      });
    }

    let oeeSheet = getOrCreateOEESheet(); let lastRowOEE = oeeSheet ? oeeSheet.getLastRow() : 0;
    let oeeData = { plannedHours: 168, targetUnits: 10000, realUnits: 9500, goodUnits: 9300, performancePct: 95.0, qualityPct: 97.9 };
    if (lastRowOEE > 1) {
      let row = oeeSheet.getRange(lastRowOEE, 1, 1, 9).getValues()[0];
      let tu = Number(row[2]) || 10000; let ru = Number(row[3]) || 9500; let gu = Number(row[4]) || 9300;
      oeeData.plannedHours = Number(row[1]) || 168;
      oeeData.targetUnits = tu; oeeData.realUnits = ru; oeeData.goodUnits = gu;
      oeeData.performancePct = tu > 0 ? (ru / tu) * 100 : 0; oeeData.qualityPct = ru > 0 ? (gu / ru) * 100 : 0;
    }

    let machineList = Object.keys(machineStats).map(k => { 
        let stat = machineStats[k];
        if (stat.firstEventTs >= nowTs && stat.historicalFailures === 0) { stat.firstEventTs = nowTs - (86400000 * 7); }
        return { 
          idMaq: stat.maquina, linea: stat.linea, idMaquina: stat.idMaquina, 
          historicalFailures: stat.historicalFailures || 0, historicalDowntimeMins: stat.historicalDowntimeMins || 0,
          hasActiveOpen: !!stat.hasActiveOpen, openStartTs: stat.openStartTs || 0, lastCloseTs: stat.lastCloseTs || stat.firstEventTs, firstEventTs: stat.firstEventTs || nowTs
        }; 
    });
    machineList.sort((a,b) => (b.hasActiveOpen ? 1 : 0) - (a.hasActiveOpen ? 1 : 0));

    let totalTareas = totalAbiertosReal + totalCerradosReal; 
    let compliancePct = totalTareas > 0 ? (totalCerradosReal / totalTareas) * 100 : 0;

    return { 
        success: true, oee: oeeData, machineList: machineList,
        scheduleCompliance: { pct: compliancePct.toFixed(1), cerrados: totalCerradosReal, total: totalTareas },
        serverNowTs: nowTs
    };
  } catch(e) { return { success: false, errorServer: "Error al calcular KPIs: " + e.toString() }; }
}

function getDashboardData() {
  try {
    checkAndInitAllSheets();
    var ss = getDB();
    var shAbiertos = ss.getSheetByName(OPEN_SHEET_NAME);
    var shCerrados = ss.getSheetByName(LOG_SHEET_NAME);
    var shEHS = ss.getSheetByName(EHS_SHEET_NAME);
    var shRef = ss.getSheetByName(REF_SHEET_NAME);
    var shPerm = ss.getSheetByName("Log_Permisos_EHS"); // 👈 NUEVO: Lee la pestaña de permisos

    var dataAbiertos = (shAbiertos && shAbiertos.getLastRow() > 1) ? shAbiertos.getDataRange().getValues() : [];
    var dataCerrados = (shCerrados && shCerrados.getLastRow() > 1) ? shCerrados.getDataRange().getValues() : [];
    var dataEHS = (shEHS && shEHS.getLastRow() > 1) ? shEHS.getDataRange().getValues() : [];
    var dataRef = (shRef && shRef.getLastRow() > 1) ? shRef.getDataRange().getValues() : [];
    var dataPerm = (shPerm && shPerm.getLastRow() > 1) ? shPerm.getDataRange().getValues() : []; // 👈 NUEVO: Extrae los permisos

    var numAbiertos = Math.max(0, dataAbiertos.length - 1);
    var numCerrados = Math.max(0, dataCerrados.length - 1);
    var totalEHSValidados = Math.max(0, dataEHS.length - 1);

    // 👈 NUEVO: Cuenta cuántos permisos están actualmente pendientes (SOLICITADO o ESPERANDO CIERRE EHS)
    var numPermisosPendientes = 0;
    if (dataPerm.length > 1) {
      for (var p = 1; p < dataPerm.length; p++) {
        var estP = String(dataPerm[p][7] || "").toUpperCase().trim();
        if (estP === "SOLICITADO" || estP === "ESPERANDO CIERRE EHS") {
          numPermisosPendientes++;
        }
      }
    }

    // 👈 MODIFICADO: Suma los pendientes de MTTO + los permisos pendientes EHS
    var totalEHSPendientes = Math.max(0, numCerrados - totalEHSValidados) + numPermisosPendientes;

    var numRefSolicitadas = Math.max(0, dataRef.length - 1);
    var numRefEntregadas = 0;
    if (dataRef.length > 1) {
      for (var i = 1; i < dataRef.length; i++) {
        var estatus = String(dataRef[i][8] || "").toUpperCase();
        var recMat = String(dataRef[i][13] || "").toUpperCase();
        if (estatus.indexOf("LISTA") !== -1 || estatus.indexOf("COMPLETO") !== -1 || estatus.indexOf("UTILIZADAS") !== -1 || estatus.indexOf("PROGRAMAR") !== -1 || recMat.indexOf("COMPLETO") !== -1) {
          numRefEntregadas++;
        }
      }
    }

    var diasSemana = ["DOM", "LUN", "MAR", "MIE", "JUE", "VIE", "SAB"];
    var hoy = new Date();
    var mapaFechas = {};
    var clavesOrdenadas = [];

    for (var d = 6; d >= 0; d--) {
      var f = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - d);
      var clave = f.getFullYear() + "-" + (f.getMonth() + 1) + "-" + f.getDate();
      var label = diasSemana[f.getDay()] + " " + (f.getDate() < 10 ? "0" : "") + f.getDate();
      
      mapaFechas[clave] = {
        label: label,
        mttoAbiertos: 0,
        mttoCerrados: 0,
        ehsCerrados: 0,
        ehsPendientesGenerados: 0, // 👈 NUEVO: Acumulador exclusivo de EHS por día
        refSolicitadas: 0,
        refEntregadas: 0
      };
      clavesOrdenadas.push(clave);
    }

    // 1. Mantenimientos Abiertos
    for (var i = 1; i < dataAbiertos.length; i++) {
      var fObj = parseAnyDate(dataAbiertos[i][10] || dataAbiertos[i][1]);
      if (fObj) {
        var c = fObj.getFullYear() + "-" + (fObj.getMonth() + 1) + "-" + fObj.getDate();
        if (mapaFechas[c]) mapaFechas[c].mttoAbiertos++;
      }
    }

    // 2. Mantenimientos Cerrados (suman a la tendencia EHS)
    for (var i = 1; i < dataCerrados.length; i++) {
      var fObj = parseAnyDate(dataCerrados[i][13] || dataCerrados[i][1]);
      if (fObj) {
        var c = fObj.getFullYear() + "-" + (fObj.getMonth() + 1) + "-" + fObj.getDate();
        if (mapaFechas[c]) {
          mapaFechas[c].mttoCerrados++;
          mapaFechas[c].ehsPendientesGenerados++; // 👈 Registra MTTO cerrado pendiente de auditoría
        }
      }
    }

    // 👈 NUEVO: Permisos EHS Solicitados (suman a la tendencia diaria de pendientes EHS)
    for (var i = 1; i < dataPerm.length; i++) {
      var fObj = parseAnyDate(dataPerm[i][1]); // Fecha de Registro del Permiso
      if (fObj) {
        var c = fObj.getFullYear() + "-" + (fObj.getMonth() + 1) + "-" + fObj.getDate();
        if (mapaFechas[c]) {
          mapaFechas[c].ehsPendientesGenerados++; // 👈 Registra Permiso solicitado en la tendencia
        }
      }
    }

    // 3. Validaciones EHS auditadas
    for (var i = 1; i < dataEHS.length; i++) {
      var fObj = parseAnyDate(dataEHS[i][2]);
      if (fObj) {
        var c = fObj.getFullYear() + "-" + (fObj.getMonth() + 1) + "-" + fObj.getDate();
        if (mapaFechas[c]) mapaFechas[c].ehsCerrados++;
      }
    }

    // 4. Refacciones
    for (var i = 1; i < dataRef.length; i++) {
      var fObjReq = parseAnyDate(dataRef[i][15]);
      if (fObjReq) {
        var cReq = fObjReq.getFullYear() + "-" + (fObjReq.getMonth() + 1) + "-" + fObjReq.getDate();
        if (mapaFechas[cReq]) mapaFechas[cReq].refSolicitadas++;
      }

      var est = String(dataRef[i][8] || "").toUpperCase();
      var rec = String(dataRef[i][13] || "").toUpperCase();
      var isEntregado = (est.indexOf("LISTA") !== -1 || est.indexOf("COMPLETO") !== -1 || est.indexOf("UTILIZADAS") !== -1 || est.indexOf("PROGRAMAR") !== -1 || rec.indexOf("COMPLETO") !== -1);
      
      if (isEntregado) {
        var fechaEntRaw = dataRef[i][18] || dataRef[i][15]; 
        var fObjEnt = parseAnyDate(fechaEntRaw);
        if (fObjEnt) {
          var cEnt = fObjEnt.getFullYear() + "-" + (fObjEnt.getMonth() + 1) + "-" + fObjEnt.getDate();
          if (mapaFechas[cEnt]) mapaFechas[cEnt].refEntregadas++;
        }
      }
    }

    var trendMtto = { categories: [], abiertos: [], cerrados: [] };
    var trendEHS = { categories: [], abiertos: [], cerrados: [] };
    var trendRef = { categories: [], abiertos: [], cerrados: [] };

    clavesOrdenadas.forEach(function(k) {
      var o = mapaFechas[k];
      
      trendMtto.categories.push(o.label);
      trendMtto.abiertos.push(o.mttoAbiertos);
      trendMtto.cerrados.push(o.mttoCerrados);

      trendEHS.categories.push(o.label);
      trendEHS.abiertos.push(o.ehsPendientesGenerados); // 👈 MODIFICADO: Muestra la suma real (MTTO + Permisos)
      trendEHS.cerrados.push(o.ehsCerrados);

      trendRef.categories.push(o.label);
      trendRef.abiertos.push(o.refSolicitadas);
      trendRef.cerrados.push(o.refEntregadas);
    });

    return {
      programados: numAbiertos,
      realizados: numCerrados,
      ehsAbiertos: totalEHSPendientes, // 👈 Ahora enviará 6 (1 MTTO + 5 Permisos)
      ehsValidados: totalEHSValidados,
      refSolicitadas: numRefSolicitadas,
      refEntregadas: numRefEntregadas,
      trendMtto: trendMtto,
      trendEHS: trendEHS,
      trendRef: trendRef
    };

  } catch (e) {
    return { errorServer: e.message };
  }
}

function abrirMantenimiento5W2H(datos) {
  try {
    // 🚨 MODIFICACIÓN RED FLAG: ANTEPONE LA ALERTA A LA DESCRIPCIÓN Y REFACCIONES
    if (datos.redFlag) {
      datos.que = "🚨 🚩 [RED FLAG] " + datos.que;
      if (datos.refDesc) { datos.refDesc = "🚨 🚩 [RED FLAG] " + datos.refDesc; }
    }

    const ss = getDB(); let openSheet = ss.getSheetByName(OPEN_SHEET_NAME); let folders = DriveApp.getFoldersByName(FOLDER_NAME); let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
let urlFoto = "";
    // 👇 ESTA LÍNEA PROCESA LA IMAGEN SI EL USUARIO LA SUBIÓ
    if (datos.foto && datos.foto.base64) { let blob = crearBlobSeguro(datos.foto.base64, datos.foto.tipo, datos.foto.nombre); if (blob) { let archivoDrive = folder.createFile(blob); try { archivoDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(err) {} urlFoto = archivoDrive.getUrl(); } }
    
    // 👇 ESTAS LÍNEAS CREAN EL FOLIO Y LA FECHA
    let folio = obtenerSiguienteFolio(); let tsNow = new Date().getTime(); let fechaHoyString = new Date(tsNow).toLocaleString('es-MX', { hour12: false });
    
    let dondeEstructurado = [datos.area, datos.linea, datos.maquina, datos.idMaquina, datos.turno].join(" || ");

    openSheet.appendRow([folio, fechaHoyString, datos.que, datos.porque, dondeEstructurado, fechaHoyString, datos.quien, datos.como, datos.requiereRef ? `SÍ (${datos.refDesc})` : "NO", urlFoto, tsNow]);
    
    if (datos.requiereRef) { 
      let refSheet = getOrCreateRefSheet(); let folioRef = obtenerSiguienteFolioRef(); 
      refSheet.appendRow([folioRef, folio, datos.area, datos.linea, datos.maquina, datos.idMaquina, urlFoto, datos.refDesc, "SOLICITUD DE REFACCIONES", "No Iniciado", "", "No Autorizado", "No Recibido", "Pendiente", "", fechaHoyString, "", "", ""]); 
      enviarCorreoCorporativo(`NUEVA REQUISICIÓN DE MATERIAL: ${folioRef}`, "SOLICITUD INICIADA", `<p>Se ha abierto un nuevo Mantenimiento que <strong>requiere refacciones</strong>.</p>`, folioRef, datos.area, datos.linea, datos.maquina, datos.refDesc, folio, urlFoto, ""); 
    } else { 
      enviarCorreoCorporativo(`NUEVO MANTENIMIENTO: ${folio}`, "MANTENIMIENTO ABIERTO", `<p>Se ha abierto un nuevo Mantenimiento en el sistema EAM.</p>`, "N/A", datos.area, datos.linea, datos.maquina, "No requerida", folio, urlFoto, ""); 
    }
    return { success: true, message: `¡Ticket abierto exitosamente! ID: ${folio}` };
  } catch(error) { return { success: false, message: "Error: " + error.toString() }; }
}

function guardarMantenimientoRealizado(datos) {
  try {
    const ss = getDB(); let logSheet = ss.getSheetByName(LOG_SHEET_NAME); let cerradosData = logSheet.getDataRange().getValues();
    for (let j = 1; j < cerradosData.length; j++) { if (String(cerradosData[j][0]).trim() === String(datos.folio).trim()) return { success: true, message: `El ticket ya estaba cerrado.` }; }
    let openSheet = ss.getSheetByName(OPEN_SHEET_NAME); let qInicial = "Dato Antiguo", dInicial = datos.maquina, qAbrio = "N/A", evInicial = "", tsApertura = "";
    if (openSheet) { let openData = openSheet.getDataRange().getValues(); for (let i = 1; i < openData.length; i++) { if (String(openData[i][0]).trim() === String(datos.folio).trim()) { qInicial = openData[i][2] || "Dato Antiguo"; dInicial = openData[i][4] || datos.maquina; qAbrio = openData[i][6] || "N/A"; evInicial = openData[i][9] || ""; tsApertura = openData[i][10] || ""; openSheet.deleteRow(i + 1); break; } } }
    let folders = DriveApp.getFoldersByName(FOLDER_NAME); let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME); let urlFoto = "";
    if (datos.foto && datos.foto.base64) { let blob = crearBlobSeguro(datos.foto.base64, datos.foto.tipo, datos.foto.nombre); if (blob) { let archivoDrive = folder.createFile(blob); try { archivoDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(err) {} urlFoto = archivoDrive.getUrl(); } }
    let tsCierre = new Date().getTime(); let fechaCierreStr = new Date(tsCierre).toLocaleString('es-MX', { hour12: false });
    logSheet.appendRow([datos.folio, fechaCierreStr, datos.tipo, datos.supervisor, datos.tecnico, datos.maquina, datos.descripcion, urlFoto, qInicial, dInicial, qAbrio, evInicial, tsApertura, tsCierre]);
    
    if (datos.folioRefVinc) {
      let refSheet = getOrCreateRefSheet();
      if (refSheet.getLastRow() > 1) {
        let refData = refSheet.getRange(2, 1, refSheet.getLastRow() - 1, 16).getValues();
        for (let i = 0; i < refData.length; i++) {
          if (String(refData[i][0]).trim() === String(datos.folioRefVinc).trim()) {
            refSheet.getRange(i + 2, 2).setValue(datos.folio); 
            refSheet.getRange(i + 2, 9).setValue("REFACCIONES UTILIZADAS"); 
            break;
          }
        }
      }
    }

    let pDonde = dInicial.split(" || "); let areaMtto = pDonde[0] || "N/A"; let lineaMtto = pDonde[1] || "N/A";
    enviarCorreoCorporativo(`MANTENIMIENTO CERRADO (PENDIENTE EHS): ${datos.folio}`, "MANTENIMIENTO CERRADO", `<p>El mantenimiento ha sido finalizado.</p>`, "N/A", areaMtto, lineaMtto, datos.maquina, "N/A", datos.folio, urlFoto, "");
    return { success: true, message: `¡Ticket cerrado! Pasa a revisión de EHS.` };
  } catch(error) { return { success: false, message: "Error: " + error.toString() }; }
}

function validarMantenimientoEHS(datos) {
  try {
    const ss = getDB(); let ehsSheet = ss.getSheetByName(EHS_SHEET_NAME); let ehsData = ehsSheet.getDataRange().getValues();
    for (let k = 1; k < ehsData.length; k++) { if (String(ehsData[k][1]).trim() === String(datos.folio).trim()) return { success: true, message: `El ticket ya fue validado.` }; }
    let folders = DriveApp.getFoldersByName(FOLDER_NAME); let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME); let urlFotoCamara = "Sin foto cámara";
    if (datos.fotoCamara) { let blobCam = crearBlobSeguro(datos.fotoCamara, "image/jpeg", `EHS_Cam_${datos.folio}.jpg`); if (blobCam) { let fileCam = folder.createFile(blobCam); try { fileCam.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {} urlFotoCamara = fileCam.getUrl(); } }
    let urlEvidencia = "";
    if (datos.fotoEvidencia && datos.fotoEvidencia.base64) { let blob = crearBlobSeguro(datos.fotoEvidencia.base64, datos.fotoEvidencia.tipo, datos.fotoEvidencia.nombre); if (blob) { let archivoDrive = folder.createFile(blob); try { archivoDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(err) {} urlEvidencia = archivoDrive.getUrl(); } }
    let idEHS = "EHS-" + Date.now().toString().slice(-4);
    ehsSheet.appendRow([idEHS, datos.folio, new Date().toLocaleString('es-MX', {hour12:false}), datos.nombre, datos.empleado, urlFotoCamara, urlEvidencia]);
    enviarCorreoCorporativo(`MANTENIMIENTO VALIDADO EHS: ${datos.folio}`, "VALIDACIÓN EHS COMPLETADA", `<p>El ticket ha sido auditado por EHS.</p>`, "N/A", "N/A", "N/A", "N/A", "N/A", datos.folio, urlEvidencia, "");
    return { success: true, message: `¡Validación EHS registrada con éxito! (ID: ${idEHS})` };
  } catch(error) { return { success: false, message: "Error: " + error.toString() }; }
}
function getRefaccionesModuloCompleto() {
  try {
    let refSheet = getOrCreateRefSheet(); 
    let maxCols = refSheet.getLastColumn(); 
    let cards = []; 
    let counts = { solicitud: 0, cotizacion: 0, requisicion: 0, aprobacion: 0, dirAprobacion: 0, ordenCompra: 0, entregadasCompletas: 0, entregadasIncompletas: 0, pendientes: 0 };
    
    if (refSheet.getLastRow() > 1 && maxCols >= 16) {
      let data = refSheet.getRange(2, 1, refSheet.getLastRow() - 1, maxCols).getValues(); 
      let hoy = new Date();
      
      for (let i = 0; i < data.length; i++) {
        let r = data[i]; 
        let folioRef = String(r[0] || "").trim(); 
        if (!folioRef || folioRef === "N/A") continue;
        
        let estatusGen = String(r[8] || "SOLICITUD DE REFACCIONES").trim(); 
        let recepcionMat = String(r[13] || "Pendiente").trim(); 
        let dirAprob = maxCols >= 22 ? String(r[21] || "No Aprobado").trim() : "No Aprobado";
        let dRegistro = parseAnyDate(r[15]); 
        let dEntrega = parseAnyDate(r[18]);
        
        let finalizado = (recepcionMat.includes("Completo") || estatusGen.includes("PROGRAMAR") || estatusGen.includes("UTILIZADAS"));
        
        if (finalizado) counts.entregadasCompletas++; 
        else if (recepcionMat.includes("Incompleto") || estatusGen.includes("PENDIENTE")) { counts.entregadasIncompletas++; counts.pendientes++; } 
        else if (String(r[12] || "") === "Recibido" || estatusGen.includes("RECEPCION")) { counts.ordenCompra++; counts.pendientes++; } 
        else if (dirAprob === "Aprobado" || estatusGen.includes("ORDEN DE COMPRA")) { counts.dirAprobacion++; counts.pendientes++; }
        else if (String(r[11] || "") === "Autorizado" || estatusGen.includes("DIRECCION")) { counts.aprobacion++; counts.pendientes++; } 
        else if (String(r[10] || "").trim() !== "" || estatusGen.includes("AUTORIZACION")) { counts.requisicion++; counts.pendientes++; } 
        else if (String(r[9] || "") === "Done" || estatusGen.includes("REQUISICION")) { counts.cotizacion++; counts.pendientes++; } 
        else if (String(r[9] || "") === "In Process" || estatusGen.includes("COTIZACION")) { counts.cotizacion++; counts.pendientes++; } 
        else { counts.solicitud++; counts.pendientes++; }
        
        let diasTranscurridos = 0; 
        if(dRegistro) { 
          let endD = finalizado ? (dEntrega ? dEntrega : hoy) : hoy; 
          diasTranscurridos = Math.max(0, Math.floor((endD - dRegistro) / 86400000)); 
        }

        cards.push({ 
          rowIdx: i + 2, 
          folioRef: folioRef, 
          folioMtto: String(r[1] || "").trim(), 
          area: String(r[2] || "N/A"), 
          linea: String(r[3] || "N/A"), 
          maquina: String(r[4] || "N/A"), 
          idMaquina: String(r[5] || "N/A"), 
          fotoFalla: String(r[6] || ""), 
          parteSolicitada: String(r[7] || "Sin especificación"), 
          estatusGeneral: estatusGen, 
          cotizacion: String(r[9] || "No Iniciado").trim(), 
          numReq: String(r[10] || "").trim(), 
          autorizacion: String(r[11] || "No Autorizado").trim(), 
          ordenCompra: String(r[12] || "No Recibido").trim(), 
          recepcionMaterial: recepcionMat, 
          fotoEvidenciaMaterial: String(r[14] || ""), 
          fechaRegistro: String(r[15] || "N/A"), 
          diasTranscurridos: diasTranscurridos, 
          finalizado: finalizado, 
          urlCotizacion: maxCols >= 17 ? String(r[16] || "") : "", 
          urlAprobacion: maxCols >= 18 ? String(r[17] || "") : "", 
          urlReqEvidencia: maxCols >= 20 ? String(r[19] || "") : "", 
          urlOCEvidencia: maxCols >= 21 ? String(r[20] || "") : "",
          dirAprobacion: dirAprob,
          urlDirAprobacion: maxCols >= 23 ? String(r[22] || "") : "",
          tsRegistro: dRegistro ? dRegistro.getTime() : null, 
          tsCotizacion: maxCols >= 24 && r[23] ? parseAnyDate(r[23])?.getTime() || null : null,
          tsRequisicion: maxCols >= 25 && r[24] ? parseAnyDate(r[24])?.getTime() || null : null,
          tsAutorizacion: maxCols >= 26 && r[25] ? parseAnyDate(r[25])?.getTime() || null : null,
          tsDirAprobacion: maxCols >= 27 && r[26] ? parseAnyDate(r[26])?.getTime() || null : null,
          tsOrdenCompra: maxCols >= 28 && r[27] ? parseAnyDate(r[27])?.getTime() || null : null,
          tsEntrega: dEntrega ? dEntrega.getTime() : (maxCols >= 29 && r[28] ? parseAnyDate(r[28])?.getTime() || null : null)
        });
      }
    }
    return { success: true, counts: counts, cards: cards };
  } catch(e) { return { errorServer: "Error cargando refacciones: " + e.toString() }; }
}

// OBTENER REFACCIONES LISTAS DE PLANTA PARA "CLOSE MTTO" (CORREGIDO PARA INCLUIR TODAS LAS LISTAS)
function getRefaccionesListasParaProgramar() {
  try {
    let refSheet = getOrCreateRefSheet(); let listas = [];
    if (refSheet && refSheet.getLastRow() > 1) {
      let maxCols = refSheet.getLastColumn();
      let data = refSheet.getRange(2, 1, refSheet.getLastRow() - 1, maxCols).getValues();
      data.forEach(r => {
        let estatusGen = String(r[8] || "").trim().toUpperCase();
        let recepcionMat = String(r[13] || "").trim().toLowerCase();
        let folioMttoVinc = String(r[1] || "").trim();
        
        let isReady = (estatusGen.includes("PROGRAMAR") || estatusGen.includes("LISTA") || recepcionMat.includes("completo") || estatusGen.includes("UTILIZADAS"));
        
        if (isReady) {
          listas.push({
            folioRef: String(r[0]),
            folioMttoVinc: folioMttoVinc,
            descripcion: `${r[0]} - ${r[4]} (${r[5]}) - ${r[7]}`
          });
        }
      });
    }
    return listas;
  } catch(e) { return []; }
}

function actualizarAvanceRefaccionCard(datos) {
  try {
    let refSheet = getOrCreateRefSheet(); 
    let rowIdx = parseInt(datos.rowIdx, 10); 
    if (!rowIdx || rowIdx < 2) return { success: false, message: "Fila inválida." };

    let cotizacion = datos.cotizacion || "No Iniciado"; 
    let numReq = datos.numReq || ""; 
    let autorizacion = datos.autorizacion || "No Autorizado"; 
    let dirAprobacion = datos.dirAprobacion || "No Aprobado"; 
    let ordenCompra = datos.ordenCompra || "No Recibido"; 
    let recepcionMaterial = datos.recepcionMaterial || "Pendiente";

    let folders = DriveApp.getFoldersByName(FOLDER_NAME); 
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME); 
    
    let urlFotoIncompleto = "", urlPDFAut = "", urlPDFReq = "", urlPDFOC = "", urlPDFDirAut = "";
    let urlPDFCotArray = [];

    if (datos.pdfCotizaciones && datos.pdfCotizaciones.length > 0) {
      datos.pdfCotizaciones.forEach((pdfObj, idx) => {
        if (pdfObj && pdfObj.base64) {
          let blob = crearBlobSeguro(pdfObj.base64, pdfObj.tipo, pdfObj.nombre || ("Cotizacion_" + (idx+1) + ".pdf"));
          if (blob) {
            let fileDrive = folder.createFile(blob);
            try { fileDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
            urlPDFCotArray.push(fileDrive.getUrl());
          }
        }
      });
    }

    if (datos.pdfReqEvidencia && datos.pdfReqEvidencia.base64) {
      let blob = crearBlobSeguro(datos.pdfReqEvidencia.base64, datos.pdfReqEvidencia.tipo, datos.pdfReqEvidencia.nombre || "Req_Evidencia.pdf");
      if (blob) {
        let fileDrive = folder.createFile(blob);
        try { fileDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
        urlPDFReq = fileDrive.getUrl();
      }
    }

    if (datos.pdfAprobacion && datos.pdfAprobacion.base64) { 
      let blob = crearBlobSeguro(datos.pdfAprobacion.base64, datos.pdfAprobacion.tipo, datos.pdfAprobacion.nombre || "Aprobacion_Evidencia.pdf"); 
      if (blob) { 
        let fileDrive = folder.createFile(blob); 
        try { fileDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){} 
        urlPDFAut = fileDrive.getUrl(); 
      } 
    }

    if (datos.pdfDirAprobacion && datos.pdfDirAprobacion.base64) {
      let blob = crearBlobSeguro(datos.pdfDirAprobacion.base64, datos.pdfDirAprobacion.tipo, datos.pdfDirAprobacion.nombre || "Dir_Aprobacion_Evidencia.pdf");
      if (blob) {
        let fileDrive = folder.createFile(blob);
        try { fileDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
        urlPDFDirAut = fileDrive.getUrl();
      }
    }

    if (datos.pdfOCEvidencia && datos.pdfOCEvidencia.base64) {
      let blob = crearBlobSeguro(datos.pdfOCEvidencia.base64, datos.pdfOCEvidencia.tipo, datos.pdfOCEvidencia.nombre || "OC_Evidencia.pdf");
      if (blob) {
        let fileDrive = folder.createFile(blob);
        try { fileDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
        urlPDFOC = fileDrive.getUrl();
      }
    }

    if (datos.fotoEvidenciaMaterial && datos.fotoEvidenciaMaterial.base64) { 
      let blob = crearBlobSeguro(datos.fotoEvidenciaMaterial.base64, datos.fotoEvidenciaMaterial.tipo, datos.fotoEvidenciaMaterial.nombre || "Recepcion_Evidencia.jpg"); 
      if (blob) { 
        let fileDrive = folder.createFile(blob); 
        try { fileDrive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){} 
        urlFotoIncompleto = fileDrive.getUrl(); 
      } 
    }
    
    let tsNowStr = new Date().toLocaleString('es-MX', {hour12: false});
    let estatusGeneral = "SOLICITUD DE REFACCIONES"; 
    let finalizadoHoy = false; 
    let tituloMail = "", msjMail = "";

    if (recepcionMaterial === "Llegó Completo") { 
      estatusGeneral = "PROGRAMAR PARA MANTENIMIENTO"; finalizadoHoy = true; tituloMail = "MATERIAL RECIBIDO COMPLETO"; msjMail = "El material ha sido recibido."; 
    } else if (recepcionMaterial === "Llegó Incompleto") { 
      estatusGeneral = "PENDIENTE POR MATERIALES"; tituloMail = "MATERIAL RECIBIDO INCOMPLETO"; msjMail = "Entrega parcial del material."; 
    } else if (ordenCompra === "Recibido") { 
      estatusGeneral = "EN ESPERA DE RECEPCION DE MATERIAL"; tituloMail = "ACTUALIZACIÓN: ORDEN DE COMPRA RECIBIDA"; msjMail = "La OC ha sido procesada."; 
    } else if (dirAprobacion === "Aprobado") { 
      estatusGeneral = "EN PROCESO ORDEN DE COMPRA"; tituloMail = "ACTUALIZACIÓN: APROBACIÓN DIRECCIÓN COMPLETADA"; msjMail = "La Dirección ha autorizado la requisición."; 
    } else if (autorizacion === "Autorizado") { 
      estatusGeneral = "EN ESPERA DE APROBACION DIRECCION"; tituloMail = "ACTUALIZACIÓN: AUTORIZACIÓN GERENCIAL COMPLETADA"; msjMail = "Pasa a aprobación de Dirección."; 
    } else if (numReq && numReq.trim() !== "") { 
      estatusGeneral = "EN PROCESO DE AUTORIZACION"; tituloMail = "ACTUALIZACIÓN: REQUISICIÓN INICIADA"; msjMail = "Se ha generado REQ."; 
    } else if (cotizacion === "Done") { 
      estatusGeneral = "EN PROCESO DE REQUISICION"; tituloMail = "ACTUALIZACIÓN: COTIZACIÓN COMPLETADA"; msjMail = "Cotización lista."; 
    } else if (cotizacion === "In Process") { 
      estatusGeneral = "EN PROCESO DE COTIZACION"; tituloMail = "ACTUALIZACIÓN: COTIZACIÓN EN PROCESO"; msjMail = "Avance en la cotización."; 
    }
    
    refSheet.getRange(rowIdx, 9).setValue(estatusGeneral); 
    refSheet.getRange(rowIdx, 10).setValue(cotizacion); 
    refSheet.getRange(rowIdx, 11).setValue(numReq); 
    refSheet.getRange(rowIdx, 12).setValue(autorizacion); 
    refSheet.getRange(rowIdx, 13).setValue(ordenCompra); 
    refSheet.getRange(rowIdx, 14).setValue(recepcionMaterial);
    refSheet.getRange(rowIdx, 22).setValue(dirAprobacion);
    
    if (urlFotoIncompleto) refSheet.getRange(rowIdx, 15).setValue(urlFotoIncompleto); 
    
    // REGISTRO DE TIMESTAMPS SOLO SI AÚN NO EXISTEN (EVITA SOBREESCRIBIR FECHA ORIGINAL)
    if (cotizacion === "In Process" || cotizacion === "Done" || urlPDFCotArray.length > 0) {
      if (urlPDFCotArray.length > 0) {
        let oldVal = String(refSheet.getRange(rowIdx, 17).getValue() || "").trim();
        if (oldVal === "undefined" || oldVal === "N/A") oldVal = "";
        let newVal = oldVal ? (oldVal + "\n" + urlPDFCotArray.join("\n")) : urlPDFCotArray.join("\n");
        refSheet.getRange(rowIdx, 17).setValue(newVal);
      }
      let valTsCot = String(refSheet.getRange(rowIdx, 24).getValue() || "").trim();
      if (!valTsCot || valTsCot === "undefined" || valTsCot === "N/A") {
        refSheet.getRange(rowIdx, 24).setValue(tsNowStr); 
      }
    }
    
    if (urlPDFReq || (numReq && numReq.trim() !== "")) {
      if (urlPDFReq) refSheet.getRange(rowIdx, 20).setValue(urlPDFReq);
      let valTsReq = String(refSheet.getRange(rowIdx, 25).getValue() || "").trim();
      if (!valTsReq || valTsReq === "undefined" || valTsReq === "N/A") {
        refSheet.getRange(rowIdx, 25).setValue(tsNowStr); 
      }
    }

    if (urlPDFAut || autorizacion === "Autorizado") {
      if (urlPDFAut) refSheet.getRange(rowIdx, 18).setValue(urlPDFAut);
      let valTsAut = String(refSheet.getRange(rowIdx, 26).getValue() || "").trim();
      if (!valTsAut || valTsAut === "undefined" || valTsAut === "N/A") {
        refSheet.getRange(rowIdx, 26).setValue(tsNowStr); 
      }
    }

    if (urlPDFDirAut || dirAprobacion === "Aprobado") {
      if (urlPDFDirAut) refSheet.getRange(rowIdx, 23).setValue(urlPDFDirAut);
      let valTsDir = String(refSheet.getRange(rowIdx, 27).getValue() || "").trim();
      if (!valTsDir || valTsDir === "undefined" || valTsDir === "N/A") {
        refSheet.getRange(rowIdx, 27).setValue(tsNowStr); 
      }
    }

    if (urlPDFOC || ordenCompra === "Recibido") {
      if (urlPDFOC) refSheet.getRange(rowIdx, 21).setValue(urlPDFOC);
      let valTsOC = String(refSheet.getRange(rowIdx, 28).getValue() || "").trim();
      if (!valTsOC || valTsOC === "undefined" || valTsOC === "N/A") {
        refSheet.getRange(rowIdx, 28).setValue(tsNowStr); 
      }
    }

    if(finalizadoHoy) { 
      let currentEntrega = String(refSheet.getRange(rowIdx, 19).getValue() || ""); 
      if(!currentEntrega || currentEntrega === "undefined" || currentEntrega === "N/A") { 
        refSheet.getRange(rowIdx, 19).setValue(tsNowStr); 
        refSheet.getRange(rowIdx, 29).setValue(tsNowStr); 
      } 
    }

    let r = refSheet.getRange(rowIdx, 1, 1, 23).getValues()[0]; 
    enviarCorreoCorporativo(`${tituloMail}: ${r[0]} - ${r[4]}`, tituloMail, `<p>${msjMail}</p>`, r[0], r[2], r[3], r[4], r[7], r[1], urlPDFCotArray[0] || r[16], urlPDFAut || r[17]);
    
    return { success: true, message: `¡Tarjeta actualizada a: ${estatusGeneral}!` };
  } catch(e) { return { success: false, message: "Error al actualizar: " + e.toString() }; }
}

function guardarRegistroManualOEE(datos) {
  try {
    let oeeSheet = getOrCreateOEESheet(); let plannedHrs = parseFloat(datos.plannedHours) || 168; let targetUnits = parseFloat(datos.targetUnits) || 10000; let realUnits = parseFloat(datos.realUnits) || 9500; let goodUnits = parseFloat(datos.goodUnits) || 9300;
    const ss = getDB(); let logSheet = ss.getSheetByName(LOG_SHEET_NAME); let totalDowntimeMins = 0;
    if (logSheet && logSheet.getLastRow() > 1) { let data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues(); data.forEach(r => { let tsA = parseInt(r[12]); let tsC = parseInt(r[13]); if (tsA && tsC && tsC >= tsA) totalDowntimeMins += (tsC - tsA) / 60000; }); }
    let downtimeHrs = totalDowntimeMins / 60; let avail = Math.max(0, (plannedHrs - downtimeHrs) / plannedHrs) * 100; let perf = targetUnits > 0 ? (realUnits / targetUnits) * 100 : 0; let qual = realUnits > 0 ? (goodUnits / realUnits) * 100 : 0; let oeeVal = (avail * perf * qual) / 10000;
    oeeSheet.appendRow([new Date().toLocaleString('es-MX', {hour12:false}), plannedHrs, targetUnits, realUnits, goodUnits, avail.toFixed(2), perf.toFixed(2), qual.toFixed(2), oeeVal.toFixed(2)]);
    return { success: true, message: "¡Parámetros OEE guardados!" };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function resetearSistemaACero(pass) {
  if (pass !== ADMIN_PASS) return { success: false, message: "Contraseña incorrecta." };
  try { const ss = getDB(); let sheetsToClear = [OPEN_SHEET_NAME, LOG_SHEET_NAME, EHS_SHEET_NAME, REF_SHEET_NAME, OEE_SHEET_NAME]; sheetsToClear.forEach(name => { let sh = ss.getSheetByName(name); if(sh && sh.getLastRow() > 1) { sh.deleteRows(2, sh.getLastRow() - 1); } }); return { success: true, message: "¡Sistema reseteado a 0!" }; } catch(e) { return { success: false, message: e.toString() }; }
}
// =========================================================================
// MÓDULO PERMISOS TRABAJO SEGURO EHS - GUARDAR SOLICITUD Y CREAR FOLIO
// =========================================================================
function solicitarPermisoTrabajoEHS(datos) { // 👈 NUEVA FUNCIÓN COMPLETA
  try {
    var ss = getDB(); // 👈 Adaptado a tu base de datos getDB()
    var sheet = ss.getSheetByName("Log_Permisos_EHS");
    
    // Crear pestaña en Google Sheets si no existe
    if (!sheet) {
      sheet = ss.insertSheet("Log_Permisos_EHS");
      sheet.appendRow([
        "Folio Permiso", "Fecha Registro", "Actividad Alto Riesgo", 
        "Técnico Ejecutor", "Supervisor", "Ubicación", "Tiempo Estimado", 
        "Estatus", "URL Evidencias Previas", "Inspector Aprobación EHS", 
        "Fecha Aprobación EHS", "URL Evidencia Aprobación EHS",
        "Fecha Cierre Supervisor", "URL Evidencia Cierre", 
        "Inspector Cierre EHS", "Fecha Cierre EHS", "URL Evidencia Cierre EHS"
      ]);
      sheet.getRange(1, 1, 1, 17).setFontWeight("bold").setBackground("#009B4A").setFontColor("#ffffff");
    }

    var lastRow = sheet.getLastRow();
    var nextNum = lastRow > 1 ? lastRow : 1;
    var numFormatted = ("0000" + nextNum).slice(-4);
    var folio = "PERM-" + numFormatted; // 👈 Genera folios tipo PERM-0001
    
    var timeZone = Session.getScriptTimeZone() || "GMT-6";
    var fechaStr = Utilities.formatDate(new Date(), timeZone, "dd/MM/yyyy, HH:mm:ss");

    // Guardar evidencias en Google Drive
    var folderName = "WorkPulse_Evidencias_Permisos_EHS";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    var urlsArr = [];
    if (datos.fotos && datos.fotos.length > 0) {
      for (var i = 0; i < datos.fotos.length; i++) {
        var fotoObj = datos.fotos[i];
        if (fotoObj && fotoObj.base64) {
          var rawBase64 = String(fotoObj.base64)
            .replace(/^data:image\/\w+;base64,/, "")
            .replace(/^data:application\/pdf;base64,/, "")
            .replace(/[\r\n\s]/g, "");
            
          var bytes = Utilities.base64Decode(rawBase64);
          var blob = Utilities.newBlob(
            bytes, 
            fotoObj.tipo || "image/jpeg", 
            folio + "_EvidenciaPrevia_" + (i + 1) + ".jpg"
          );
          
          var file = folder.createFile(blob);
          
          try {
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch (eSharing) {}
          
          urlsArr.push(file.getUrl());
        }
      }
    }

    // Registrar fila en Log_Permisos_EHS
    sheet.appendRow([
      folio,
      fechaStr,
      datos.actividad,
      datos.tecnico,
      datos.supervisor,
      datos.ubicacion,
      datos.tiempoEstimado,
      "SOLICITADO", // 👈 Estatus inicial para que pase a Validación EHS
      urlsArr.join("\n"),
      "", "", "", "", "", "", "", ""
    ]);

    return {
      success: true,
      message: "🛡️ PERMISO EHS SOLICITADO CON ÉXITO: " + folio
    };

  } catch (e) {
    return {
      success: false,
      message: "Error al registrar permiso EHS: " + e.message
    };
  }
}
// =========================================================================
// MÓDULO TARJETAS MANTENIMIENTO - CÁLCULO Y CARGA DE 5W2H (👈 NUEVO)
// =========================================================================
function getTarjetasMttoModuloCompleto() { // 👈 NUEVA FUNCIÓN COMPLETA
  try {
    const ss = getDB();
    let sheet = ss.getSheetByName(OPEN_SHEET_NAME);
    let cards = [];
    let counts = { total: 0, conRef: 0, sinRef: 0, criticos: 0, turnoD: 0, turnoE: 0, turnoF: 0 };
    let nowTs = Date.now();

    if (sheet && sheet.getLastRow() > 1) {
      let maxCols = sheet.getLastColumn();
      let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, maxCols).getValues();

      for (let i = 0; i < data.length; i++) {
        let r = data[i];
        let folio = String(r[0] || "").trim();
        if (!folio || folio === "N/A") continue;

        let fechaStr = String(r[1] || "N/A");
        let que = String(r[2] || "N/A");
        let porque = String(r[3] || "N/A");
        let dondeStr = String(r[4] || "");
        let cuando = String(r[5] || "N/A");
        let quien = String(r[6] || "N/A");
        let como = String(r[7] || "N/A");
        let refaccionesInfo = String(r[8] || "NO");
        let fotoUrl = String(r[9] || "");
        let tsApertura = parseInt(r[10]) || nowTs;

        let diffMs = Math.max(0, nowTs - tsApertura);
        let esCritico = (diffMs >= 86400000); // Más de 24 horas abierto

        counts.total++;
        if (esCritico) counts.criticos++;

        let tieneRef = refaccionesInfo.toUpperCase().indexOf("SÍ") !== -1 || refaccionesInfo.toUpperCase().indexOf("SI") !== -1;
        if (tieneRef) counts.conRef++; else counts.sinRef++;

        let pDonde = dondeStr.indexOf("||") !== -1 ? dondeStr.split("||").map(s => s.trim()) : [];
        let area = pDonde[0] || "N/A";
        let linea = pDonde[1] || "N/A";
        let maquina = pDonde[2] || "MÁQUINA GENERAL";
        let idMaquina = pDonde[3] || "N/A";
        let turno = pDonde[4] || "N/A";

        if (turno.indexOf("Turno D") !== -1) counts.turnoD++;
        else if (turno.indexOf("Turno E") !== -1) counts.turnoE++;
        else if (turno.indexOf("Turno F") !== -1) counts.turnoF++;

        cards.push({
          rowIdx: i + 2,
          folio: folio,
          fecha: fechaStr,
          que: que,
          porque: porque,
          dondeStr: dondeStr,
          area: area,
          linea: linea,
          maquina: maquina,
          idMaquina: idMaquina,
          turno: turno,
          cuando: cuando,
          quien: quien,
          como: como,
          refaccionesInfo: refaccionesInfo,
          tieneRef: tieneRef,
          fotoUrl: fotoUrl,
          tsApertura: tsApertura,
          esCritico: esCritico
        });
      }
    }
    return { success: true, counts: counts, cards: cards };
  } catch(e) {
    return { success: false, errorServer: "Error cargando tarjetas de mantenimiento: " + e.toString() };
  }
}
// =========================================================================
// MÓDULO TODAS LAS MÁQUINAS - CÁLCULO PARETO 80/20 (👈 TODAS INCLUIDAS)
// =========================================================================
function getTarjetasMttoModuloCompleto() {
  try {
    const ss = getDB();
    let sheet = ss.getSheetByName(OPEN_SHEET_NAME);
    let cards = [];
    let counts = { total: 0, conRef: 0, sinRef: 0, criticos: 0, turnoD: 0, turnoE: 0, turnoF: 0 };
    let nowTs = Date.now();

    if (sheet && sheet.getLastRow() > 1) {
      let maxCols = Math.max(17, sheet.getLastColumn());
      let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, maxCols).getValues();

      for (let i = 0; i < data.length; i++) {
        let r = data[i];
        let folio = String(r[0] || "").trim();
        if (!folio || folio === "N/A") continue;

        let fechaStr = String(r[1] || "N/A");
        let que = String(r[2] || "N/A");
        let porque = String(r[3] || "N/A");
        let dondeStr = String(r[4] || "");
        let cuando = String(r[5] || "N/A");
        let quien = String(r[6] || "N/A");
        let como = String(r[7] || "N/A");
        let refaccionesInfo = String(r[8] || "NO");
        let fotoUrl = String(r[9] || "");
        let tsApertura = parseInt(r[10]) || nowTs;

        // 🟢 CAMPOS DEL PLAN DE MANTENIMIENTO (COLUMNAS L A Q)
        let tiempoSolicitado = String(r[11] || "").trim();
        let tecNoEmp = String(r[12] || "").trim();
        let tecnicoAsignado = String(r[13] || "").trim();
        let supNoEmp = String(r[14] || "").trim();
        let supervisorAsignado = String(r[15] || "").trim();
        let actividadProgramada = String(r[16] || "").trim();
        let planGuardado = (tiempoSolicitado !== "" && tecnicoAsignado !== "" && supervisorAsignado !== "");

        let diffMs = Math.max(0, nowTs - tsApertura);
        let esCritico = (diffMs >= 86400000);

        counts.total++;
        if (esCritico) counts.criticos++;

        let tieneRef = refaccionesInfo.toUpperCase().indexOf("SÍ") !== -1 || refaccionesInfo.toUpperCase().indexOf("SI") !== -1;
        if (tieneRef) counts.conRef++; else counts.sinRef++;

        let pDonde = dondeStr.indexOf("||") !== -1 ? dondeStr.split("||").map(s => s.trim()) : [];
        let area = pDonde[0] || "N/A";
        let linea = pDonde[1] || "N/A";
        let maquina = pDonde[2] || "MÁQUINA GENERAL";
        let idMaquina = pDonde[3] || "N/A";
        let turno = pDonde[4] || "N/A";

        if (turno.indexOf("Turno D") !== -1) counts.turnoD++;
        else if (turno.indexOf("Turno E") !== -1) counts.turnoE++;
        else if (turno.indexOf("Turno F") !== -1) counts.turnoF++;

        cards.push({
          rowIdx: i + 2,
          folio: folio,
          fecha: fechaStr,
          que: que,
          porque: porque,
          dondeStr: dondeStr,
          area: area,
          linea: linea,
          maquina: maquina,
          idMaquina: idMaquina,
          turno: turno,
          cuando: cuando,
          quien: quien,
          como: como,
          refaccionesInfo: refaccionesInfo,
          tieneRef: tieneRef,
          fotoUrl: fotoUrl,
          tsApertura: tsApertura,
          esCritico: esCritico,
          tiempoSolicitado: tiempoSolicitado,
          tecNoEmp: tecNoEmp,
          tecnicoAsignado: tecnicoAsignado,
          supNoEmp: supNoEmp,
          supervisorAsignado: supervisorAsignado,
          actividadProgramada: actividadProgramada,
          planGuardado: planGuardado
        });
      }
    }
    return { success: true, counts: counts, cards: cards };
  } catch(e) {
    return { success: false, errorServer: "Error cargando tarjetas de mantenimiento: " + e.toString() };
  }
}
// =========================================================================
// MÓDULO PERSONAL - GESTIÓN DE PLANTILLA DE MANTENIMIENTO
// =========================================================================
function getOrCreatePersonalSheet() {
  const PERSONAL_SHEET_NAME = "Log_Personal";
  const ss = getDB();
  let sheet = ss.getSheetByName(PERSONAL_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PERSONAL_SHEET_NAME);
    sheet.appendRow(["NO. EMPLEADO", "NOMBRE", "TURNO"]);
    sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#2563eb").setFontColor("white");

    // 📋 PRECARGA AUTOMÁTICA DE TUS 19 EMPLEADOS DE PERSONAL.XLSX
    const personalInicial = [
      [1077, "MEZA MONRREAL, RAMON EDUARDO", "Turno D"],
      [1451, "RAMIREZ HERNANDEZ, JOSE MARTIN", "Turno D"],
      [256, "GOMEZ LEYVA, JESUS MIGUEL", "Turno F"],
      [19524, "RAMIREZ ARROYO, NICOLAS", "Turno D"],
      [318, "BAUTISTA MONTES, ANTONIO", "Turno F"],
      [329, "DORIA RODRIGUEZ, SIMON", "Turno F"],
      [544, "MARTINEZ AVILA, JOSE FRANCISCO", "Turno E"],
      [948, "FLORES SANTIAGO, JUAN", "Turno D"],
      [13686, "ESPINOZA OLIVARES, ADRIAN", "Turno F"],
      [27447, "LOPEZ GUTIERREZ, MARCO URIEL", "Turno E"],
      [199, "VEGA CANALES, JOSE ANGEL", "Turno D"],
      [227, "CAVAZOS NUÑEZ, FRANCISCO ISAAC", "Turno E"],
      [267, "EXELENTE HERNANDEZ, IVAN LOPSA", "Turno D"],
      [325, "BLANCO HERNANDEZ, MIGUEL PATRI", "Turno D"],
      [588, "MONTERO BELLO, MANUEL", "Turno D"],
      [717, "BAUTISTA MORALES, ROGELIO", "Turno D"],
      [1105, "BAUTISTA HERNANDEZ, ALEJANDRO", "Turno D"],
      [1454, "CRUZ JUAREZ, MANUEL ENOC", "Turno F"],
      [30514, "CORTAZAR RODRIGUEZ, JUAN DANIEL", "Turno D"]
    ];

    personalInicial.forEach(p => sheet.appendRow(p));
  }
  return sheet;
}

/// 🔐 CONSULTAR PLANTILLA DE PERSONAL
function getPersonalData(pass) {
  if (pass !== ADMIN_PASS) {
    return { success: false, message: "Contraseña de Administrador incorrecta." };
  }
  try {
    let sheet = getOrCreatePersonalSheet();
    let data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues() : [];
    let list = data.map((r, idx) => ({
      rowIdx: idx + 2,
      noEmpleado: String(r[0]).trim(),
      nombre: String(r[1]).trim(),
      turno: String(r[2]).trim()
    }));
    return { success: true, personal: list };
  } catch(e) {
    return { success: false, message: "Error al consultar plantilla de personal: " + e.toString() };
  }
}

// 💾 GUARDAR / EDITAR EMPLEADO
function guardarPersonal(datos, pass) {
  if (pass !== ADMIN_PASS) {
    return { success: false, message: "Contraseña de Administrador incorrecta." };
  }
  try {
    let sheet = getOrCreatePersonalSheet();
    let noEmp = String(datos.noEmpleado).trim();
    let nombre = String(datos.nombre).trim().toUpperCase();
    let turno = String(datos.turno).trim();

    if (!noEmp || !nombre || !turno) {
      return { success: false, message: "Completa todos los campos obligatorios." };
    }

    let data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues() : [];
    let existingRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === noEmp) {
        existingRow = i + 2;
        break;
      }
    }

    if (existingRow > 1) {
      sheet.getRange(existingRow, 2).setValue(nombre);
      sheet.getRange(existingRow, 3).setValue(turno);
      return { success: true, message: `Empleado N° ${noEmp} actualizado con éxito.` };
    } else {
      sheet.appendRow([parseInt(noEmp) || noEmp, nombre, turno]);
      return { success: true, message: `Empleado ${nombre} agregado con éxito.` };
    }
  } catch(e) {
    return { success: false, message: "Error al guardar empleado: " + e.toString() };
  }
}

// 🗑️ ELIMINAR EMPLEADO
function eliminarPersonal(noEmpleado, pass) {
  if (pass !== ADMIN_PASS) {
    return { success: false, message: "Contraseña de Administrador incorrecta." };
  }
  try {
    let sheet = getOrCreatePersonalSheet();
    let data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues() : [];
    let targetRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(noEmpleado).trim()) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow > 1) {
      sheet.deleteRow(targetRow);
      return { success: true, message: `Empleado N° ${noEmpleado} eliminado correctamente.` };
    } else {
      return { success: false, message: "No se encontró el número de empleado en la base de datos." };
    }
  } catch(e) {
    return { success: false, message: "Error al eliminar empleado: " + e.toString() };
  }
}
// =========================================================================
// MÓDULO PERFORMANCE DE EMPLEADOS - DOBLE RANKING (CIERRES Y APERTURAS)
// =========================================================================
function getPerformanceEmpleadoData() {
  try {
    const ss = getDB();
    let openSheet = ss.getSheetByName(OPEN_SHEET_NAME);
    let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    let personalSheet = ss.getSheetByName("Log_Personal");

    let abiertosData = (openSheet && openSheet.getLastRow() > 1) 
      ? openSheet.getRange(2, 1, openSheet.getLastRow() - 1, openSheet.getLastColumn()).getValues() 
      : [];

    let cerradosData = (logSheet && logSheet.getLastRow() > 1) 
      ? logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues() 
      : [];

    let personalData = (personalSheet && personalSheet.getLastRow() > 1)
      ? personalSheet.getRange(2, 1, personalSheet.getLastRow() - 1, 3).getValues()
      : [];

    // 1. Conteo de CIERRES (Técnico - Col E / idx 4)
    let tecCounts = {};
    cerradosData.forEach(row => {
      let tecRaw = String(row[4] || "").trim().toUpperCase(); 
      if (tecRaw && tecRaw !== "N/A" && tecRaw !== "UNDEFINED") {
        tecCounts[tecRaw] = (tecCounts[tecRaw] || 0) + 1;
      }
    });

    // 2. Conteo de APERTURAS (Quién en Abiertos + Quién Abrió en Cerrados)
    let abridorCounts = {};

    abiertosData.forEach(row => {
      let abridorRaw = String(row[6] || "").trim().toUpperCase(); // Col G (Quién)
      if (abridorRaw && abridorRaw !== "N/A" && abridorRaw !== "UNDEFINED") {
        abridorCounts[abridorRaw] = (abridorCounts[abridorRaw] || 0) + 1;
      }
    });

    cerradosData.forEach(row => {
      let abridorRaw = String(row[10] || "").trim().toUpperCase(); // Col K (Quién Abrió)
      if (abridorRaw && abridorRaw !== "N/A" && abridorRaw !== "UNDEFINED") {
        abridorCounts[abridorRaw] = (abridorCounts[abridorRaw] || 0) + 1;
      }
    });

    // Función auxiliar para vincular con 'Log_Personal'
    let mapearConPersonal = (countsMap) => {
      let list = Object.keys(countsMap).map(person => {
        let matchPersonal = personalData.find(p => {
          let noEmp = String(p[0]).trim();
          let nombre = String(p[1]).trim().toUpperCase();
          return person.includes(noEmp) || person.includes(nombre) || nombre.includes(person);
        });

        return {
          persona: person,
          total: countsMap[person],
          noEmpleado: matchPersonal ? matchPersonal[0] : "N/A",
          nombreCompleto: matchPersonal ? matchPersonal[1] : person,
          turno: matchPersonal ? matchPersonal[2] : "Turno N/A"
        };
      });
      list.sort((a, b) => b.total - a.total);
      return list;
    };

    let rankingCierres = mapearConPersonal(tecCounts);
    let rankingAperturas = mapearConPersonal(abridorCounts);

    return {
      success: true,
      totalCerrados: cerradosData.length,
      totalAperturas: abiertosData.length + cerradosData.length,
      rankingCierres: rankingCierres,
      rankingAperturas: rankingAperturas
    };
  } catch(e) {
    return { success: false, errorServer: "Error al calcular performance: " + e.toString() };
  }
}
// =========================================================================
// MÓDULO EQUIPOS - GESTIÓN DE CATÁLOGO TÉCNICO DE MAQUINARIA
// =========================================================================
function getOrCreateEquiposSheet() {
  const EQUIPOS_SHEET_NAME = "Log_Equipos";
  const ss = getDB();
  let sheet = ss.getSheetByName(EQUIPOS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(EQUIPOS_SHEET_NAME);
    sheet.appendRow(["ITEM", "LINEA", "ID MAQUINA", "EQUIPO"]);
    sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#0284c7").setFontColor("white");

    // 📋 PRECARGA AUTOMÁTICA DE LOS 438 EQUIPOS DE EQUIPOS.XLSX
    const equiposInicial = [
      [1, "1", "MRC-0110", "PRENSA DE POLEAS"],
      [2, "1", "MRE-0105", "EMBOBINADORA"],
      [3, "1", "MRE-0108", "Transferencia"],
      [4, "1", "MRE-0110", "TRANSFERENCIA"],
      [5, "1", "MRE-0113", "PRENSA G AUXILIAR"],
      [6, "1", "MRE-0114", "BLOQUER"],
      [7, "1", "MRE-0118", "EMBOBINADORA"],
      [8, "1", "MRE-0119", "CONVEYOR"],
      [9, "1", "MRE-0125", "INSULADORA 4 ESTACIONES"],
      [10, "1", "MRE-0127", "AMARRADORA CHINA #1 BXSIIL4 - 160/160"],
      [11, "1", "MRE-0128", "AMARRADORA CHINA #2 BXSIIL4 - 160/160"],
      [12, "1", "MRE-0129", "ROBOT YAZKAWA"],
      [13, "1", "MRE-0201", "CLEATER"],
      [14, "1", "MRE-0214", "PRENSA G PRINCIPAL"],
      [15, "1", "MRE-0410", "EMBOBINADORA"],
      [16, "1", "MRE-0414", "INSULADORA DE ESTATOR"],
      [17, "1", "MRE-0415", "FORMADORA FINAL"],
      [18, "1", "MRE-0417", "INSERTADORA DE BOBINAS"],
      [19, "1", "MRE-0419", "FORMADORA INTERMEDIA"],
      [20, "1", "MRE-0422", "PROBADOR DE ESTATOR BAKER"],
      [21, "2", "MRE-0102", "PRENSA DE POLEAS"],
      [22, "2", "MRE-0107", "EMBOBINADORA DE BOBINA PRINCIPAL"],
      [23, "2", "MRE-0111", "INSERTADORA DE BOBINAS"],
      [24, "2", "MRE-0112", "FORMADORA INTERMEDIA"],
      [25, "2", "MRE-0116", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [26, "2", "MRE-0120", "AMARRADORA"],
      [27, "2", "MRE-0121", "FORMADORA FINAL"],
      [28, "2", "MRE-0122", "AMARRADORA"],
      [29, "2", "MRE-0124", "PROBADOR DE ESTATOR BAKER"],
      [30, "2", "MRE-0126", "CLEATER INSERCION Y CORTE DE CLEAT"],
      [31, "2", "MRE-0203", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [32, "2", "MRE-0204", "CONVEYOR INTERMEDIO DE TRANSFERENCIA DE ESTATOR"],
      [33, "2", "MRE-0210", "TRANSFERENCIA BANCO DE TRABAJO Y GIRO DE ESTATOR"],
      [34, "2", "MRE-0412", "EMBOBINADORA DE BOBINAS DE AUXILIAR Y PRINCIPAL"],
      [35, "2", "MRE-0413", "PRENSA G AUXILIAR APLICADORA DE TERMINAL"],
      [36, "3", "MRE-0103", "CLEATER DE 4 ESTACIONES"],
      [37, "3", "MRE-0106", "INSULADORA DE DRAFT"],
      [38, "3", "MRE-0109", "PRENSA G APLICADORA DE TERMINALES AUXILIAR"],
      [39, "3", "MRE-0115", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [40, "3", "MRE-0117", "AMARRADORA"],
      [41, "3", "MRE-0123", "FORMADORA FINAL"],
      [42, "3", "MRE-0202", "PRENSA DE POLEAS"],
      [43, "3", "MRE-0205", "PRENSA G APLICADORA DE TERMINALES PRINCIPAL"],
      [44, "3", "MRE-0207", "INSERTADORA DE BOBINAS"],
      [45, "3", "MRE-0208", "FORMADORA INTERMEDIA"],
      [46, "3", "MRE-0209", "AMARRADORA DE ESTATOR"],
      [47, "3", "MRE-0212", "PROBADOR BAKER"],
      [48, "3", "MRE-0213", "EMBOBINADORA DE BOBINA PRINCIPAL"],
      [49, "3", "MRE-0411", "EMBOBINADORA DE BOBINA AUXILIAR"],
      [50, "3", "MRE-0420", "PROBADOR DE BAKER MANUAL"],
      [51, "4", "MRE-0301", "PRENSA DE POLEAS"],
      [52, "4", "MRE-0302", "EMBOBINADORA DE BOBINA AUXILIAR"],
      [53, "4", "MRE-0303", "PRENSA G APLICADORA DE TERMINALES AUXILIAR"],
      [54, "4", "MRE-0304", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [55, "4", "MRE-0305", "EMBOBINADORA DE BOBINA PRINCIPAL"],
      [56, "4", "MRE-0306", "PRENSA G APLICADORA DE TERMINALES PRINCIPAL"],
      [57, "4", "MRE-0307", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [58, "4", "MRE-0308", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [59, "4", "MRE-0309", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [60, "4", "MRE-0310", "FORMADORA INTERMEDIA DE ESTATOR"],
      [61, "4", "MRE-0311", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [62, "4", "MRE-0312", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [63, "4", "MRE-0313", "FORMADORA FINAL DE ESTATOR"],
      [64, "4", "MRE-0314", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [65, "4", "MRE-0315", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [66, "5", "MRE-0501", "PRENSA DE POLEAS"],
      [67, "5", "MRE-0502", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [68, "5", "MRE-0503", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [69, "5", "MRE-0504", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [70, "5", "MRE-0505", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [71, "5", "MRE-0506", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [72, "5", "MRE-0507", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [73, "5", "MRE-0508", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [74, "5", "MRE-0509", "FORMADORA INTERMEDIA DE ESTATOR"],
      [75, "5", "MRE-0510", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [76, "5", "MRE-0511", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [77, "5", "MRE-0512", "FORMADORA FINAL DE ESTATOR"],
      [78, "5", "MRE-0513", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [79, "5", "MRE-0514", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [80, "6", "MRE-0601", "EMBOBINADORA DE BOBINA AUXILIAR"],
      [81, "6", "MRE-0602", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [82, "6", "MRE-0603", "PRENSA DE POLEAS"],
      [83, "6", "MRE-0604", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [84, "6", "MRE-0605", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [85, "6", "MRE-0606", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [86, "6", "MRE-0607", "EMBOBINADORA DE BOBINA PRINCIPAL"],
      [87, "6", "MRE-0608", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [88, "6", "MRE-0609", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [89, "6", "MRE-0610", "FORMADORA INTERMEDIA DE ESTATOR"],
      [90, "6", "MRE-0611", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [91, "6", "MRE-0612", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [92, "6", "MRE-0613", "FORMADORA FINAL DE ESTATOR"],
      [93, "6", "MRE-0614", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [94, "6", "MRE-0615", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [95, "7", "MRE-0701", "PRENSA DE POLEAS"],
      [96, "7", "MRE-0702", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [97, "7", "MRE-0703", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [98, "7", "MRE-0704", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [99, "7", "MRE-0705", "EMBOBINADORA DE BOBINA AUXILIAR"],
      [100, "7", "MRE-0706", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [101, "7", "MRE-0707", "EMBOBINADORA DE BOBINA PRINCIPAL"],
      [102, "7", "MRE-0708", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [103, "7", "MRE-0709", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [104, "7", "MRE-0710", "FORMADORA INTERMEDIA DE ESTATOR"],
      [105, "7", "MRE-0711", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [106, "7", "MRE-0712", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [107, "7", "MRE-0713", "FORMADORA FINAL DE ESTATOR"],
      [108, "7", "MRE-0714", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [109, "7", "MRE-0715", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [110, "8", "MRE-0801", "PRENSA DE POLEAS"],
      [111, "8", "MRE-0802", "EMBOBINADORA DE BOBINA AUXILIAR"],
      [112, "8", "MRE-0803", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [113, "8", "MRE-0804", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [114, "8", "MRE-0805", "EMBOBINADORA DE BOBINA PRINCIPAL"],
      [115, "8", "MRE-0806", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [116, "8", "MRE-0807", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [117, "8", "MRE-0808", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [118, "8", "MRE-0809", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [119, "8", "MRE-0810", "FORMADORA INTERMEDIA DE ESTATOR"],
      [120, "8", "MRE-0811", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [121, "8", "MRE-0812", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [122, "8", "MRE-0813", "FORMADORA FINAL DE ESTATOR"],
      [123, "8", "MRE-0814", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [124, "8", "MRE-0815", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [125, "9", "MRE-0901", "PRENSA DE POLEAS"],
      [126, "9", "MRE-0902", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [127, "9", "MRE-0903", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [128, "9", "MRE-0904", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [129, "9", "MRE-0905", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [130, "9", "MRE-0906", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [131, "9", "MRE-0907", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [132, "9", "MRE-0908", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [133, "9", "MRE-0909", "FORMADORA INTERMEDIA DE ESTATOR"],
      [134, "9", "MRE-0910", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [135, "9", "MRE-0911", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [136, "9", "MRE-0912", "FORMADORA FINAL DE ESTATOR"],
      [137, "9", "MRE-0913", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [138, "9", "MRE-0914", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [139, "10", "MRE-1001", "PRENSA DE POLEAS"],
      [140, "10", "MRE-1002", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [141, "10", "MRE-1003", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [142, "10", "MRE-1004", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [143, "10", "MRE-1005", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [144, "10", "MRE-1006", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [145, "10", "MRE-1007", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [146, "10", "MRE-1008", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [147, "10", "MRE-1009", "FORMADORA INTERMEDIA DE ESTATOR"],
      [148, "10", "MRE-1010", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [149, "10", "MRE-1011", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [150, "10", "MRE-1012", "FORMADORA FINAL DE ESTATOR"],
      [151, "10", "MRE-1013", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [152, "10", "MRE-1014", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [153, "11", "MRE-1101", "PRENSA DE POLEAS"],
      [154, "11", "MRE-1102", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [155, "11", "MRE-1103", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [156, "11", "MRE-1104", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [157, "11", "MRE-1105", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [158, "11", "MRE-1106", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [159, "11", "MRE-1107", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [160, "11", "MRE-1108", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [161, "11", "MRE-1109", "FORMADORA INTERMEDIA DE ESTATOR"],
      [162, "11", "MRE-1110", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [163, "11", "MRE-1111", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [164, "11", "MRE-1112", "FORMADORA FINAL DE ESTATOR"],
      [165, "11", "MRE-1113", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [166, "11", "MRE-1114", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [167, "12", "MRE-1201", "PRENSA DE POLEAS"],
      [168, "12", "MRE-1202", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [169, "12", "MRE-1203", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [170, "12", "MRE-1204", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [171, "12", "MRE-1205", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [172, "12", "MRE-1206", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [173, "12", "MRE-1207", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [174, "12", "MRE-1208", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [175, "12", "MRE-1209", "FORMADORA INTERMEDIA DE ESTATOR"],
      [176, "12", "MRE-1210", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [177, "12", "MRE-1211", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [178, "12", "MRE-1212", "FORMADORA FINAL DE ESTATOR"],
      [179, "12", "MRE-1213", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [180, "12", "MRE-1214", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [181, "13", "MRE-1301", "PRENSA DE POLEAS"],
      [182, "13", "MRE-1302", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [183, "13", "MRE-1303", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [184, "13", "MRE-1304", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [185, "13", "MRE-1305", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [186, "13", "MRE-1306", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [187, "13", "MRE-1307", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [188, "13", "MRE-1308", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [189, "13", "MRE-1309", "FORMADORA INTERMEDIA DE ESTATOR"],
      [190, "13", "MRE-1310", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [191, "13", "MRE-1311", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [192, "13", "MRE-1312", "FORMADORA FINAL DE ESTATOR"],
      [193, "13", "MRE-1313", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [194, "13", "MRE-1314", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [195, "14", "MRE-1401", "PRENSA DE POLEAS"],
      [196, "14", "MRE-1402", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [197, "14", "MRE-1403", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [198, "14", "MRE-1404", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [199, "14", "MRE-1405", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [200, "14", "MRE-1406", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [201, "14", "MRE-1407", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [202, "14", "MRE-1408", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [203, "14", "MRE-1409", "FORMADORA INTERMEDIA DE ESTATOR"],
      [204, "14", "MRE-1410", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [205, "14", "MRE-1411", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [206, "14", "MRE-1412", "FORMADORA FINAL DE ESTATOR"],
      [207, "14", "MRE-1413", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [208, "14", "MRE-1414", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [209, "15", "MRE-1501", "PRENSA DE POLEAS"],
      [210, "15", "MRE-1502", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [211, "15", "MRE-1503", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [212, "15", "MRE-1504", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [213, "15", "MRE-1505", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [214, "15", "MRE-1506", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [215, "15", "MRE-1507", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [216, "15", "MRE-1508", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [217, "15", "MRE-1509", "FORMADORA INTERMEDIA DE ESTATOR"],
      [218, "15", "MRE-1510", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [219, "15", "MRE-1511", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [220, "15", "MRE-1512", "FORMADORA FINAL DE ESTATOR"],
      [221, "15", "MRE-1513", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [222, "15", "MRE-1514", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [223, "16", "MRE-1601", "PRENSA DE POLEAS"],
      [224, "16", "MRE-1602", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [225, "16", "MRE-1603", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [226, "16", "MRE-1604", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [227, "16", "MRE-1605", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [228, "16", "MRE-1606", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [229, "16", "MRE-1607", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [230, "16", "MRE-1608", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [231, "16", "MRE-1609", "FORMADORA INTERMEDIA DE ESTATOR"],
      [232, "16", "MRE-1610", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [233, "16", "MRE-1611", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [234, "16", "MRE-1612", "FORMADORA FINAL DE ESTATOR"],
      [235, "16", "MRE-1613", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [236, "16", "MRE-1614", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [237, "17", "MRE-1701", "PRENSA DE POLEAS"],
      [238, "17", "MRE-1702", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [239, "17", "MRE-1703", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [240, "17", "MRE-1704", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [241, "17", "MRE-1705", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [242, "17", "MRE-1706", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [243, "17", "MRE-1707", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [244, "17", "MRE-1708", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [245, "17", "MRE-1709", "FORMADORA INTERMEDIA DE ESTATOR"],
      [246, "17", "MRE-1710", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [247, "17", "MRE-1711", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [248, "17", "MRE-1712", "FORMADORA FINAL DE ESTATOR"],
      [249, "17", "MRE-1713", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [250, "17", "MRE-1714", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [251, "18", "MRE-1801", "PRENSA DE POLEAS"],
      [252, "18", "MRE-1802", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [253, "18", "MRE-1803", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [254, "18", "MRE-1804", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [255, "18", "MRE-1805", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [256, "18", "MRE-1806", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [257, "18", "MRE-1807", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [258, "18", "MRE-1808", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [259, "18", "MRE-1809", "FORMADORA INTERMEDIA DE ESTATOR"],
      [260, "18", "MRE-1810", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [261, "18", "MRE-1811", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [262, "18", "MRE-1812", "FORMADORA FINAL DE ESTATOR"],
      [263, "18", "MRE-1813", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [264, "18", "MRE-1814", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [265, "19", "MRE-1901", "PRENSA DE POLEAS"],
      [266, "19", "MRE-1902", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [267, "19", "MRE-1903", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [268, "19", "MRE-1904", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [269, "19", "MRE-1905", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [270, "19", "MRE-1906", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [271, "19", "MRE-1907", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [272, "19", "MRE-1908", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [273, "19", "MRE-1909", "FORMADORA INTERMEDIA DE ESTATOR"],
      [274, "19", "MRE-1910", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [275, "19", "MRE-1911", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [276, "19", "MRE-1912", "FORMADORA FINAL DE ESTATOR"],
      [277, "19", "MRE-1913", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [278, "19", "MRE-1914", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [279, "20", "MRE-2001", "PRENSA DE POLEAS"],
      [280, "20", "MRE-2002", "EMBOBINADORA DE BOBINAS DE PRINCIPAL Y AUXILIAR"],
      [281, "20", "MRE-2003", "PRENSA G APLICADORA DE TERMINAL Y CONEXION DE LEAD WIRE"],
      [282, "20", "MRE-2004", "PRENSA G APLICADORA DE TERMINAL AUXILIAR DE LEAD WIRE"],
      [283, "20", "MRE-2005", "CLEATER INSERCION Y CORTE DE CLEAT DE DRAFT"],
      [284, "20", "MRE-2006", "INSULADORA DE DRAFT DE 4 ESTACIONES"],
      [285, "20", "MRE-2007", "INSULADORA DE ESTATOR DE 4 ESTACIONES"],
      [286, "20", "MRE-2008", "INSERTADORA DE BOBINAS PRINCIPAL Y AUXILIAR"],
      [287, "20", "MRE-2009", "FORMADORA INTERMEDIA DE ESTATOR"],
      [288, "20", "MRE-2010", "AMARRADORA DE ESTATOR LADO LEAD WIRE"],
      [289, "20", "MRE-2011", "AMARRADORA DE ESTATOR LADO OPUESTO"],
      [290, "20", "MRE-2012", "FORMADORA FINAL DE ESTATOR"],
      [291, "20", "MRE-2013", "PROBADOR BAKER AUTOMATICO DE ESTATOR"],
      [292, "20", "MRE-2014", "PROBADOR BAKER MANUAL DE ESTATOR"],
      [293, "30", "MRE-0316", "ENSAMBLADORA Y PROBADORA DE MOTOR"],
      [294, "30", "MRE-0317", "PRENSA ARBOR"],
      [295, "30", "MRE-0318", "PRENSA SCHMIDT APLICADORA DE POLEA EN FLECHA DE ROTOR"],
      [296, "30", "MRE-0319", "PRENSA HYDROKRAFT APLICADORA DE BALERO EN FLECHA"],
      [297, "30", "MRE-0320", "PRENSA MESA ROTATIVA Y ALINEADORA DE SHIELD DE ALUMINIO"],
      [298, "30", "MRE-0321", "PROBADOR DE AIRE Y FUGAS EN TAPA LADO OPUESTO Y LADO FLECHA"],
      [299, "30", "MRE-0322", "PRENSA DAKE RECTIFICADORA DE SHIELD DE ALUMINIO"],
      [300, "30", "MRE-0323", "PRENSA PNEUPOWER APLICADORA DE BALERO EN FLECHA DE ROTOR"],
      [301, "30", "MRE-0324", "MESA INTERMEDIA DE BANCO DE TRABAJO Y ENSAMBLE DE MOTOR"],
      [302, "30", "MRE-0325", "BALANCEADORA AUTOMATICA DE ROTOR SCHENCK"],
      [303, "30", "MRE-0326", "ATORNILLADORA DENSITRON Y BANCO DE TRABAJO"],
      [304, "30", "MRE-0327", "PRENSA MESA ROTATIVA Y ALINEADORA DE SHIELD"],
      [305, "30", "MRE-0328", "PROBADOR BAKER MANUAL DE MOTOR ENSAMBLADO"],
      [306, "30", "MRE-0329", "CAJA DE EMPAQUE Y EMBALAJE DE MOTOR FINAL"],
      [307, "30", "MRE-0330", "CONVEYOR PRINCIPAL DE MOTOR ENSAMBLADO"],
      [308, "30", "MRE-0331", "ETIQUETADORA AUTOMATICA ZEBRA DE DATOS DE MOTOR"],
      [309, "31", "MRE-3101", "PRENSA DE POLEAS"],
      [310, "31", "MRE-3102", "ENSAMBLADORA Y PROBADORA DE MOTOR"],
      [311, "31", "MRE-3103", "PRENSA ARBOR"],
      [312, "31", "MRE-3104", "PRENSA SCHMIDT APLICADORA DE POLEA"],
      [313, "31", "MRE-3105", "PRENSA HYDROKRAFT APLICADORA DE BALERO"],
      [314, "31", "MRE-3106", "PRENSA MESA ROTATIVA Y ALINEADORA DE SHIELD"],
      [315, "31", "MRE-3107", "PROBADOR DE AIRE Y FUGAS"],
      [316, "31", "MRE-3108", "PRENSA DAKE RECTIFICADORA DE SHIELD"],
      [317, "31", "MRE-3109", "PRENSA PNEUPOWER APLICADORA DE BALERO"],
      [318, "31", "MRE-3110", "BALANCEADORA AUTOMATICA SCHENCK"],
      [319, "31", "MRE-3111", "ATORNILLADORA DENSITRON"],
      [320, "31", "MRE-3112", "PROBADOR BAKER MANUAL DE MOTOR"],
      [321, "31", "MRE-3113", "ETIQUETADORA AUTOMATICA ZEBRA"],
      [322, "32", "MRE-3201", "PRENSA DE POLEAS"],
      [323, "32", "MRE-3202", "ENSAMBLADORA Y PROBADORA DE MOTOR"],
      [324, "32", "MRE-3203", "PRENSA ARBOR"],
      [325, "32", "MRE-3204", "PRENSA SCHMIDT APLICADORA DE POLEA"],
      [326, "32", "MRE-3205", "PRENSA HYDROKRAFT APLICADORA DE BALERO"],
      [327, "32", "MRE-3206", "PRENSA MESA ROTATIVA Y ALINEADORA DE SHIELD"],
      [328, "32", "MRE-3207", "PROBADOR DE AIRE Y FUGAS"],
      [329, "32", "MRE-3208", "PRENSA DAKE RECTIFICADORA DE SHIELD"],
      [330, "32", "MRE-3209", "PRENSA PNEUPOWER APLICADORA DE BALERO"],
      [331, "32", "MRE-3210", "BALANCEADORA AUTOMATICA SCHENCK"],
      [332, "32", "MRE-3211", "ATORNILLADORA DENSITRON"],
      [333, "32", "MRE-3212", "PROBADOR BAKER MANUAL DE MOTOR"],
      [334, "32", "MRE-3213", "ETIQUETADORA AUTOMATICA ZEBRA"],
      [335, "33", "MRE-3301", "PRENSA DE POLEAS"],
      [336, "33", "MRE-3302", "ENSAMBLADORA Y PROBADORA DE MOTOR"],
      [337, "33", "MRE-3303", "PRENSA ARBOR"],
      [338, "33", "MRE-3304", "PRENSA SCHMIDT APLICADORA DE POLEA"],
      [339, "33", "MRE-3305", "PRENSA HYDROKRAFT APLICADORA DE BALERO"],
      [340, "33", "MRE-3306", "PRENSA MESA ROTATIVA Y ALINEADORA DE SHIELD"],
      [341, "33", "MRE-3307", "PROBADOR DE AIRE Y FUGAS"],
      [342, "33", "MRE-3308", "PRENSA DAKE RECTIFICADORA DE SHIELD"],
      [343, "33", "MRE-3309", "PRENSA PNEUPOWER APLICADORA DE BALERO"],
      [344, "33", "MRE-3310", "BALANCEADORA AUTOMATICA SCHENCK"],
      [345, "33", "MRE-3311", "ATORNILLADORA DENSITRON"],
      [346, "33", "MRE-3312", "PROBADOR BAKER MANUAL DE MOTOR"],
      [347, "33", "MRE-3313", "ETIQUETADORA AUTOMATICA ZEBRA"],
      [348, "34", "MRE-3401", "PRENSA DE POLEAS"],
      [349, "34", "MRE-3402", "ENSAMBLADORA Y PROBADORA DE MOTOR"],
      [350, "34", "MRE-3403", "PRENSA ARBOR"],
      [351, "34", "MRE-3404", "PRENSA SCHMIDT APLICADORA DE POLEA"],
      [352, "34", "MRE-3405", "PRENSA HYDROKRAFT APLICADORA DE BALERO"],
      [353, "34", "MRE-3406", "PRENSA MESA ROTATIVA Y ALINEADORA DE SHIELD"],
      [354, "34", "MRE-3407", "PROBADOR DE AIRE Y FUGAS"],
      [355, "34", "MRE-3408", "PRENSA DAKE RECTIFICADORA DE SHIELD"],
      [356, "34", "MRE-3409", "PRENSA PNEUPOWER APLICADORA DE BALERO"],
      [357, "34", "MRE-3410", "BALANCEADORA AUTOMATICA SCHENCK"],
      [358, "34", "MRE-3411", "ATORNILLADORA DENSITRON"],
      [359, "34", "MRE-3412", "PROBADOR BAKER MANUAL DE MOTOR"],
      [360, "34", "MRE-3413", "ETIQUETADORA AUTOMATICA ZEBRA"],
      [361, "35", "MRE-3501", "PRENSA DE POLEAS"],
      [362, "35", "MRE-3502", "ENSAMBLADORA Y PROBADORA DE MOTOR"],
      [363, "35", "MRE-3503", "PRENSA ARBOR"],
      [364, "35", "MRE-3504", "PRENSA SCHMIDT APLICADORA DE POLEA"],
      [365, "35", "MRE-3505", "PRENSA HYDROKRAFT APLICADORA DE BALERO"],
      [366, "35", "MRE-3506", "PRENSA MESA ROTATIVA Y ALINEADORA DE SHIELD"],
      [367, "35", "MRE-3507", "PROBADOR DE AIRE Y FUGAS"],
      [368, "35", "MRE-3508", "PRENSA DAKE RECTIFICADORA DE SHIELD"],
      [369, "35", "MRE-3509", "PRENSA PNEUPOWER APLICADORA DE BALERO"],
      [370, "35", "MRE-3510", "BALANCEADORA AUTOMATICA SCHENCK"],
      [371, "35", "MRE-3511", "ATORNILLADORA DENSITRON"],
      [372, "35", "MRE-3512", "PROBADOR BAKER MANUAL DE MOTOR"],
      [373, "35", "MRE-3513", "ETIQUETADORA AUTOMATICA ZEBRA"],
      [374, "36", "MRE-3601", "PRENSA DE POLEAS"],
      [375, "36", "MRE-3602", "ENSAMBLADORA Y PROBADORA DE MOTOR"],
      [376, "36", "MRE-3603", "PRENSA ARBOR"],
      [377, "36", "MRE-3604", "PRENSA SCHMIDT APLICADORA DE POLEA"],
      [378, "36", "MRE-3605", "PRENSA HYDROKRAFT APLICADORA DE BALERO"],
      [379, "36", "MRE-3606", "PRENSA MESA ROTATIVA Y ALINEADORA DE SHIELD"],
      [380, "36", "MRE-3607", "PROBADOR DE AIRE Y FUGAS"],
      [381, "36", "MRE-3608", "PRENSA DAKE RECTIFICADORA DE SHIELD"],
      [382, "36", "MRE-3609", "PRENSA PNEUPOWER APLICADORA DE BALERO"],
      [383, "36", "MRE-3610", "BALANCEADORA AUTOMATICA SCHENCK"],
      [384, "36", "MRE-3611", "ATORNILLADORA DENSITRON"],
      [385, "36", "MRE-3612", "PROBADOR BAKER MANUAL DE MOTOR"],
      [386, "36", "MRE-3613", "ETIQUETADORA AUTOMATICA ZEBRA"],
      [387, "37", "MRE-3701", "PRENSA DE POLEAS"],
      [388, "37", "MRE-3702", "ENSAMBLADORA Y PROBADORA DE MOTOR"],
      [389, "37", "MRE-3703", "PRENSA ARBOR"],
      [390, "37", "MRE-3704", "PRENSA SCHMIDT APLICADORA DE POLEA"],
      [391, "37", "MRE-3705", "PRENSA HYDROKRAFT APLICADORA DE BALERO"],
      [392, "37", "MRE-3706", "PRENSA MESA ROTATIVA Y ALINEADORA DE SHIELD"],
      [393, "37", "MRE-3707", "PROBADOR DE AIRE Y FUGAS"],
      [394, "37", "MRE-3708", "PRENSA DAKE RECTIFICADORA DE SHIELD"],
      [395, "37", "MRE-3709", "PRENSA PNEUPOWER APLICADORA DE BALERO"],
      [396, "37", "MRE-3710", "BALANCEADORA AUTOMATICA SCHENCK"],
      [397, "37", "MRE-3711", "ATORNILLADORA DENSITRON"],
      [398, "37", "MRE-3712", "PROBADOR BAKER MANUAL DE MOTOR"],
      [399, "37", "MRE-3713", "ETIQUETADORA AUTOMATICA ZEBRA"],
      [400, "38", "MRE-3801", "PRENSA DE POLEAS"],
      [401, "38", "MRE-3802", "ENSAMBLADORA Y PROBADORA DE MOTOR"],
      [402, "38", "MRE-3803", "PRENSA ARBOR"],
      [403, "38", "MRE-3804", "PRENSA SCHMIDT APLICADORA DE POLEA"],
      [404, "38", "MRE-3805", "PRENSA HYDROKRAFT APLICADORA DE BALERO"],
      [405, "38", "MRE-3806", "PRENSA MESA ROTATIVA Y ALINEADORA DE SHIELD"],
      [406, "38", "MRE-3807", "PROBADOR DE AIRE Y FUGAS"],
      [407, "38", "MRE-3808", "PRENSA DAKE RECTIFICADORA DE SHIELD"],
      [408, "38", "MRE-3809", "PRENSA PNEUPOWER APLICADORA DE BALERO"],
      [409, "38", "MRE-3810", "BALANCEADORA AUTOMATICA SCHENCK"],
      [410, "38", "MRE-3811", "ATORNILLADORA DENSITRON"],
      [411, "38", "MRE-3812", "PROBADOR BAKER MANUAL DE MOTOR"],
      [412, "38", "MRE-3813", "ETIQUETADORA AUTOMATICA ZEBRA"],
      [413, "39", "MRE-3901", "PRENSA DE POLEAS"],
      [414, "39", "MRE-3902", "ENSAMBLADORA Y PROBADORA DE MOTOR"],
      [415, "39", "MRE-3903", "PRENSA ARBOR"],
      [416, "39", "MRE-3904", "PRENSA SCHMIDT APLICADORA DE POLEA"],
      [417, "39", "MRE-3905", "PRENSA HYDROKRAFT APLICADORA DE BALERO"],
      [418, "39", "MRE-3906", "PRENSA MESA ROTATIVA Y ALINEADORA DE SHIELD"],
      [419, "39", "MRE-3907", "PROBADOR DE AIRE Y FUGAS"],
      [420, "39", "MRE-3908", "PRENSA DAKE RECTIFICADORA DE SHIELD"],
      [421, "39", "MRE-3909", "PRENSA PNEUPOWER APLICADORA DE BALERO"],
      [422, "39", "MRE-3910", "BALANCEADORA AUTOMATICA SCHENCK"],
      [423, "39", "MRE-3911", "ATORNILLADORA DENSITRON"],
      [424, "39", "MRE-3912", "PROBADOR BAKER MANUAL DE MOTOR"],
      [425, "39", "MRE-3913", "ETIQUETADORA AUTOMATICA ZEBRA"],
      [426, "40", "MRE-4001", "TROQUELADORA"],
      [427, "40", "MRE-4002", "PRENSA DAKE"],
      [428, "40", "MRE-4003", "DOBLADORA DE TUBO"],
      [429, "40", "MRE-4004", "CORORTADORA DE TUBO"],
      [430, "40", "N/A", "TERMINAL BIHLER"],
      [431, "40", "N/A", "TORNO"],
      [432, "40", "N/A", "FRESADORA"],
      [433, "40", "N/A", "TALADRO"],
      [434, "40", "N/A", "RECTIFICADORA"],
      [435, "40", "N/A", "SOLDADORA"],
      [436, "40", "N/A", "SOLDADORA MESA"],
      [437, "40", "N/A", "SIERRA VERTICAL"],
      [438, "40", "N/A", "SIERRA HORIZONTAL"]
    ];

    sheet.getRange(2, 1, equiposInicial.length, 4).setValues(equiposInicial);
  }
  return sheet;
}

// 🔐 OBTENER LISTA DE EQUIPOS (VALIDA CONTRASEÑA ADMIN)
function getEquiposData(pass) {
  if (pass !== ADMIN_PASS) {
    return { success: false, message: "Contraseña de Administrador incorrecta." };
  }
  try {
    let sheet = getOrCreateEquiposSheet();
    let data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues() : [];
    let list = data.map((r, idx) => ({
      rowIdx: idx + 2,
      item: parseInt(r[0]) || (idx + 1),
      linea: String(r[1]).trim(),
      idMaquina: String(r[2]).trim(),
      equipo: String(r[3]).trim()
    }));
    return { success: true, equipos: list };
  } catch(e) {
    return { success: false, message: "Error al consultar catálogo de equipos: " + e.toString() };
  }
}

// 💾 GUARDAR O EDITAR EQUIPO (VALIDA CONTRASEÑA ADMIN)
function guardarEquipo(datos, pass) {
  if (pass !== ADMIN_PASS) {
    return { success: false, message: "Contraseña de Administrador incorrecta." };
  }
  try {
    let sheet = getOrCreateEquiposSheet();
    let itemInput = datos.item ? parseInt(datos.item) : 0;
    let linea = String(datos.linea || "N/A").trim();
    let idMaquina = String(datos.idMaquina || "N/A").trim().toUpperCase();
    let equipo = String(datos.equipo || "").trim().toUpperCase();

    if (!equipo) {
      return { success: false, message: "Escribe el nombre del equipo." };
    }

    let data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues() : [];
    let existingRow = -1;

    if (itemInput > 0) {
      for (let i = 0; i < data.length; i++) {
        if (parseInt(data[i][0]) === itemInput) {
          existingRow = i + 2;
          break;
        }
      }
    }

    if (existingRow > 1) {
      sheet.getRange(existingRow, 2).setValue(linea);
      sheet.getRange(existingRow, 3).setValue(idMaquina);
      sheet.getRange(existingRow, 4).setValue(equipo);
      return { success: true, message: `Equipo ITEM ${itemInput} actualizado con éxito.` };
    } else {
      let maxItem = 0;
      data.forEach(r => {
        let it = parseInt(r[0]) || 0;
        if (it > maxItem) maxItem = it;
      });
      let newItem = maxItem + 1;
      sheet.appendRow([newItem, linea, idMaquina, equipo]);
      return { success: true, message: `Equipo ITEM ${newItem} (${equipo}) agregado con éxito.` };
    }
  } catch(e) {
    return { success: false, message: "Error al guardar equipo: " + e.toString() };
  }
}

// 🗑️ ELIMINAR EQUIPO (VALIDA CONTRASEÑA ADMIN)
function eliminarEquipo(item, pass) {
  if (pass !== ADMIN_PASS) {
    return { success: false, message: "Contraseña de Administrador incorrecta." };
  }
  try {
    let sheet = getOrCreateEquiposSheet();
    let itemTarget = parseInt(item);
    let data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues() : [];
    let targetRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (parseInt(data[i][0]) === itemTarget) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow > 1) {
      sheet.deleteRow(targetRow);
      return { success: true, message: `Equipo ITEM ${itemTarget} eliminado correctamente.` };
    } else {
      return { success: false, message: "No se encontró el ITEM en la base de datos." };
    }
  } catch(e) {
    return { success: false, message: "Error al eliminar equipo: " + e.toString() };
  }
}
// =========================================================================
// OBTENER CATÁLOGOS PÚBLICOS PARA FORMULARIO OPEN MTTO
// =========================================================================
function getPersonalListPublic() {
  try {
    let sheet = getOrCreatePersonalSheet();
    let data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues() : [];
    let list = data.map(r => ({
      noEmpleado: String(r[0]).trim(),
      nombre: String(r[1]).trim(),
      turno: String(r[2]).trim()
    }));
    return { success: true, personal: list };
  } catch(e) {
    return { success: false, personal: [] };
  }
}

function getEquiposListPublic() {
  try {
    let sheet = getOrCreateEquiposSheet();
    let data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues() : [];
    let list = data.map(r => ({
      item: r[0],
      linea: String(r[1]).trim(),
      idMaquina: String(r[2]).trim(),
      equipo: String(r[3]).trim()
    }));
    return { success: true, equipos: list };
  } catch(e) {
    return { success: false, equipos: [] };
  }
}
// =========================================================================
// MÓDULO PRODUCCIÓN - PERSISTENCIA Y SINCRONIZACIÓN EN TIEMPO REAL (15 COLS)
// =========================================================================
const PROD_SHEET_NAME = "Log_Produccion_Pizarron";

function getOrCreateProdSheet() {
  const ss = getDB();
  let sh = ss.getSheetByName(PROD_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PROD_SHEET_NAME);
    sh.appendRow(["Linea", "Fecha", "HoraIndex", "Meta", "ProduccionReal", "Hombre", "Maquina", "Material", "Metodo", "Calidad", "Produccion", "Otros", "Comentario", "UltimaActualizacion", "FALLA"]);
    sh.getRange("A1:O1").setFontWeight("bold").setBackground("#009B4A").setFontColor("white");
  }
  return sh;
}

// 🛠️ NORMALIZADOR BLINDADO CONTRA DESFASES DE ZONA HORARIA
function normalizarFechaJS(val) {
  if (!val) return "";
  var d, m, y;

  if (val instanceof Date) {
    // Extraer valores locales directamente del objeto sin Utilities.formatDate para evitar desfases GMT
    d = val.getDate();
    m = val.getMonth() + 1;
    y = val.getFullYear();
  } else {
    var str = String(val).trim().split(" ")[0]; // Quitar horas si las hay
    var parts = str.split(/[\/\-\.]/);
    if (parts.length >= 3) {
      if (parts[0].length === 4) { // Formato YYYY-MM-DD
        y = parseInt(parts[0], 10);
        m = parseInt(parts[1], 10);
        d = parseInt(parts[2], 10);
      } else { // Formato DD/MM/YYYY o D/M/YYYY
        d = parseInt(parts[0], 10);
        m = parseInt(parts[1], 10);
        y = parseInt(parts[2], 10);
      }
    }
  }

  if (d && m && y) {
    return d + "/" + m + "/" + y; // Estandariza siempre a "1/9/2026"
  }
  return String(val).trim();
}

// 💾 GUARDAR O ACTUALIZAR UNA HORA EN LA NUBE DESDE EL PIZARRÓN
function guardarDatoPizarron(datos) {
  try {
    let sh = getOrCreateProdSheet();
    let data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 15).getValues() : [];
    let targetRow = -1;
    let lineaStr = String(datos.linea).trim();
    let fechaNormTarget = normalizarFechaJS(datos.fecha);
    let horaIdx = parseInt(datos.horaIndex, 10);

    for (let i = 0; i < data.length; i++) {
      let lineaCelda = String(data[i][0]).trim();
      let fechaCeldaNorm = normalizarFechaJS(data[i][1]);
      let horaIdxCelda = parseInt(data[i][2], 10);

      if (lineaCelda === lineaStr && fechaCeldaNorm === fechaNormTarget && horaIdxCelda === horaIdx) {
        targetRow = i + 2;
        break;
      }
    }

    let timeZone = Session.getScriptTimeZone() || "GMT-6";
    let tsNow = Utilities.formatDate(new Date(), timeZone, "dd/MM/yyyy HH:mm:ss");
    let rowValues = [
      lineaStr, 
      datos.fecha, 
      horaIdx, 
      datos.meta, 
      datos.real,
      datos.hombre, 
      datos.maquina, 
      datos.material, 
      datos.metodo,
      datos.calidad, 
      datos.produccion, 
      datos.otros, 
      datos.comentario, 
      tsNow,
      datos.falla || ""
    ];

    if (targetRow > 1) {
      sh.getRange(targetRow, 1, 1, 15).setValues([rowValues]);
    } else {
      sh.appendRow(rowValues);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 📡 LECTURA EN TIEMPO REAL / HISTÓRICO PARA EL PIZARRÓN
function obtenerDatosPizarron(linea, fecha) {
  try {
    let sh = getOrCreateProdSheet();
    if (sh.getLastRow() <= 1) return { success: true, registros: [] };
    
    let data = sh.getRange(2, 1, sh.getLastRow() - 1, 15).getDisplayValues();
    let lineaStr = String(linea).trim();
    let fechaNormTarget = normalizarFechaJS(fecha);
    let result = [];

    for (let i = 0; i < data.length; i++) {
      let lineaCelda = String(data[i][0]).trim();
      let fechaCeldaNorm = normalizarFechaJS(data[i][1]);

      if (lineaCelda === lineaStr && fechaCeldaNorm === fechaNormTarget) {
        result.push({
          horaIndex: parseInt(data[i][2], 10),
          meta: parseInt(data[i][3], 10) || 0,
          real: data[i][4] !== undefined && data[i][4] !== null && data[i][4] !== "" ? data[i][4] : "",
          hombre: parseInt(data[i][5], 10) || 0,
          maquina: parseInt(data[i][6], 10) || 0,
          material: parseInt(data[i][7], 10) || 0,
          metodo: parseInt(data[i][8], 10) || 0,
          calidad: parseInt(data[i][9], 10) || 0,
          produccion: parseInt(data[i][10], 10) || 0,
          otros: parseInt(data[i][11], 10) || 0,
          comentario: String(data[i][12] || "").toUpperCase(),
          falla: String(data[i][14] || "").toUpperCase()
        });
      }
    }
    return { success: true, registros: result };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
// =========================================================================
// MÓDULO ADMIN PRODUCCIÓN - LECTURA DE CATÁLOGOS Y CONFIGURACIÓN DE METAS
// =========================================================================
const CONFIG_PROD_SHEET = "Config_Lineas_Produccion";

function getOrCreateConfigProdSheet() {
  const ss = getDB();
  let sh = ss.getSheetByName(CONFIG_PROD_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG_PROD_SHEET);
    sh.appendRow(["Linea", "Area", "Supervisor", "Turno", "M_07_08", "M_08_09", "M_09_10", "M_10_11", "M_11_12", "M_12_13", "M_13_14", "M_14_15", "M_15_16", "UltimaActualizacion"]);
    sh.getRange("A1:N1").setFontWeight("bold").setBackground("#009B4A").setFontColor("white");
  }
  return sh;
}

// 📋 LECTURA DE SUPERVISORES DE LA PESTAÑA 'Log_Supervisores'
function obtenerSupervisoresProduccion() {
  try {
    const ss = getDB();
    let sh = ss.getSheetByName("Log_Supervisores");
    if (!sh || sh.getLastRow() <= 1) return { success: true, supervisores: [] };
    let data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    let list = data.map(r => ({
      noEmpleado: String(r[0]).trim(),
      nombre: String(r[1]).trim(),
      turno: String(r[2]).trim()
    })).filter(x => x.nombre !== "");
    return { success: true, supervisores: list };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// 📋 LECTURA DE ÁREAS DE LA PESTAÑA 'Log_Areas'
function obtenerAreasProduccion() {
  try {
    const ss = getDB();
    let sh = ss.getSheetByName("Log_Areas");
    if (!sh || sh.getLastRow() <= 1) return { success: true, areas: [] };
    let data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    let list = data.map(r => String(r[0]).trim()).filter(x => x !== "" && x.toUpperCase() !== "AREAS");
    return { success: true, areas: list };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}



// 💾 GUARDAR LA CONFIGURACIÓN 24 HORAS DE UNA LÍNEA (ADMIN)
function guardarConfiguracionLineaAdmin(datos) {
  try {
    let sh = getOrCreateConfigProdSheet();
    // 🎯 Leemos 29 columnas (A hasta AC)
    let data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 29).getValues() : [];
    let lineaStr = String(datos.linea).trim();
    let targetRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === lineaStr) {
        targetRow = i + 2;
        break;
      }
    }

    let timeZone = Session.getScriptTimeZone() || "GMT-6";
    let tsNow = Utilities.formatDate(new Date(), timeZone, "dd/MM/yyyy HH:mm:ss");
    
    let metas = datos.metas || [];
    while(metas.length < 24) metas.push(0); // Asegura que siempre sean 24 horas

    // 🎯 Empaquetamos la fila (4 Info Base + 24 Metas + 1 Timestamp en AC)
    let rowValues = [lineaStr, datos.area, datos.supervisor, datos.turno];
    for (let m = 0; m < 24; m++) { 
        rowValues.push(metas[m]); 
    }
    rowValues.push(tsNow);

    if (targetRow > 1) {
      sh.getRange(targetRow, 1, 1, 29).setValues([rowValues]);
    } else {
      sh.appendRow(rowValues);
    }
    return { success: true, message: "¡Configuración de 24 Horas de la Línea " + lineaStr + " guardada exitosamente!" };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// 📡 LECTURA DE CONFIGURACIÓN 24 HORAS (ADMIN)
function obtenerConfiguracionLineaAdmin(numLinea) {
  try {
    let sh = getOrCreateConfigProdSheet();
    if (sh.getLastRow() <= 1) return { success: true, config: null };
    
    // 🎯 Leemos 29 columnas (4 info base + 24 horas + 1 timestamp)
    let data = sh.getRange(2, 1, sh.getLastRow() - 1, 29).getValues();
    let lineaStr = String(numLinea).trim();

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === lineaStr) {
        
        let metas24 = [];
        // Desempaquetamos las 24 columnas (empiezan en el índice 4 = Columna E)
        for (let m = 0; m < 24; m++) {
          metas24.push(parseInt(data[i][4 + m]) || 0);
        }

        return {
          success: true,
          config: {
            linea: lineaStr,
            area: String(data[i][1] || "DRYERS"),
            supervisor: String(data[i][2] || ""),
            turno: String(data[i][3] || ""),
            metas: metas24
          }
        };
      }
    }
    return { success: true, config: null };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}




// 📋 LECTURA DE TURNOS DE LA PESTAÑA 'Log_Turnos'
function obtenerTurnosProduccion() {
  try {
    const ss = getDB();
    let sh = ss.getSheetByName("Log_Turnos");
    if (!sh || sh.getLastRow() <= 1) return { success: true, turnos: [] };
    
    // 🎯 CLAVE: getDisplayValues() extrae el texto exacto ("07:00 - 17:06") sin romper el formato
    let data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues();
    
    let list = data.map(r => {
      // Filtramos celdas vacías y unimos con " / "
      let partes = r.filter(x => String(x).trim() !== "");
      return partes.join(" / ");
    }).filter(x => x !== "");
    
    return { success: true, turnos: list };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// =========================================================================
// MÓDULO SUPERVISOR PERFORMANCE - GENERAL, MULTI-ÁREA Y DINÁMICO
// =========================================================================
function obtenerDatosSupervisorPerformance() {
  try {
    let ss = getDB();
    let hoy = new Date();
    let fechaHoyNorm = normalizarFechaSegura(hoy);

    // 1. OBTENER CONFIGURACIÓN DE LÍNEAS ACTIVAS
    let shConfig = ss.getSheetByName("Config_Lineas_Produccion") || 
                   ss.getSheetByName("Configuracion_Lineas_Admin") || 
                   ss.getSheetByName("Log_Configuracion_Lineas");
                   
    if (!shConfig || shConfig.getLastRow() <= 1) return { success: true, ranking: [] };
    let configData = shConfig.getDataRange().getDisplayValues();

    let lineasConfig = {};
    for (let c = 1; c < configData.length; c++) {
      let r = configData[c];
      let numL = parseInt(String(r[0] || '').replace(/[^0-9]/g, ''), 10);
      if (isNaN(numL)) continue;

      let area = String(r[1] || "").trim();
      let supervisor = String(r[2] || "").trim().toUpperCase();
      let turno = String(r[3] || "").trim().toUpperCase();

      let invalidSup = ['', 'ASIGNADO POR ADMIN', 'SIN ASIGNAR', 'VACIO', 'VACÍO', 'N/A', 'N/D', '-', 'NINGUNO', 'SIN SUPERVISOR'];
      let invalidTurno = ['', 'VACIO', 'VACÍO', 'SIN ASIGNAR', 'OFF', 'DESACTIVADO', 'INACTIVO', 'N/A', '-', 'NO'];

      if (invalidSup.indexOf(supervisor) === -1 && invalidTurno.indexOf(turno) === -1) {
        let metas = [];
        for (let m = 0; m < 9; m++) {
          metas.push(parseInt(r[4 + m]) || 300);
        }
        lineasConfig[numL] = { area: area, supervisor: supervisor, metas: metas };
      }
    }

    // 2. OBTENER REGISTROS DE PRODUCCIÓN DE HOY
    let shProd = ss.getSheetByName("Log_Pizarron_Produccion") || ss.getSheetByName("Log_Produccion_Pizarron");
    if (!shProd || shProd.getLastRow() <= 1) return { success: true, ranking: [] };
    
    let prodData = shProd.getDataRange().getDisplayValues();
    let prodPorLineaHora = {};
    let logMetasHora = {};

    prodData.forEach(r => {
      let lStr = parseInt(String(r[0] || '').replace(/[^0-9]/g, ''), 10);
      let fNorm = normalizarFechaSegura(r[1]);
      let hIdx = parseInt(r[2], 10);
      let meta = parseInt(r[3], 10) || 0;
      let real = parseInt(r[4], 10) || 0;

      if (fNorm === fechaHoyNorm && lineasConfig[lStr] !== undefined) {
        if (!prodPorLineaHora[lStr]) prodPorLineaHora[lStr] = {};
        if (!logMetasHora[lStr]) logMetasHora[lStr] = {};
        
        prodPorLineaHora[lStr][hIdx] = real;
        if (meta > 0) logMetasHora[lStr][hIdx] = meta;
      }
    });

    // 3. AGRUPAR LÍNEAS POR SUPERVISOR
    let supMap = {};
    Object.keys(lineasConfig).forEach(lStr => {
      let l = parseInt(lStr, 10);
      let sup = lineasConfig[l].supervisor;
      if (!supMap[sup]) supMap[sup] = [];
      supMap[sup].push(l);
    });

    // 4. CALCULAR PERFORMANCE PROMEDIADO
    let ranking = [];
    Object.keys(supMap).forEach(sup => {
      let lineasSup = supMap[sup];
      let perfHorasSup = [];
      let sumaPromHoras = 0;
      let casillasMayorACeroContador = 0;

      for (let h = 0; h < 9; h++) {
        let sumaPerfHora = 0;
        let lineasContadas = 0;

        lineasSup.forEach(l => {
          let meta = (logMetasHora[l] && logMetasHora[l][h] > 0) ? logMetasHora[l][h] : (lineasConfig[l].metas[h] || 300);
          let real = (prodPorLineaHora[l] && prodPorLineaHora[l][h] !== undefined) ? prodPorLineaHora[l][h] : 0;
          let perf = meta > 0 ? (real / meta) * 100 : 0;

          sumaPerfHora += perf;
          lineasContadas++;
        });

        let perfPromHoraLineas = lineasContadas > 0 ? (sumaPerfHora / lineasContadas) : 0;
        let perfHoraRedondeado = Math.round(perfPromHoraLineas);
        perfHorasSup.push(perfHoraRedondeado);

        if (perfHoraRedondeado > 0) {
          sumaPromHoras += perfHoraRedondeado;
          casillasMayorACeroContador++;
        }
      }

      let perfGlobal = casillasMayorACeroContador > 0 ? Math.round(sumaPromHoras / casillasMayorACeroContador) : 0;

      ranking.push({
        supervisor: sup,
        lineas: lineasSup.join(", "),
        cantLineas: lineasSup.length,
        perfGlobal: perfGlobal,
        perfHoras: perfHorasSup,
        horasReportadas: casillasMayorACeroContador
      });
    });

    ranking.sort((a, b) => b.perfGlobal - a.perfGlobal);

    return { success: true, ranking: ranking, fecha: fechaHoyNorm };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =========================================================================
// MÓDULO KPI PRODUCCIÓN GLOBAL - CÁLCULO EXACTO DE DT Y FECHA ROBUSTA (CODE.GS)
// =========================================================================
function obtenerKpisProduccionData() {
  try {
    let ss = getDB();
    let hoy = new Date();
    let fechaHoyNorm = normalizarFechaJS(hoy);

    // 1. LEER METAS CONFIGURADAS POR LÍNEA
    let shConfig = ss.getSheetByName(CONFIG_PROD_SHEET);
    let lineasConfig = {};
    if (shConfig && shConfig.getLastRow() > 1) {
      let configData = shConfig.getRange(2, 1, shConfig.getLastRow() - 1, 14).getValues();
      configData.forEach(r => {
        let linea = String(r[0]).trim();
        if (linea !== "") {
          lineasConfig[linea] = {
            metas: [
              parseInt(r[4]) || 300, parseInt(r[5]) || 300, parseInt(r[6]) || 200,
              parseInt(r[7]) || 300, parseInt(r[8]) || 300, parseInt(r[9]) || 150,
              parseInt(r[10]) || 300, parseInt(r[11]) || 300, parseInt(r[12]) || 150
            ]
          };
        }
      });
    }

    // 2. LEER REGISTROS DE PRODUCCIÓN DE HOY (VALORES VISIBLES BLINDADOS)
    let shProd = getOrCreateProdSheet();
    let prodData = shProd.getLastRow() > 1 ? shProd.getRange(2, 1, shProd.getLastRow() - 1, 14).getDisplayValues() : [];

    let totalRealProd = 0;
    let totalMetaProd = 0;
    let totalDtMinutos = 0;
    let totalTiempoDisponibleMinutos = 0;

    let registrosHoy = {};
    prodData.forEach(r => {
      let lStr = String(r[0]).trim();
      let fNorm = normalizarFechaJS(r[1]);
      let hIdx = parseInt(r[2], 10);
      let real = parseInt(r[4], 10) || 0;
      let suma6M = (parseInt(r[5])||0) + (parseInt(r[6])||0) + (parseInt(r[7])||0) + 
                   (parseInt(r[8])||0) + (parseInt(r[9])||0) + (parseInt(r[10])||0) + (parseInt(r[11])||0);
      let com = String(r[12] || "").trim();

      if (fNorm === fechaHoyNorm) {
        if (!registrosHoy[lStr]) registrosHoy[lStr] = {};
        
        let metaConfig = (lineasConfig[lStr] && lineasConfig[lStr].metas[hIdx] !== undefined) 
                         ? lineasConfig[lStr].metas[hIdx] 
                         : (parseInt(r[3]) || 300);

        // Hora activa/corriendo si tiene producción real > 0, 6M > 0 o comentario
        let esActiva = (real > 0) || (suma6M > 0) || (com !== "" && com !== "-");

        // 🎯 CÁLCULO DE DT: Toma los 6M o calcula los minutos perdidos por faltante de producción
        let dtCalculado = 0;
        if (suma6M > 0) {
          dtCalculado = suma6M;
        } else if (esActiva && real < metaConfig && metaConfig > 0) {
          dtCalculado = Math.round((metaConfig - real) * (60 / metaConfig));
        }

        registrosHoy[lStr][hIdx] = {
          real: real,
          dt: dtCalculado,
          meta: metaConfig,
          esActiva: esActiva
        };
      }
    });

    // 3. SUMATORIA EXCLUSIVA DE HORAS CORRIENDO / ACTIVAS
    Object.keys(registrosHoy).forEach(lStr => {
      Object.keys(registrosHoy[lStr]).forEach(hIdx => {
        let reg = registrosHoy[lStr][hIdx];
        if (reg.esActiva) {
          totalRealProd += reg.real;
          totalMetaProd += reg.meta;
          totalDtMinutos += reg.dt;
          totalTiempoDisponibleMinutos += 60; // +60 min de Tiempo Disponible por cada hora activa
        }
      });
    });

    // 4. CÁLCULO DE PORCENTAJES Y METRICAS
    let perfPct = totalMetaProd > 0 ? ((totalRealProd / totalMetaProd) * 100) : 0;
    let deltaProd = totalRealProd - totalMetaProd;

    let dtPct = totalTiempoDisponibleMinutos > 0 ? ((totalDtMinutos / totalTiempoDisponibleMinutos) * 100) : 0;
    let dtHoras = (totalDtMinutos / 60).toFixed(1);
    let tiempoDispHoras = (totalTiempoDisponibleMinutos / 60).toFixed(1);

    return {
      success: true,
      perf: {
        pct: parseFloat(perfPct.toFixed(1)),
        plan: totalMetaProd,
        real: totalRealProd,
        delta: deltaProd,
        cumple: perfPct >= 85
      },
      dt: {
        pct: parseFloat(dtPct.toFixed(1)),
        tiempoDisponibleHoras: tiempoDispHoras,
        dtMinutos: totalDtMinutos,
        dtHoras: dtHoras,
        cumple: dtPct <= 2.5
      },
      fecha: fechaHoyNorm
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}








// =========================================================================
// MÓDULO DASHBOARD DE PRODUCCIÓN - DIAGNÓSTICO DE TIEMPOS MUERTOS (SIN DUPLICADOS)
// =========================================================================
function obtenerDatosDashboardProduccion(fIni, fFin, esHoy) {
  try {
    var ss = getDB();
    var hoy = new Date();
    var fechaHoyNorm = normalizarFechaJS(hoy);
    var esHoyCalculado = (esHoy !== undefined) ? !!esHoy : true;

    var normIni = (!esHoyCalculado && fIni) ? normalizarFechaJS(fIni) : null;
    var normFin = (!esHoyCalculado && fFin) ? normalizarFechaJS(fFin) : null;

    // 1. CONFIGURACIÓN DE LÍNEAS
    var shConfig = ss.getSheetByName(CONFIG_PROD_SHEET);
    var lineasConfig = {};
    if (shConfig && shConfig.getLastRow() > 1) {
      var configData = shConfig.getRange(2, 1, shConfig.getLastRow() - 1, 14).getValues();
      configData.forEach(function(r) {
        var linea = String(r[0]).trim();
        if (linea !== "") {
          lineasConfig[linea] = {
            area: String(r[1] || "DRYERS").trim(),
            supervisor: String(r[2] || "Sin Asignar").trim(),
            turno: String(r[3] || "Turno D").trim()
          };
        }
      });
    }

    // 2. REGISTROS DE PRODUCCIÓN Y PAROS
    var shProd = getOrCreateProdSheet();
    var lastRow = shProd.getLastRow();
    var fechaLabel = esHoyCalculado ? "HOY" : "RANGO SELECCIONADO";

    if (lastRow <= 1) return { success: true, ranking: [], fechaFiltroLabel: fechaLabel };

    var maxCols = Math.min(shProd.getLastColumn(), 15);
    var prodData = shProd.getRange(2, 1, lastRow - 1, maxCols).getDisplayValues();

    var horasLabel = ["07:00-08:00", "08:00-09:00", "09:00-10:00", "10:00-11:00", "11:00-12:00", "12:00-13:00", "13:00-14:00", "14:00-15:00", "15:00-16:00"];

    // Map unificado para eliminar duplicados de la misma hora
    var mapaHorasUnicas = {};

    for (var i = 0; i < prodData.length; i++) {
      var r = prodData[i];
      var lStr = String(r[0] || "").trim();
      var fecCeldaNorm = normalizarFechaJS(r[1]);
      var hIdx = parseInt(r[2], 10);

      var coincide = false;

      if (esHoyCalculado) {
        coincide = (fecCeldaNorm === fechaHoyNorm);
      } else {
        if (normIni && normFin) {
          var tsFila = parseFechaTimestampNorm(fecCeldaNorm);
          var tsIni = parseFechaTimestampNorm(normIni);
          var tsFin = parseFechaTimestampNorm(normFin);
          coincide = (tsFila !== null && tsFila >= tsIni && tsFila <= tsFin);
        } else {
          coincide = true;
        }
      }

      if (lStr !== "" && !isNaN(hIdx) && coincide) {
        var claveUnica = lStr + "_" + hIdx + "_" + fecCeldaNorm;
        
        mapaHorasUnicas[claveUnica] = {
          linea: lStr,
          fecha: r[1] || "",
          horaIdx: hIdx,
          hombre: parseInt(r[5], 10) || 0,
          maquina: parseInt(r[6], 10) || 0,
          material: parseInt(r[7], 10) || 0,
          metodo: parseInt(r[8], 10) || 0,
          calidad: parseInt(r[9], 10) || 0,
          produccion: parseInt(r[10], 10) || 0,
          otros: parseInt(r[11], 10) || 0,
          comentario: String(r[12] || "").trim(),
          falla: (r.length >= 15) ? String(r[14] || "").trim() : ""
        };
      }
    }

    // 3. CONSOLIDACIÓN DE RESULTADOS SIN SOBRECONTEO
    var lineasDataMap = {};

    Object.keys(mapaHorasUnicas).forEach(function(k) {
      var item = mapaHorasUnicas[k];
      var lStr = item.linea;

      if (!lineasDataMap[lStr]) {
        var cfg = lineasConfig[lStr] || { area: "DRYERS", supervisor: "Sin Asignar", turno: "Turno D" };
        lineasDataMap[lStr] = {
          linea: lStr,
          area: cfg.area,
          supervisor: cfg.supervisor,
          turno: cfg.turno,
          hombre: 0, maquina: 0, material: 0, metodo: 0, calidad: 0, produccion: 0, otros: 0, totalDT: 0,
          detalles: []
        };
      }

      var dtFila = item.hombre + item.maquina + item.material + item.metodo + item.calidad + item.produccion + item.otros;

      lineasDataMap[lStr].hombre += item.hombre;
      lineasDataMap[lStr].maquina += item.maquina;
      lineasDataMap[lStr].material += item.material;
      lineasDataMap[lStr].metodo += item.metodo;
      lineasDataMap[lStr].calidad += item.calidad;
      lineasDataMap[lStr].produccion += item.produccion;
      lineasDataMap[lStr].otros += item.otros;
      lineasDataMap[lStr].totalDT += dtFila;

      var com = item.comentario;
      var fal = item.falla;

      if (dtFila > 0 || (com !== "" && com !== "-") || (fal !== "" && fal !== "-")) {
        lineasDataMap[lStr].detalles.push({
          hora: horasLabel[item.horaIdx] || ("Hora " + (item.horaIdx + 1)),
          fecha: item.fecha,
          hombre: item.hombre, maquina: item.maquina, material: item.material, metodo: item.metodo, calidad: item.calidad, produccion: item.produccion, otros: item.otros,
          dtTotal: dtFila,
          comentario: (com && com !== "-") ? com : "-",
          falla: (fal && fal !== "-") ? fal : "-"
        });
      }
    });

    var rankingLineas = Object.values(lineasDataMap);
    rankingLineas.sort(function(a, b) { return b.totalDT - a.totalDT; });

    return { 
      success: true, 
      ranking: rankingLineas, 
      fechaFiltroLabel: fechaLabel 
    };
  } catch (e) {
    return { success: false, error: e.toString(), ranking: [] };
  }
}











// =========================================================================
// PROGRAMAR MTTO (BOTÓN NARANJA) - GUARDA PLAN EN MANTENIMIENTOS ABIERTOS
// =========================================================================
function guardarPlanMttoCard(datos) {
  try {
    const ss = getDB();
    let sheetOpen = ss.getSheetByName(OPEN_SHEET_NAME);
    let folioTarget = String(datos.folio || "").trim().toUpperCase();

    if (!sheetOpen) {
      return { success: false, message: "No se encontró la pestaña MTTOs Abiertos." };
    }

    // 1. Asegurar encabezados de plan en columnas L a Q (Cols 12 a 17)
    let maxCols = Math.max(17, sheetOpen.getLastColumn());
    let headers = sheetOpen.getRange(1, 1, 1, maxCols).getValues()[0];
    if (!headers[11]) sheetOpen.getRange(1, 12).setValue("TiempoSolicitado").setFontWeight("bold");
    if (!headers[12]) sheetOpen.getRange(1, 13).setValue("TecNoEmp").setFontWeight("bold");
    if (!headers[13]) sheetOpen.getRange(1, 14).setValue("TecnicoAsignado").setFontWeight("bold");
    if (!headers[14]) sheetOpen.getRange(1, 15).setValue("SupNoEmp").setFontWeight("bold");
    if (!headers[15]) sheetOpen.getRange(1, 16).setValue("SupervisorAsignado").setFontWeight("bold");
    if (!headers[16]) sheetOpen.getRange(1, 17).setValue("ActividadProgramada").setFontWeight("bold");

    // 2. Buscar la fila del folio en MTTOs Abiertos y actualizar valores sin cerrar ni borrar
    if (sheetOpen.getLastRow() > 1) {
      let dataOpen = sheetOpen.getRange(2, 1, sheetOpen.getLastRow() - 1, maxCols).getValues();
      for (let i = 0; i < dataOpen.length; i++) {
        if (String(dataOpen[i][0]).trim().toUpperCase() === folioTarget) {
          let row = i + 2;
          sheetOpen.getRange(row, 12).setValue(datos.tiempoSolicitado);
          sheetOpen.getRange(row, 13).setValue(datos.tecNoEmp);
          sheetOpen.getRange(row, 14).setValue(datos.tecnicoAsignado);
          sheetOpen.getRange(row, 15).setValue(datos.supNoEmp);
          sheetOpen.getRange(row, 16).setValue(datos.supervisorAsignado);
          sheetOpen.getRange(row, 17).setValue(datos.actividadProgramada);

          return { success: true, message: "Mantenimiento programado y asignado exitosamente." };
        }
      }
    }

    return { success: false, message: "No se encontró el Folio " + folioTarget + " en MTTOs Abiertos." };
  } catch (e) {
    return { success: false, message: "Error en servidor: " + e.toString() };
  }
}
// =========================================================================
// MÓDULO ASIGNAR Y CERRAR - LECTURA DE CATÁLOGOS Y TIEMPOS
// =========================================================================
function obtenerCatalogosCierreMtto() {
  try {
    const ss = getDB();
    
    // 1. Personal (Técnicos y Supervisores)
    let shPers = ss.getSheetByName("Log_Personal");
    let personal = [];
    if (shPers && shPers.getLastRow() > 1) {
      let dataP = shPers.getRange(2, 1, shPers.getLastRow() - 1, 3).getValues();
      personal = dataP.map(r => ({
        noEmpleado: String(r[0]).trim(),
        nombre: String(r[1]).trim(),
        turno: String(r[2]).trim()
      }));
    }

    // 2. Equipos
    let shEq = ss.getSheetByName("Log_Equipos");
    let equipos = [];
    if (shEq && shEq.getLastRow() > 1) {
      let dataE = shEq.getRange(2, 1, shEq.getLastRow() - 1, 4).getValues();
      equipos = dataE.map(r => ({
        item: r[0],
        linea: String(r[1]).trim(),
        idMaquina: String(r[2]).trim(),
        equipo: String(r[3]).trim()
      }));
    }

    return { success: true, personal: personal, equipos: equipos };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
// =========================================================================
// PERSISTENCIA DEL PLAN Y GUARDADO DIRECTO EN LOG CERRADOS
// =========================================================================
function guardarPlanMttoCard(datos) {
  try {
    const ss = getDB();
    let sheetOpen = ss.getSheetByName(OPEN_SHEET_NAME);
    let sheetLog = ss.getSheetByName(LOG_SHEET_NAME);
    let folioTarget = String(datos.folio || "").trim().toUpperCase();

    if (!sheetLog) {
      return { success: false, message: "No existe la pestaña Log Cerrados en la base de datos." };
    }

    // 1. Crear encabezados adicionales en Log Cerrados si no existen (Columnas O, P, Q)
    let maxCols = Math.max(17, sheetLog.getLastColumn());
    let headers = sheetLog.getRange(1, 1, 1, maxCols).getValues()[0];
    if (!headers[14]) sheetLog.getRange(1, 15).setValue("TiempoSolicitado").setFontWeight("bold");
    if (!headers[15]) sheetLog.getRange(1, 16).setValue("TecNoEmp").setFontWeight("bold");
    if (!headers[16]) sheetLog.getRange(1, 17).setValue("SupNoEmp").setFontWeight("bold");

    // 2. Si el Folio está en MTTOs Abiertos, extraemos su información de origen y eliminamos la fila
    let qInicial = "Dato Registrado", dInicial = datos.maquina || "", qAbrio = "N/A", evInicial = "", tsApertura = Date.now();
    if (sheetOpen && sheetOpen.getLastRow() > 1) {
      let dataOpen = sheetOpen.getRange(2, 1, sheetOpen.getLastRow() - 1, sheetOpen.getLastColumn()).getValues();
      for (let i = 0; i < dataOpen.length; i++) {
        if (String(dataOpen[i][0]).trim().toUpperCase() === folioTarget) {
          qInicial = dataOpen[i][2] || "Dato Registrado";
          dInicial = dataOpen[i][4] || datos.maquina;
          qAbrio = dataOpen[i][6] || "N/A";
          evInicial = dataOpen[i][9] || "";
          tsApertura = dataOpen[i][10] || Date.now();
          
          sheetOpen.deleteRow(i + 2); // Remueve de Abiertos
          break;
        }
      }
    }

    // 3. Revisar si el Folio ya existe en Log Cerrados para actualizar o insertar
    let dataLog = sheetLog.getLastRow() > 1 ? sheetLog.getRange(2, 1, sheetLog.getLastRow() - 1, sheetLog.getLastColumn()).getValues() : [];
    let existingRow = -1;
    for (let i = 0; i < dataLog.length; i++) {
      if (String(dataLog[i][0]).trim().toUpperCase() === folioTarget) {
        existingRow = i + 2;
        break;
      }
    }

    let tsCierre = Date.now();
    let fechaCierreStr = new Date(tsCierre).toLocaleString('es-MX', { hour12: false });

    if (existingRow > 1) {
      // Actualizar registro existente
      sheetLog.getRange(existingRow, 3).setValue(datos.tipo || "Correctivo (Fugas)");
      sheetLog.getRange(existingRow, 4).setValue(datos.supervisorAsignado);
      sheetLog.getRange(existingRow, 5).setValue(datos.tecnicoAsignado);
      sheetLog.getRange(existingRow, 6).setValue(datos.maquina || "");
      sheetLog.getRange(existingRow, 7).setValue(datos.actividadProgramada);
      sheetLog.getRange(existingRow, 15).setValue(datos.tiempoSolicitado);
      sheetLog.getRange(existingRow, 16).setValue(datos.tecNoEmp);
      sheetLog.getRange(existingRow, 17).setValue(datos.supNoEmp);
    } else {
      // Insertar nueva fila completa en Log Cerrados
      sheetLog.appendRow([
        datos.folio,
        fechaCierreStr,
        datos.tipo || "Correctivo (Fugas)",
        datos.supervisorAsignado,
        datos.tecnicoAsignado,
        datos.maquina || "",
        datos.actividadProgramada,
        "", // Evidencia Final
        qInicial,
        dInicial,
        qAbrio,
        evInicial,
        tsApertura,
        tsCierre,
        datos.tiempoSolicitado,
        datos.tecNoEmp,
        datos.supNoEmp
      ]);
    }

    // 4. Si vincularon una refacción, actualizar la pestaña Log Refacciones
    if (datos.folioRefVinc) {
      let refSheet = ss.getSheetByName(REF_SHEET_NAME);
      if (refSheet && refSheet.getLastRow() > 1) {
        let refData = refSheet.getRange(2, 1, refSheet.getLastRow() - 1, 9).getValues();
        for (let i = 0; i < refData.length; i++) {
          if (String(refData[i][0]).trim().toUpperCase() === String(datos.folioRefVinc).trim().toUpperCase()) {
            refSheet.getRange(i + 2, 2).setValue(datos.folio);
            refSheet.getRange(i + 2, 9).setValue("REFACCIONES UTILIZADAS");
            break;
          }
        }
      }
    }

    return { success: true, message: "¡Mantenimiento guardado exitosamente en Log Cerrados!" };
  } catch (e) {
    return { success: false, message: "Error en servidor: " + e.toString() };
  }
}
/**
 * Registra el tiempo de inicio de la alarma EHS en la columna "TIEMPO 1 P1 I"
 */
function registrarInicioAlarmaEHS() {
  try {
    let ss = getDB();
    let sheet = ss.getSheetByName("REGISTRO DE PRIMEROS AUXILIOS");
    if (!sheet) return { success: false, error: "Pestaña 'REGISTRO DE PRIMEROS AUXILIOS' no encontrada." };

    let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    let colInicioIdx = -1;

    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim().toUpperCase() === "TIEMPO 1 P1 I") {
        colInicioIdx = i + 1;
        break;
      }
    }

    if (colInicioIdx === -1) {
      return { success: false, error: "Columna 'TIEMPO 1 P1 I' no encontrada." };
    }

    let nextRow = sheet.getLastRow() + 1;
    if (nextRow < 2) nextRow = 2;

    let fechaInicio = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT-6", "dd/MM/yyyy HH:mm:ss");

    sheet.getRange(nextRow, colInicioIdx).setValue(fechaInicio);

    return { success: true, rowId: nextRow, timestamp: fechaInicio };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
/**
 * Registra el tiempo de inicio de la alarma EHS en la columna "TIEMPO 1 P1 I"
 */
function registrarInicioAlarmaEHS() {
  try {
    let ss = getDB();
    let targetTabName = "REGISTRO DE PRIMEROS AUXILIOS";
    let sheet = null;
    let sheets = ss.getSheets();

    // Búsqueda flexible tolerante a espacios extras o minúsculas
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().trim().toUpperCase() === targetTabName) {
        sheet = sheets[i];
        break;
      }
    }

    // Autocreación de emergencia si la pestaña no existe en la hoja activa
    if (!sheet) {
      sheet = ss.insertSheet(targetTabName);
      sheet.appendRow(["TIEMPO 1 P1 I", "TIEMPO 1 P1 F"]);
      sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#dc2626").setFontColor("white");
    }

    let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    let colInicioIdx = -1;

    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim().toUpperCase() === "TIEMPO 1 P1 I") {
        colInicioIdx = i + 1;
        break;
      }
    }

    if (colInicioIdx === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue("TIEMPO 1 P1 I");
      colInicioIdx = sheet.getLastColumn();
    }

    let nextRow = sheet.getLastRow() + 1;
    if (nextRow < 2) nextRow = 2;

    let fechaInicio = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT-6", "dd/MM/yyyy HH:mm:ss");
    sheet.getRange(nextRow, colInicioIdx).setValue(fechaInicio);

    return { success: true, rowId: nextRow, timestamp: fechaInicio };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Registra el tiempo final de la alarma EHS en la columna "TIEMPO 1 P1 F"
 */
function registrarFinAlarmaEHS(rowId) {
  try {
    let ss = getDB();
    let targetTabName = "REGISTRO DE PRIMEROS AUXILIOS";
    let sheet = null;
    let sheets = ss.getSheets();

    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().trim().toUpperCase() === targetTabName) {
        sheet = sheets[i];
        break;
      }
    }

    if (!sheet) return { success: false, error: "Pestaña 'REGISTRO DE PRIMEROS AUXILIOS' no encontrada." };

    let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    let colFinIdx = -1;

    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim().toUpperCase() === "TIEMPO 1 P1 F") {
        colFinIdx = i + 1;
        break;
      }
    }

    if (colFinIdx === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue("TIEMPO 1 P1 F");
      colFinIdx = sheet.getLastColumn();
    }

    let targetRow = parseInt(rowId, 10) || sheet.getLastRow();
    if (targetRow < 2) targetRow = 2;

    let fechaFin = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT-6", "dd/MM/yyyy HH:mm:ss");
    sheet.getRange(targetRow, colFinIdx).setValue(fechaFin);

    return { success: true, timestamp: fechaFin };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
/**
 * Convierte los audios de Google Drive a Base64 para reproducción en el navegador
 */
function obtenerAudiosBase64EHS() {
  var idsMap = {
    p02_inicio: "1htsrj_kwPSnGiAbp9k-AxG_91bQ-Xo3M",         // Pasos 2
    p03_tiempo_obj: "1AwLymC9CS1st7-OLQZcomzXAXHvvMIek",     // Pasos 3, 9
    p04_actividades: "1Sm6VblnnqLHua5ulK7u6P27poArylszp",    // Pasos 4
    p05_notificar: "1ZP2W5_bmHqP4wHWNsDyy_u9fHJqzyZoo",      // Pasos 5
    p06_operador: "1nCC4Ok2zXQibeBuv8h5xfjIuJ4dWd_dl",       // Pasos 6
    p11_todolisto: "1yc2pESbTwF2Mk9WCorYHMq5KFbgcl8P3",      // Pasos 11
    p13_rest1min: "14AAgir9tzx4Bd-kTtjmRkraHeFtnNnCO",       // Pasos 13, 14
    p19_rest30s: "1_ydnAbPFmXK5Uc4ED3-vtYQiAgWmIGn4",        // Pasos 19
    p24_fintiempo: "1Y4kOmeEWxCz-Txw1y9W51lupYCv1j8Yf",      // Pasos 24
    p26_excedido_proto: "1IZrLoQ4PHIZrpQjtn2Q76qM_oS8P-YBP", // Pasos 26, 28, 30, 32, 35, 37, 39, 42, 44, 47, 51, 53, 55
    p33_excedido30: "1pQS52Y_y5yFKkYoynwcWXmUBQ_W0JSKi",     // Pasos 33
    p40_excedido60: "1wWkx_UAct6PkQxAkgzaRS0lCkxQFL1EN",     // Pasos 40
    p49_excedido90: "1CtHKMaSdCuEdngwktDTjCLWXjN63blUT",     // Pasos 49
    p56_excedido120: "1HxftvGMbFbd5r2fRCFS8KmXsn5KX5kdX",    // Pasos 56
    p57_desactiva4m: "1VHnlVjmMxffnd1VC25Bp-cImC6Fyo5zy",    // Pasos 57
    p58_cancela: "1iduaIp5p8efmVvk788VGi6MTQ11qkM6f"         // Pasos 58 (Cancelación anticipada)
  };

  var base64Audios = {
    alarm: "https://actions.google.com/sounds/v1/alarms/alarm_clock.ogg"
  };

  for (var key in idsMap) {
    try {
      var id = idsMap[key];
      var file = DriveApp.getFileById(id);
      var blob = file.getBlob();
      var mime = blob.getContentType() || "audio/mp3";
      base64Audios[key] = "data:" + mime + ";base64," + Utilities.base64Encode(blob.getBytes());
    } catch (e) {
      console.error("Error cargando audio " + key + ": " + e.toString());
    }
  }

  return base64Audios;
}
// =========================================================================
// MÓDULO PASE ENTRADA / SALIDA RRHH (FA-RH-029) - BACKEND REVISADO
// =========================================================================

function obtenerAsuntosRRHH() {
  try {
    var ss = getDB(); // 👈 FIX 1: Conector maestro a prueba de fallos
    var sheet = ss.getSheetByName("ASUNTOS RRHH");
    if (!sheet) return { success: false, asuntos: [] };
    
    // 👈 FIX 2: Lectura rápida que no explota con rangos invertidos
    var data = sheet.getDataRange().getValues();
    var asuntosValidos = [];
    
    // Empezamos en i = 1 para saltar el encabezado
    for (var i = 1; i < data.length; i++) {
      var valorAsunto = String(data[i][0] || "").trim();
      if (valorAsunto !== "") {
        asuntosValidos.push(valorAsunto);
      }
    }
    
    return { success: true, asuntos: asuntosValidos };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function buscarGenericoRRHH(nombrePestana, noEmpleado) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(nombrePestana);
    if (!sheet) return { success: false, message: "Pestaña '" + nombrePestana + "' no encontrada" };
    
    var data = sheet.getDataRange().getValues();
    var numBuscado = String(noEmpleado || "").trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === numBuscado) {
        return {
          success: true,
          encontrado: true,
          noEmpleado: String(data[i][0] || ""),
          nombre: String(data[i][1] || ""),
          depto: String(data[i][2] || ""),
          turno: String(data[i][3] || "")
        };
      }
    }
    return { success: true, encontrado: false };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function buscarEmpleadoRRHH(no) { return buscarGenericoRRHH("EMPLEADOS RRHH", no); }
function buscarSupervisorRRHH(no) { return buscarGenericoRRHH("SUPERVISORES RRHH", no); }
function buscarGuardiaRRHH(no) { return buscarGenericoRRHH("GURDIAS RRHH", no); }

function MapeoColumnasPase(headersRow) {
  var map = { folio:0, estatus:1, fechaReg:2, tsCreacion:3, tipoPase:4, horaEntrada:5, horaSalida:6, fechaPase:7, trabajador:8, noEmp:9, depto:10, turno:11, asunto:12, firmaTrabajador:13, origNombre:14, origNoEmp:15, firmaOrig:16, supNombre:17, supNoEmp:18, firmaSup:19, guardiaNombre:20, guardiaNoEmp:21, firmaGuardia:22 };
  if (!headersRow || !headersRow.length) return map;

  for (var c = 0; c < headersRow.length; c++) {
    var h = String(headersRow[c] || "").trim().toUpperCase();
    if (h === "FOLIO") map.folio = c;
    else if (h === "ESTATUS") map.estatus = c;
    else if (h.indexOf("FECHA REGISTRO") !== -1) map.fechaReg = c;
    else if (h.indexOf("TS_CREACION") !== -1 || h.indexOf("TS CREACION") !== -1) map.tsCreacion = c;
    else if (h.indexOf("TIPO") !== -1) map.tipoPase = c;
    else if (h.indexOf("ENTRADA") !== -1 && h.indexOf("HORA") !== -1) map.horaEntrada = c;
    else if (h.indexOf("SALIDA") !== -1 && h.indexOf("HORA") !== -1) map.horaSalida = c;
    else if (h === "FECHA PASE" || (h.indexOf("FECHA") !== -1 && h.indexOf("REGISTRO") === -1)) map.fechaPase = c;
    else if (h === "TRABAJADOR") map.trabajador = c;
    else if (h === "NO. EMPLEADO" || h === "NO EMPLEADO") map.noEmp = c;
    else if (h.indexOf("DEPTO") !== -1) map.depto = c;
    else if (h.indexOf("TURNO") !== -1) map.turno = c;
    else if (h.indexOf("ASUNTO") !== -1) map.asunto = c;
    else if (h.indexOf("FIRMA TRABAJADOR") !== -1) map.firmaTrabajador = c;
    else if (h.indexOf("ORIGINADO POR NOMBRE") !== -1 || h.indexOf("ORIGINADO NOMBRE") !== -1) map.origNombre = c;
    else if (h.indexOf("ORIGINADO POR NO") !== -1 || h.indexOf("ORIGINADO NO") !== -1) map.origNoEmp = c;
    else if (h.indexOf("FIRMA ORIGINADO") !== -1) map.firmaOrig = c;
    else if (h.indexOf("SUPERVISOR NOMBRE") !== -1) map.supNombre = c;
    else if (h.indexOf("SUPERVISOR NO") !== -1) map.supNoEmp = c;
    else if (h.indexOf("FIRMA SUPERVISOR") !== -1) map.firmaSup = c;
    else if (h.indexOf("GUARDIA NOMBRE") !== -1) map.guardiaNombre = c;
    else if (h.indexOf("GUARDIA NO") !== -1) map.guardiaNoEmp = c;
    else if (h.indexOf("FIRMA GUARDIA") !== -1) map.firmaGuardia = c;
  }
  return map;
}

function guardarPaseInicialRRHH(datos) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PASE RRHH");
    var headersOficiales = ["FOLIO", "ESTATUS", "FECHA REGISTRO", "TS_CREACION", "TIPO PASE", "HORA ENTRADA", "HORA SALIDA", "FECHA PASE", "TRABAJADOR", "NO. EMPLEADO", "DEPTO", "TURNO", "ASUNTO", "FIRMA TRABAJADOR", "ORIGINADO POR NOMBRE", "ORIGINADO POR NO. EMP", "FIRMA ORIGINADO", "SUPERVISOR NOMBRE", "SUPERVISOR NO. EMP", "FIRMA SUPERVISOR", "GUARDIA NOMBRE", "GUARDIA NO. EMP", "FIRMA GUARDIA"];

    if (!sheet) { sheet = ss.insertSheet("PASE RRHH"); sheet.appendRow(headersOficiales); } 
    else {
      var currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
      if (currentHeaders.length < 5 || String(currentHeaders[0] || "").toUpperCase() !== "FOLIO") {
        sheet.insertRowBefore(1); sheet.getRange(1, 1, 1, headersOficiales.length).setValues([headersOficiales]);
      }
    }
    
    var numFolio = sheet.getLastRow() > 1 ? sheet.getLastRow() : 1;
    var folioStr = "PAS-" + new Date().getFullYear() + "-" + ("000" + numFolio).slice(-4);
    var ahoraTs = Date.now();
    var ahoraFechaStr = new Date().toLocaleString('es-MX', { dateStyle: 'short', timeStyle: 'medium' });

    sheet.appendRow([folioStr, "PENDIENTE SUPERVISOR", ahoraFechaStr, ahoraTs, datos.tipoPase||"", datos.horaEntrada||"", datos.horaSalida||"", datos.fechaPase||"", datos.trabajadorNombre||"", datos.trabajadorNoEmp||"", datos.depto||"", datos.turno||"", datos.asunto||"", datos.firmaTrabajador||"", datos.originadoNombre||"", datos.originadoNoEmp||"", datos.firmaOriginado||"", "", "", "", "", "", ""]);

    return { success: true, folio: folioStr, message: "Pase creado con éxito." };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function obtenerPasesPorEstatusRRHH(estatusBuscado) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PASE RRHH");
    if (!sheet || sheet.getLastRow() < 2) return { success: true, pases: [] };
    
    var data = sheet.getDataRange().getValues();
    var col = MapeoColumnasPase(data[0]);
    var ahora = Date.now();
    var unDiaMs = 24 * 60 * 60 * 1000;

    for (var i = 1; i < data.length; i++) {
      var estatusRow = String(data[i][col.estatus] || "").trim().toUpperCase();
      var tsCreacion = parseInt(data[i][col.tsCreacion]) || 0;
      if (tsCreacion > 0 && (estatusRow === "PENDIENTE GUARDIA" || estatusRow === "PENDIENTE ENTRADA") && (ahora - tsCreacion > unDiaMs)) {
        var row = i + 1;
        if (!String(data[i][col.horaEntrada] || "").trim() || String(data[i][col.horaEntrada] || "").trim() === "PENDIENTE") {
          sheet.getRange(row, col.horaEntrada + 1).setValue("NO REGRESÓ");
          data[i][col.horaEntrada] = "NO REGRESÓ";
        }
        sheet.getRange(row, col.estatus + 1).setValue("COMPLETADO");
        data[i][col.estatus] = "COMPLETADO";
      }
    }
    
    var pases = [];
    var estBuscadoClean = String(estatusBuscado || "").trim().toUpperCase();

    for (var i = 1; i < data.length; i++) {
      var est = String(data[i][col.estatus] || "").trim().toUpperCase();
      if ((est === estBuscadoClean) || (estBuscadoClean === "PENDIENTE GUARDIA" && (est === "PENDIENTE GUARDIA" || est === "PENDIENTE ENTRADA"))) {
        pases.push({
          folio: String(data[i][col.folio] || ""), estatus: est, fecha: String(data[i][col.fechaPase] || ""),
          trabajador: String(data[i][col.trabajador] || ""), noEmp: String(data[i][col.noEmp] || ""), asunto: String(data[i][col.asunto] || "")
        });
      }
    }
    return { success: true, pases: pases.reverse() };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function obtenerPasePorFolioRRHH(folioBuscado) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PASE RRHH");
    if (!sheet) return { success: false, message: "Pestaña 'PASE RRHH' no encontrada" };
    
    var data = sheet.getDataRange().getValues();
    var col = MapeoColumnasPase(data[0]);
    var folioClean = String(folioBuscado || "").trim().toUpperCase();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][col.folio] || "").trim().toUpperCase() === folioClean) {
        return {
          success: true,
          pase: {
            folio: String(data[i][col.folio] || ""), estatus: String(data[i][col.estatus] || ""),
            fechaRegistro: String(data[i][col.fechaReg] || ""), tipoPase: String(data[i][col.tipoPase] || ""),
            horaEntrada: String(data[i][col.horaEntrada] || ""), horaSalida: String(data[i][col.horaSalida] || ""),
            fechaPase: String(data[i][col.fechaPase] || ""), trabajadorNombre: String(data[i][col.trabajador] || ""),
            trabajadorNoEmp: String(data[i][col.noEmp] || ""), depto: String(data[i][col.depto] || ""),
            turno: String(data[i][col.turno] || ""), asunto: String(data[i][col.asunto] || ""),
            firmaTrabajador: String(data[i][col.firmaTrabajador] || ""), originadoNombre: String(data[i][col.origNombre] || ""),
            originadoNoEmp: String(data[i][col.origNoEmp] || ""), firmaOriginado: String(data[i][col.firmaOrig] || ""),
            supervisorNombre: String(data[i][col.supNombre] || ""), supervisorNoEmp: String(data[i][col.supNoEmp] || ""),
            firmaSupervisor: String(data[i][col.firmaSup] || ""), guardiaNombre: String(data[i][col.guardiaNombre] || ""),
            guardiaNoEmp: String(data[i][col.guardiaNoEmp] || ""), firmaGuardia: String(data[i][col.firmaGuardia] || "")
          }
        };
      }
    }
    return { success: false, message: "Folio no encontrado" };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function autorizarPaseSupervisorRRHH(datos) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PASE RRHH");
    var data = sheet.getDataRange().getValues();
    var col = MapeoColumnasPase(data[0]);

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][col.folio] || "").trim().toUpperCase() === String(datos.folio || "").trim().toUpperCase()) {
        var row = i + 1;
        sheet.getRange(row, col.estatus + 1).setValue("PENDIENTE GUARDIA");
        sheet.getRange(row, col.supNombre + 1).setValue(datos.supervisorNombre);
        sheet.getRange(row, col.supNoEmp + 1).setValue(datos.supervisorNoEmp);
        sheet.getRange(row, col.firmaSup + 1).setValue(datos.firmaSupervisor);
        return { success: true, message: "Pase " + datos.folio + " autorizado y enviado a Caseta." };
      }
    }
    return { success: false, message: "Folio no encontrado" };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function concluirPaseGuardiaRRHH(datos) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PASE RRHH");
    var data = sheet.getDataRange().getValues();
    var col = MapeoColumnasPase(data[0]);

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][col.folio] || "").trim().toUpperCase() === String(datos.folio || "").trim().toUpperCase()) {
        var row = i + 1;
        
        sheet.getRange(row, col.estatus + 1).setValue("COMPLETADO");
        sheet.getRange(row, col.horaEntrada + 1).setValue(datos.horaEntrada || data[i][col.horaEntrada]);
        sheet.getRange(row, col.horaSalida + 1).setValue(datos.horaSalida || data[i][col.horaSalida]);
        sheet.getRange(row, col.guardiaNombre + 1).setValue(datos.guardiaNombre);
        sheet.getRange(row, col.guardiaNoEmp + 1).setValue(datos.guardiaNoEmp);
        sheet.getRange(row, col.firmaGuardia + 1).setValue(datos.firmaGuardia);

        var rU = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
        var paseObj = {
          folio: String(rU[col.folio]||""), estatus: "COMPLETADO", fechaPase: String(rU[col.fechaPase]||""), tipoPase: String(rU[col.tipoPase]||""),
          horaEntrada: String(rU[col.horaEntrada]||"N/A"), horaSalida: String(rU[col.horaSalida]||"N/A"),
          trabajadorNombre: String(rU[col.trabajador]||""), trabajadorNoEmp: String(rU[col.noEmp]||""), depto: String(rU[col.depto]||""),
          turno: String(rU[col.turno]||""), asunto: String(rU[col.asunto]||""), firmaTrabajador: String(rU[col.firmaTrabajador]||""),
          originadoNombre: String(rU[col.origNombre]||""), originadoNoEmp: String(rU[col.origNoEmp]||""), firmaOriginado: String(rU[col.firmaOrig]||""),
          supervisorNombre: String(rU[col.supNombre]||""), supervisorNoEmp: String(rU[col.supNoEmp]||""), firmaSupervisor: String(rU[col.firmaSup]||""),
          guardiaNombre: String(rU[col.guardiaNombre]||""), guardiaNoEmp: String(rU[col.guardiaNoEmp]||""), firmaGuardia: String(rU[col.firmaGuardia]||"")
        };

        enviarCorreoEjecutivoRRHH(paseObj);
        return { success: true, message: "Pase COMPLETADO. Documento inmutable enviado por correo a Recursos Humanos." };
      }
    }
    return { success: false, message: "Folio no encontrado" };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function enviarCorreoEjecutivoRRHH(p) {
  var destinatario = "jose.h.ramirez@nidec-ga.com";
  var asunto = "✅ REPORTE PASE CONCLUIDO (FA-RH-029): " + p.folio + " - " + p.trabajadorNombre;

  var renderImg = function(b64) {
    return (b64 && b64.indexOf("data:image") !== -1) ? '<img src="' + b64 + '" style="max-height:55px; max-width:180px; object-fit:contain;">' : '<span style="color:#94a3b8; font-size:11px;">N/A</span>';
  };

  var body = `
    <div style="font-family:'Segoe UI',Arial,sans-serif; background:#f8fafc; padding:30px; color:#0f172a;">
      <div style="max-width:850px; margin:0 auto; background:#ffffff; border-radius:10px; border-top:8px solid #009B4A; box-shadow:0 10px 25px rgba(0,0,0,0.1); padding:30px;">
        <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px; border-bottom:3px solid #009B4A; padding-bottom:15px;">
          <tr>
            <td width="30%"><img src="https://www.nidec.com/files/user/www-nidec-com/corporate/about/slogan/img/img_slogan_01.png" style="height:45px;"></td>
            <td width="40%" align="center"><h2 style="margin:0; font-size:22px; color:#000;">PASE ENTRADA / SALIDA</h2></td>
            <td width="30%" align="right"><div style="color:#dc2626; font-size:22px; font-weight:900; font-family:monospace;">${p.folio}</div><div style="font-size:12px; font-weight:bold; color:#475569;">MOTORES REYNOSA</div></td>
          </tr>
        </table>
        
        <div style="background:#e0f2fe; border:1px solid #bae6fd; padding:12px; border-radius:6px; text-align:center; font-weight:900; margin-bottom:20px; color:#0369a1; text-transform:uppercase;">
          REPORTE CORPORATIVO DE AUDITORÍA - ESTATUS: ${p.estatus}
        </div>

        <table width="100%" cellpadding="10" cellspacing="0" border="1" style="border-collapse:collapse; border:2px solid #000; font-size:13px; text-transform:uppercase;">
          <tr style="background:#f1f5f9;"><td colspan="6" style="font-weight:900;">TIPO DE PASE: <span style="color:#2563eb;">${p.tipoPase}</span></td></tr>
          <tr><td width="16%" style="background:#f8fafc; font-weight:bold;">HORA ENTRADA:</td><td width="17%"><b>${p.horaEntrada}</b></td><td width="16%" style="background:#f8fafc; font-weight:bold;">HORA SALIDA:</td><td width="17%"><b>${p.horaSalida}</b></td><td width="10%" style="background:#f8fafc; font-weight:bold;">FECHA:</td><td width="24%"><b>${p.fechaPase}</b></td></tr>
          <tr><td style="background:#f8fafc; font-weight:bold;">TRABAJADOR:</td><td colspan="3"><b>${p.trabajadorNombre}</b></td><td style="background:#f8fafc; font-weight:bold;">NO. EMP:</td><td><b>${p.trabajadorNoEmp}</b></td></tr>
          <tr><td style="background:#f8fafc; font-weight:bold;">DEPTO:</td><td colspan="3"><b>${p.depto}</b></td><td style="background:#f8fafc; font-weight:bold;">TURNO:</td><td><b>${p.turno}</b></td></tr>
          <tr><td style="background:#f8fafc; font-weight:bold;">ASUNTO:</td><td colspan="3"><b>${p.asunto}</b></td><td style="background:#f8fafc; font-weight:bold;">FIRMA TRAB:</td><td align="center">${renderImg(p.firmaTrabajador)}</td></tr>
          <tr><td style="background:#f8fafc; font-weight:bold;">ORIGINADO POR:</td><td colspan="3"><b>${p.originadoNombre} (#${p.originadoNoEmp})</b></td><td style="background:#f8fafc; font-weight:bold;">FIRMA ORIG:</td><td align="center">${renderImg(p.firmaOriginado)}</td></tr>
          <tr><td style="background:#fffbeb; font-weight:bold; color:#b45309;">SUPERVISOR:</td><td colspan="3" style="background:#fffbeb;"><b>${p.supervisorNombre} (#${p.supervisorNoEmp})</b></td><td style="background:#fffbeb; font-weight:bold; color:#b45309;">FIRMA SUP:</td><td align="center" style="background:#fffbeb;">${renderImg(p.firmaSupervisor)}</td></tr>
          <tr><td style="background:#eff6ff; font-weight:bold; color:#1e40af;">GUARDIA:</td><td colspan="3" style="background:#eff6ff;"><b>${p.guardiaNombre} (#${p.guardiaNoEmp})</b></td><td style="background:#eff6ff; font-weight:bold; color:#1e40af;">FIRMA GRD:</td><td align="center" style="background:#eff6ff;">${renderImg(p.firmaGuardia)}</td></tr>
        </table>
        
        <p style="text-align:center; font-size:11px; color:#64748b; margin-top:30px; font-weight:bold;">
          Documento Electrónico Inmutable - WorkPulse ERP System<br>Generado Automáticamente en Cierre de Caseta (FA-RH-029)
        </p>
      </div>
    </div>
  `;
  MailApp.sendEmail({ to: destinatario, subject: asunto, htmlBody: body });
}

function obtenerTodosLosPasesRRHH(pass) {
  if (String(pass || "").trim() !== "ADMINMB2026") { return { success: false, message: "🔒 Contraseña de Acceso Incorrecta." }; }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PASE RRHH");
    if (!sheet || sheet.getLastRow() < 2) return { success: true, pases: [] };
    
    var data = sheet.getDataRange().getValues();
    var col = MapeoColumnasPase(data[0]);
    var pases = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][col.folio] || "").trim() !== "") {
        pases.push({
          folio: String(data[i][col.folio] || ""), estatus: String(data[i][col.estatus] || ""),
          fecha: String(data[i][col.fechaPase] || ""), trabajador: String(data[i][col.trabajador] || ""),
          noEmp: String(data[i][col.noEmp] || ""), asunto: String(data[i][col.asunto] || "")
        });
      }
    }
    return { success: true, pases: pases.reverse() };
  } catch (e) { return { success: false, error: e.toString() }; }
}
function obtenerUrlSupervisorPerformance() {
  // Retorna la URL del Script que carga la vista de Performance de Supervisores
  return ScriptApp.getService().getUrl() + "?p=supervisorPerformance";
}



// =========================================================================
// MÓDULO TOP 10 MÁQUINAS - PROCESAMIENTO BACKEND PARETO 80/20
// =========================================================================
function getTop10MaquinasData() {
  try {
    let ss = getDB();
    let sh = ss.getSheetByName(OPEN_SHEET_NAME);
    if (!sh) return { success: false, errorServer: "Pestaña 'MTTOs Abiertos' no encontrada" };

    let data = sh.getLastRow() > 1 ? sh.getDataRange().getValues() : [];
    if (data.length <= 1) return { success: true, top10: [], categories: [], countsData: [], cumulativeData: [], totalReportes: 0 };

    let maquinaCounts = {};
    let totalTickets = 0;

    for (let i = 1; i < data.length; i++) {
      let folio = String(data[i][0]).trim();
      if (!folio || folio === "N/A") continue;

      let dondeStr = String(data[i][4] || "");
      
      // NOS ASEGURAMOS DE USAR LA FUNCIÓN PARSE DONDE
      let dInfo = parseDonde(dondeStr); 
      
      let maqName = dInfo.maquina && dInfo.maquina !== "N/A" ? dInfo.maquina : "General";
      let key = maqName + "||" + (dInfo.idMaquina || "N/A") + "||" + (dInfo.linea || "N/A") + "||" + (dInfo.area || "N/A");
      
      maquinaCounts[key] = (maquinaCounts[key] || 0) + 1;
      totalTickets++;
    }

    let maqArray = [];
    for (let k in maquinaCounts) {
      let p = k.split("||");
      maqArray.push({
        maquina: p[0],
        idMaquina: p[1],
        linea: p[2],
        area: p[3],
        count: maquinaCounts[k]
      });
    }
    
    // ORDEN DE MAYOR A MENOR (PARETO)
    maqArray.sort((a, b) => b.count - a.count);

    let top10 = maqArray.slice(0, 10);

    let categories = [];
    let countsData = [];
    let cumulativeData = [];
    let acumulado = 0;

    for (let j = 0; j < top10.length; j++) {
      categories.push(top10[j].maquina);
      countsData.push(top10[j].count);
      
      acumulado += top10[j].count;
      let pctAcumulado = totalTickets > 0 ? parseFloat(((acumulado / totalTickets) * 100).toFixed(1)) : 0;
      cumulativeData.push(pctAcumulado);
    }

    return { 
      success: true, 
      top10: top10, 
      categories: categories, 
      countsData: countsData, 
      cumulativeData: cumulativeData, 
      totalReportes: totalTickets 
    };

  } catch (e) {
    return { success: false, errorServer: e.toString() };
  }
}

// =========================================================================
// FUNCIÓN AUXILIAR BACKEND PARA PARSEAR COLUMNA "DONDE"
// =========================================================================
function parseDonde(str) { 
  if(!str) return {area:"N/A", linea:"N/A", maquina:"N/A", idMaquina:"N/A", turno:"N/A"}; 
  if(str.indexOf("||") !== -1) { 
      let p = str.split("||").map(s => s.trim()); 
      return {area: p[0]||"N/A", linea: p[1]||"N/A", maquina: p[2]||"N/A", idMaquina: p[3]||"N/A", turno: p[4]||"N/A"}; 
  } else if(str.indexOf("/") !== -1) { 
      let p = str.split("/").map(s => s.trim()); 
      return {area: p[0]||"N/A", linea: p[1]||"N/A", maquina: p[2]||"N/A", idMaquina: "N/A", turno: p[3]||"N/A"}; 
  } 
  return {area: str, linea: "N/A", maquina: "N/A", idMaquina:"N/A", turno:"N/A"}; 
}
// =========================================================================
// OBTENER TODOS LOS LOGS (USADO POR MÚLTIPLES MÓDULOS)
// =========================================================================
function getAllLogs() {
  try {
    const ss = getDB();
    
    // Función auxiliar para obtener datos de forma segura
    const getSheetData = (exactName) => { 
      try { 
        let sheet = ss.getSheetByName(exactName); 
        if (!sheet) return []; 
        return sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues(); 
      } catch(err) { 
        return []; 
      } 
    };
    
    return { 
      success: true, 
      abiertos: getSheetData(OPEN_SHEET_NAME), 
      cerrados: getSheetData(LOG_SHEET_NAME), 
      ehs: getSheetData(EHS_SHEET_NAME), 
      refacciones: getSheetData(REF_SHEET_NAME), 
      permisos: getSheetData("Log_Permisos_EHS") 
    };
  } catch(e) { 
    return { success: false, errorServer: e.toString() }; 
  }
}
// =========================================================================
// BACKEND WCM: CONTROLADOR UNIFICADO DE PRODUCCIÓN Y KPI (CORREGIDO)
// =========================================================================
function obtenerResumenEjecutivoDirecto() {
  try {
    var ss = getDB();
    var hoy = new Date();
    var fechaHoyNorm = normalizarFechaJS(hoy);

    // 1. OBTENER LÍNEAS ACTIVAS DEL ÁREA Y TURNO
    var shConfig = ss.getSheetByName(CONFIG_PROD_SHEET);
    var lineasConfig = {};
    var lineasActivasList = [];

    if (shConfig && shConfig.getLastRow() > 1) {
      var configData = shConfig.getRange(2, 1, shConfig.getLastRow() - 1, Math.min(14, shConfig.getLastColumn())).getValues();
      configData.forEach(function(r) {
        var linea = String(r[0] || "").trim();
        var areaVal = String(r[1] || "DRYERS").trim().toUpperCase();
        var supVal = String(r[2] || "").trim().toUpperCase();
        var turnoVal = String(r[3] || "").trim().toUpperCase();

        var tieneSupervisor = (supVal !== "" && supVal !== "ASIGNADO POR ADMIN" && supVal !== "SIN ASIGNAR" && supVal !== "VACIO");
        var esTurnoD = (turnoVal === "TURNO D" || turnoVal === "D");
        var esDryers = (areaVal === "DRYERS");

        if (linea !== "") {
          lineasConfig[linea] = {
            area: areaVal,
            supervisor: String(r[2] || "Sin Asignar").trim(),
            turno: String(r[3] || "Turno D").trim(),
            activa: (tieneSupervisor && esTurnoD && esDryers)
          };

          if (tieneSupervisor && esTurnoD && esDryers) {
            lineasActivasList.push(linea);
          }
        }
      });
    }

    if (lineasActivasList.length === 0) {
      lineasActivasList = ["1", "2", "7", "14", "16"];
    }

    // 2. EXTRAER VALORES OFICIALES DESDE KPIS DE CUMPLIMIENTO Y DOWN TIME
    var resKpis = obtenerKpisProduccionData();
    var perfData = (resKpis && resKpis.success && resKpis.perf) ? resKpis.perf : {};
    var dtData = (resKpis && resKpis.success && resKpis.dt) ? resKpis.dt : {};

    var valPerformancePct = Number(perfData.pct) || 0;
    var valProdPlaneada = Number(perfData.plan) || 0;
    var valProdReal = Number(perfData.real) || 0;
    var valDowntimeMins = Number(dtData.dtMinutos) || 0;

    // 3. CONSOLIDAR LÍNEAS DE PRODUCCIÓN Y SUMATORIA DEL PIZARRÓN DESDUPLICADAS
    var shPizarron = ss.getSheetByName("Log_Pizarron_Produccion");
    var pizarronMap = {};
    var sumatorioDeltaPzs = 0;
    var sumatorioPallets = 0;

    if (shPizarron && shPizarron.getLastRow() > 1) {
      var pizData = shPizarron.getRange(2, 1, shPizarron.getLastRow() - 1, Math.min(15, shPizarron.getLastColumn())).getDisplayValues();
      var mapaPizUnico = {};

      pizData.forEach(function(r) {
        var lin = String(r[0] || "").trim();
        var fecPizNorm = normalizarFechaJS(r[1]);
        var hIdx = parseInt(r[2], 10);
        var metaVal = parseInt(r[3], 10) || 0;
        var realVal = parseInt(r[4], 10);

        if (lin !== "" && !isNaN(hIdx) && (fecPizNorm === fechaHoyNorm || fecPizNorm === "")) {
          var kPiz = lin + "_" + hIdx;
          mapaPizUnico[kPiz] = {
            linea: lin,
            meta: metaVal,
            real: (!isNaN(realVal) && realVal >= 0) ? realVal : 0
          };
        }
      });

      Object.keys(mapaPizUnico).forEach(function(kPiz) {
        var itemP = mapaPizUnico[kPiz];
        if (!pizarronMap[itemP.linea]) pizarronMap[itemP.linea] = { realAcum: 0, metaAcum: 0 };
        pizarronMap[itemP.linea].metaAcum += itemP.meta;
        pizarronMap[itemP.linea].realAcum += itemP.real;
      });
    }

    // 4. AGRUPAR TIEMPOS MUERTOS 6M + O DESDUPLICADOS
    var shProd = getOrCreateProdSheet();
    var lastRow = shProd.getLastRow();
    var lineasDataMap = {};
    var horasLabel = ["07:00-08:00", "08:00-09:00", "09:00-10:00", "10:00-11:00", "11:00-12:00", "12:00-13:00", "13:00-14:00", "14:00-15:00", "15:00-16:00"];

    lineasActivasList.forEach(function(l) {
      var cfg = lineasConfig[l] || { area: "DRYERS", supervisor: "Sin Asignar", turno: "Turno D" };
      var pizInfo = pizarronMap[l] || { realAcum: 0, metaAcum: 0 };

      lineasDataMap[l] = {
        linea: l,
        area: cfg.area,
        supervisor: cfg.supervisor,
        turno: cfg.turno,
        hombre: 0, maquina: 0, material: 0, metodo: 0, calidad: 0, produccion: 0, otros: 0, totalDT: 0,
        metaTotal: pizInfo.metaAcum,
        realTotal: pizInfo.realAcum,
        detalles: []
      };
    });

    if (lastRow > 1) {
      var maxCols = Math.min(shProd.getLastColumn(), 15);
      var prodData = shProd.getRange(2, 1, lastRow - 1, maxCols).getDisplayValues();
      var mapaProdUnico = {};

      prodData.forEach(function(r) {
        var lStr = String(r[0] || "").trim();
        var fecFilaNorm = normalizarFechaJS(r[1]);
        var hIdx = parseInt(r[2], 10);

        var coincideFecha = (fecFilaNorm === fechaHoyNorm || fecFilaNorm === "");

        if (lStr !== "" && !isNaN(hIdx) && lineasDataMap[lStr] && coincideFecha) {
          var kProd = lStr + "_" + hIdx;
          mapaProdUnico[kProd] = {
            linea: lStr,
            fecha: r[1] || "",
            horaIdx: hIdx,
            hombre: parseInt(r[5], 10) || 0,
            maquina: parseInt(r[6], 10) || 0,
            material: parseInt(r[7], 10) || 0,
            metodo: parseInt(r[8], 10) || 0,
            calidad: parseInt(r[9], 10) || 0,
            produccion: parseInt(r[10], 10) || 0,
            otros: parseInt(r[11], 10) || 0,
            comentario: String(r[12] || "").trim(),
            falla: (r.length >= 15) ? String(r[14] || "").trim() : ""
          };
        }
      });

      Object.keys(mapaProdUnico).forEach(function(kP) {
        var item = mapaProdUnico[kP];
        var lStr = item.linea;

        var dtFila = item.hombre + item.maquina + item.material + item.metodo + item.calidad + item.produccion + item.otros;

        lineasDataMap[lStr].hombre += item.hombre;
        lineasDataMap[lStr].maquina += item.maquina;
        lineasDataMap[lStr].material += item.material;
        lineasDataMap[lStr].metodo += item.metodo;
        lineasDataMap[lStr].calidad += item.calidad;
        lineasDataMap[lStr].produccion += item.produccion;
        lineasDataMap[lStr].otros += item.otros;
        lineasDataMap[lStr].totalDT += dtFila;

        var com = item.comentario;
        var fal = item.falla;

        if (dtFila > 0 || (com !== "" && com !== "-") || (fal !== "" && fal !== "-")) {
          lineasDataMap[lStr].detalles.push({
            hora: horasLabel[item.horaIdx] || ("Hora " + (item.horaIdx + 1)),
            fecha: item.fecha,
            hombre: item.hombre, maquina: item.maquina, material: item.material, metodo: item.metodo, calidad: item.calidad, produccion: item.produccion, otros: item.otros,
            dtTotal: dtFila,
            comentario: (com && com !== "-") ? com : "-",
            falla: (fal && fal !== "-") ? fal : "-"
          });
        }
      });
    }

    var rankingLineas = [];
    for (var kKey in lineasDataMap) {
      if (Object.prototype.hasOwnProperty.call(lineasDataMap, kKey)) {
        rankingLineas.push(lineasDataMap[kKey]);
      }
    }
    rankingLineas.sort(function(a, b) { return b.totalDT - a.totalDT; });

    if (valProdPlaneada === 0) {
      rankingLineas.forEach(function(item) {
        valProdPlaneada += item.metaTotal;
        valProdReal += item.realTotal;
      });
    }

    sumatorioDeltaPzs = valProdReal - valProdPlaneada;
    sumatorioPallets = (valProdReal / 198);

    var resSup = typeof obtenerDatosSupervisorPerformance === "function" ? obtenerDatosSupervisorPerformance() : null;
    var rankingSupervisores = (resSup && resSup.success && resSup.ranking) ? resSup.ranking : [];

    return {
      success: true,
      lineasEvaluadasCount: lineasActivasList.length,
      downtimeMins: valDowntimeMins,
      performancePct: valPerformancePct,
      prodPlaneada: valProdPlaneada,
      prodReal: valProdReal,
      deltaPzs: sumatorioDeltaPzs,
      palletsReal: sumatorioPallets,
      deltaPallets: (sumatorioDeltaPzs / 198),
      ranking: rankingLineas,
      rankingSupervisores: rankingSupervisores,
      fechaFiltroLabel: "HOY"
    };
  } catch (e) {
    return { success: false, error: "Error en servidor: " + e.toString() };
  }
}
// =========================================================================
// BACKEND: CALCULADORA DE GENERAL PERFORMANCE POR LÍNEA (CODE.GS)
// =========================================================================
function obtenerDatosGeneralPerformance() {
  try {
    let ss = getDB();
    let hoy = new Date();
    let fechaHoyNorm = normalizarFechaSegura(hoy);

    // 1. OBTENER CONFIGURACIÓN DE LÍNEAS ACTIVAS (CUALQUIER ÁREA Y TURNO HABILITADO)
    let shConfig = ss.getSheetByName("Config_Lineas_Produccion") || 
                   ss.getSheetByName("Configuracion_Lineas_Admin") || 
                   ss.getSheetByName("Log_Configuracion_Lineas");
                   
    if (!shConfig || shConfig.getLastRow() <= 1) return { success: true, ranking: [] };
    let configData = shConfig.getDataRange().getDisplayValues();

    let lineasConfig = {};
    for (let c = 1; c < configData.length; c++) {
      let r = configData[c];
      let numL = parseInt(String(r[0] || '').replace(/[^0-9]/g, ''), 10);
      if (isNaN(numL)) continue;

      let area = String(r[1] || "DRYERS").trim().toUpperCase();
      let supervisor = String(r[2] || "").trim().toUpperCase();
      let turno = String(r[3] || "").trim().toUpperCase();

      let invalidSup = ['', 'ASIGNADO POR ADMIN', 'SIN ASIGNAR', 'VACIO', 'VACÍO', 'N/A', 'N/D', '-', 'NINGUNO', 'SIN SUPERVISOR'];
      let invalidTurno = ['', 'VACIO', 'VACÍO', 'SIN ASIGNAR', 'OFF', 'DESACTIVADO', 'INACTIVO', 'N/A', '-', 'NO'];

      let tieneSupervisor = invalidSup.indexOf(supervisor) === -1;
      let tieneTurno = invalidTurno.indexOf(turno) === -1;

      if (tieneSupervisor && tieneTurno) {
        let metas = [];
        for (let m = 0; m < 9; m++) {
          metas.push(parseInt(r[4 + m], 10) || 300);
        }
        lineasConfig[numL] = { area: area, supervisor: supervisor, metas: metas };
      }
    }

    // 2. OBTENER REGISTROS DE PRODUCCIÓN DE HOY (BÚSQUEDA MULTI-PESTAÑA)
    let shProd = ss.getSheetByName("Log_Pizarron_Produccion") || ss.getSheetByName("Log_Produccion_Pizarron");
    let prodPorLineaHora = {};
    let logMetasHora = {};

    if (shProd && shProd.getLastRow() > 1) {
      let prodData = shProd.getDataRange().getDisplayValues();
      for (let i = 1; i < prodData.length; i++) {
        let r = prodData[i];
        let lNum = parseInt(String(r[0] || '').replace(/[^0-9]/g, ''), 10);
        let fNorm = normalizarFechaSegura(r[1]);
        let hIdx = parseInt(r[2], 10);
        let meta = parseInt(r[3], 10) || 0;
        let real = parseInt(r[4], 10) || 0;

        if (!isNaN(lNum) && !isNaN(hIdx) && fNorm === fechaHoyNorm) {
          if (!prodPorLineaHora[lNum]) prodPorLineaHora[lNum] = {};
          if (!logMetasHora[lNum]) logMetasHora[lNum] = {};

          prodPorLineaHora[lNum][hIdx] = real;
          if (meta > 0) logMetasHora[lNum][hIdx] = meta;
        }
      }
    }

    // 3. CALCULAR PERFORMANCE INDIVIDUAL POR LÍNEA
    let ranking = [];
    Object.keys(lineasConfig).forEach(function(lStr) {
      let numL = parseInt(lStr, 10);
      let cfg = lineasConfig[numL];
      let perfHoras = [];
      let sumaPerfValidas = 0;
      let horasContadas = 0;

      for (let h = 0; h < 9; h++) {
        let metaH = (logMetasHora[numL] && logMetasHora[numL][h] > 0) ? logMetasHora[numL][h] : (cfg.metas[h] || 300);
        let realH = (prodPorLineaHora[numL] && prodPorLineaHora[numL][h] !== undefined) ? prodPorLineaHora[numL][h] : 0;
        let pctH = metaH > 0 ? Math.round((realH / metaH) * 100) : 0;

        perfHoras.push(pctH);

        if (pctH > 0) {
          sumaPerfValidas += pctH;
          horasContadas++;
        }
      }

      let perfGlobal = horasContadas > 0 ? Math.round(sumaPerfValidas / horasContadas) : 0;

      ranking.push({
        linea: numL,
        area: cfg.area,
        supervisor: cfg.supervisor,
        perfGlobal: perfGlobal,
        perfHoras: perfHoras
      });
    });

    // 4. ORDENAR DE MAYOR A MENOR RENDIMIENTO GLOBAL
    ranking.sort(function(a, b) { return b.perfGlobal - a.perfGlobal; });

    return { success: true, ranking: ranking, fecha: fechaHoyNorm };

  } catch (err) {
    return { success: false, message: err.toString(), ranking: [] };
  }
}
// =========================================================================
// MÓDULO CENTRO DE COMANDO (8 KPIs) - BÚSQUEDA MULTI-PESTAÑA DE METAS Y REAL
// =========================================================================

function normalizarFechaSegura(val) {
  if (!val) return '';
  if (val instanceof Date) {
    let d = String(val.getDate()).padStart(2, '0');
    let m = String(val.getMonth() + 1).padStart(2, '0');
    let a = val.getFullYear();
    return d + '/' + m + '/' + a;
  }
  let str = String(val).trim().split(' ')[0];
  let p = str.split(/[\/\-]/);
  if (p.length === 3) {
    let n1 = parseInt(p[0], 10), n2 = parseInt(p[1], 10), n3 = parseInt(p[2], 10);
    if (isNaN(n1) || isNaN(n2) || isNaN(n3)) return '';
    if (n1 > 31) {
      return String(n3).padStart(2, '0') + '/' + String(n2).padStart(2, '0') + '/' + String(n1);
    }
    return String(n1).padStart(2, '0') + '/' + String(n2).padStart(2, '0') + '/' + String(n3);
  }
  return str;
}

function parseFechaSegura(fecStr) {
  if (!fecStr) return null;
  if (fecStr instanceof Date) return fecStr;
  let norm = normalizarFechaSegura(fecStr);
  let p = norm.split('/');
  if (p.length === 3) {
    let d = parseInt(p[0], 10);
    let m = parseInt(p[1], 10) - 1;
    let a = parseInt(p[2], 10);
    if (!isNaN(d) && !isNaN(m) && !isNaN(a)) {
      return new Date(a, m, d);
    }
  }
  return null;
}

function obtenerDatosCentroComando() {
  try {
    let ss = getDB();
    let hoy = new Date();
    let hoyNorm = normalizarFechaSegura(hoy);
    let hoyYear = hoy.getFullYear();
    let hoyMonth = hoy.getMonth();

    // 🎯 1. BÚSQUEDA DE PESTAÑA DE CONFIGURACIÓN (ACEPTA Config_Lineas_Produccion)
    let lineasActivasSet = {};
    let configMetasPorLinea = {};

    let shConfig = ss.getSheetByName("Config_Lineas_Produccion") || 
                   ss.getSheetByName("Configuracion_Lineas_Admin") || 
                   ss.getSheetByName("Log_Configuracion_Lineas");

    if (shConfig && shConfig.getLastRow() > 1) {
      let cfgData = shConfig.getDataRange().getDisplayValues();
      for (let c = 1; c < cfgData.length; c++) {
        let numL = parseInt(String(cfgData[c][0] || '').replace(/[^0-9]/g, ''), 10);
        if (isNaN(numL)) continue;

        let supVal = String(cfgData[c][2] || '').toUpperCase().trim();
        let turnoVal = String(cfgData[c][3] || '').toUpperCase().trim();

        let invalidSup = ['', 'ASIGNADO POR ADMIN', 'SIN ASIGNAR', 'VACIO', 'VACÍO', 'N/A', 'N/D', '-', 'NINGUNO', 'SIN SUPERVISOR'];
        let invalidTurno = ['', 'VACIO', 'VACÍO', 'SIN ASIGNAR', 'OFF', 'DESACTIVADO', 'INACTIVO', 'N/A', '-', 'NO'];

        let tieneSupervisor = invalidSup.indexOf(supVal) === -1;
        let tieneTurno = invalidTurno.indexOf(turnoVal) === -1;

        if (tieneSupervisor && tieneTurno) {
          lineasActivasSet[numL] = true;
          let metasArr = [];
          for (let m = 0; m < 24; m++) {
            metasArr.push(parseInt(cfgData[c][4 + m], 10) || 0);
          }
          configMetasPorLinea[numL] = metasArr;
        }
      }
    }

    // 🎯 2. EXTRAER REGISTROS Y METAS DIRECTAS DEL LOG DE PRODUCCIÓN DE HOY
    let shProd = ss.getSheetByName("Log_Pizarron_Produccion") || ss.getSheetByName("Log_Produccion_Pizarron");
    let rawLogData = [];
    let logMetasHoyMap = {};

    if (shProd && shProd.getLastRow() > 1) {
      rawLogData = shProd.getDataRange().getDisplayValues();
      for (let i = 1; i < rawLogData.length; i++) {
        let nL = parseInt(String(rawLogData[i][0] || '').replace(/[^0-9]/g, ''), 10);
        let fStr = String(rawLogData[i][1] || '').trim();
        let hIdx = parseInt(rawLogData[i][2], 10);
        let metaVal = parseInt(rawLogData[i][3], 10) || 0;

        if (!isNaN(nL) && !isNaN(hIdx) && normalizarFechaSegura(fStr) === hoyNorm) {
          lineasActivasSet[nL] = true; // Marca activa la línea si registró actividad hoy
          if (metaVal > 0) {
            logMetasHoyMap[nL + "_" + hIdx] = metaVal;
          }
        }
      }
    }

    // Respaldo de seguridad en caso de tablas sin inicializar
    if (Object.keys(lineasActivasSet).length === 0) {
      [1, 2, 7, 14, 16, 30].forEach(function(l) { lineasActivasSet[l] = true; });
    }

    // 🎯 3. SUMATORIA DINÁMICA DE METAS HORARIAS DE TODAS LAS LÍNEAS ACTIVAS (07:00 A 06:00)
    let metas24hExactas = new Array(24).fill(0);

    Object.keys(lineasActivasSet).forEach(function(lStr) {
      let numL = parseInt(lStr, 10);
      let arrM = configMetasPorLinea[numL] || [];

      for (let h = 0; h < 24; h++) {
        let claveLog = numL + "_" + h;
        // Prioriza la meta grabada en el log; si no existe, toma la meta configurada
        let metaHoraLinea = (logMetasHoyMap.hasOwnProperty(claveLog) && logMetasHoyMap[claveLog] > 0)
          ? logMetasHoyMap[claveLog]
          : (arrM[h] || 0);

        metas24hExactas[h] += metaHoraLinea;
      }
    });

    let metaSumaTotal = metas24hExactas.reduce(function(a, b) { return a + b; }, 0);

    let res = {
      dia: { real: new Array(24).fill(0), meta: metas24hExactas.slice() },
      semana: { real: new Array(7).fill(0), meta: new Array(7).fill(0) }, 
      mes: { real: new Array(9).fill(0), meta: new Array(9).fill(0) },     
      ano: { real: new Array(12).fill(0), meta: new Array(12).fill(0) },   
      turno: { real: 0, meta: metaSumaTotal, perfPct: 0, dt6M: [0,0,0,0,0,0,0], criticas: [] },
      fechaTitulo: ""
    };

    let diaSemana = hoy.getDay(); 
    let offsetLunes = diaSemana === 0 ? 6 : diaSemana - 1; 
    let lunesSemana = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - offsetLunes);
    lunesSemana.setHours(0,0,0,0);
    let domingoSemana = new Date(lunesSemana.getTime() + 6 * 86400000);
    domingoSemana.setHours(23,59,59,999);

    // 🎯 4. ACUMULACIÓN DE PRODUCCIÓN REAL
    if (rawLogData.length > 1) {
      let ultimasCapturasMap = {}; 

      for (let i = 1; i < rawLogData.length; i++) {
        let r = rawLogData[i];
        let numLinea = parseInt(String(r[0] || '').replace(/[^0-9]/g, ''), 10);
        let fechaStr = String(r[1] || "").trim();
        let hIdx = parseInt(r[2], 10);
        
        if (isNaN(numLinea) || isNaN(hIdx)) continue;
        if (!lineasActivasSet[numLinea]) continue;

        let fNorm = normalizarFechaSegura(fechaStr);
        let fObj = parseFechaSegura(fechaStr); 
        if (!fObj) continue; 

        let claveUnica = numLinea + "_" + hIdx + "_" + fNorm;
        let realVal = Math.max(0, parseInt(r[4], 10) || 0);

        ultimasCapturasMap[claveUnica] = {
          numLinea: numLinea,
          hIdx: hIdx,
          fNorm: fNorm,
          fObj: fObj,
          realVal: realVal
        };
      }

      Object.keys(ultimasCapturasMap).forEach(function(key) {
        let item = ultimasCapturasMap[key];
        let fObj = item.fObj;
        let fNorm = item.fNorm;
        let hIdx = item.hIdx;
        let realVal = item.realVal;

        if (fObj.getFullYear() === hoyYear && fObj.getMonth() <= hoyMonth) {
          let mIdx = fObj.getMonth();
          res.ano.real[mIdx] += realVal;

          let primerDiaAno = new Date(fObj.getFullYear(), 0, 1);
          let dAno = Math.floor((fObj - primerDiaAno) / (24 * 60 * 60 * 1000));
          let numSemanaAno = Math.ceil((fObj.getDay() + 1 + dAno) / 7);

          let semIdx = numSemanaAno - 31;
          if (semIdx >= 0 && semIdx < 9) {
            res.mes.real[semIdx] += realVal;
          }
        }

        if (fObj.getTime() >= lunesSemana.getTime() && fObj.getTime() <= domingoSemana.getTime()) {
          let dJS = fObj.getDay(); 
          let dIdx = dJS === 0 ? 6 : dJS - 1;
          if (dIdx >= 0 && dIdx < 7) {
            res.semana.real[dIdx] += realVal;
          }
        }

        if (fNorm === hoyNorm) {
          let posIdx = (hIdx >= 0 && hIdx < 24) ? hIdx : -1;
          if (posIdx >= 0 && posIdx < 24) {
            res.dia.real[posIdx] += realVal;
          }
        }
      });
    }

    // 🎯 5. ASIGNAR META TOTAL DIARIA A LOS DÍAS LABORADOS EN LA GRÁFICA SEMANAL
    for (let d = 0; d < 7; d++) {
      if (res.semana.real[d] > 0 || d === (hoy.getDay() === 0 ? 6 : hoy.getDay() - 1)) {
        res.semana.meta[d] = metaSumaTotal;
      }
    }

    // 🎯 6. RESUMEN EJECUTIVO (KPIs 5 A 8)
    try {
      if (typeof obtenerResumenEjecutivoDirecto === "function") {
        let resEjecutivo = obtenerResumenEjecutivoDirecto();
        if (resEjecutivo && resEjecutivo.success) {
          res.turno.perfPct = parseFloat(Number(resEjecutivo.performancePct).toFixed(1));
          res.turno.real = resEjecutivo.prodReal;
          res.turno.meta = resEjecutivo.prodPlaneada;

          let ranking = resEjecutivo.ranking || [];
          let totH = 0, totMaq = 0, totMat = 0, totMet = 0, totCal = 0, totProd = 0, totOtr = 0;
          let tiempoMuertoTotalGeneral = 0;
          let listaLineasConParo = [];

          ranking.forEach(function(r) {
            totH += Number(r.hombre) || 0;
            totMaq += Number(r.maquina) || 0;
            totMat += Number(r.material) || 0;
            totMet += Number(r.metodo) || 0;
            totCal += Number(r.calidad) || 0;
            totProd += Number(r.produccion) || 0;
            totOtr += Number(r.otros) || 0;

            let lDT = Number(r.totalDT) || 0;
            tiempoMuertoTotalGeneral += lDT;

            let at = "", fa = "";
            if (r.detalles && r.detalles.length) {
              for (let j = 0; j < r.detalles.length; j++) {
                let d = r.detalles[j];
                let comVal = String(d.comentario || '').trim();
                let falVal = String(d.falla || '').trim();
                
                if (comVal && comVal !== '-' && comVal !== 'OK' && comVal !== 'N A' && !at) at = comVal;
                if (falVal && falVal !== '-' && falVal !== 'OK' && falVal !== 'N A' && !fa) fa = falVal;
              }
            }

            let textoDetalle = (at && fa && at !== fa) ? at + " - " + fa : (fa || at || "SIN FALLA REGISTRADA");

            if (lDT > 0) {
              listaLineasConParo.push({
                linea: r.linea,
                dtMins: lDT,
                falla: textoDetalle
              });
            }
          });

          res.turno.dt6M = [totH, totMaq, totMat, totMet, totCal, totProd, totOtr];

          listaLineasConParo.forEach(function(item) {
            item.perf = tiempoMuertoTotalGeneral > 0 
              ? parseFloat(((item.dtMins / tiempoMuertoTotalGeneral) * 100).toFixed(1)) 
              : 0;
          });

          listaLineasConParo.sort(function(a, b) { return b.dtMins - a.dtMins; });
          res.turno.criticas = listaLineasConParo;
        }
      }
    } catch(errExec) {
      Logger.log("Error en Resumen Ejecutivo: " + errExec.toString());
    }

    let fManana = new Date(hoy.getTime() + 86400000);
    res.fechaTitulo = hoy.getDate().toString().padStart(2,'0') + "/" + (hoy.getMonth()+1).toString().padStart(2,'0') + "/" + hoy.getFullYear() + " - " + 
                      fManana.getDate().toString().padStart(2,'0') + "/" + (fManana.getMonth()+1).toString().padStart(2,'0') + "/" + fManana.getFullYear();

    return { success: true, data: res };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}