/**
 * JoieM — backend de Google Apps Script.
 *
 * A diferencia de los otros clientes, acá NO hay un Excel previo: el tablero
 * se arma desde el diagnóstico y el Sheet se crea solo y se va llenando con
 * lo que se carga en la web.
 *
 * CÓMO SE USA LA PRIMERA VEZ:
 * 1. Pegar este archivo en un proyecto de Apps Script de hola@soyevelyna.com.
 * 2. Ejecutar crearSheet() una vez: crea el Google Sheet "Evelyna I
 *    CONSULTORÍA: JoieM" dentro de la carpeta de Drive del cliente, con las
 *    hojas "01 I Plan de trabajo" y "02 I Proceso de trabajo" ya armadas
 *    (encabezados, desplegables y colores), y deja su id en el registro.
 * 3. Copiar ese id en SHEET_ID acá abajo, guardar y publicar como Web App.
 *
 * DESPUÉS:
 * - Objetivo y prioridades se editan desde la web y se escriben en la hoja 02.
 * - Las tareas se crean, editan y borran desde la web, en la hoja 02.
 * - Las reuniones se cargan desde la web, en la hoja 01.
 * - Los links del entregable van en OBSERVACIONES, como texto clickeable.
 * - Autenticación: un token compartido (ACCESS_TOKEN), el mismo que la web.
 */

/* Para ejecutar a mano desde el editor (el menú no muestra funciones con "_").
   Van primero porque el editor corre la primera función del archivo. */
function crearSheet() { Logger.log(JSON.stringify(crearSheet_(), null, 2)); }
function verEstado() { Logger.log(JSON.stringify(diagnostico_(), null, 2)); }

/* Deja el desplegable de RESPONSABLE de la hoja 02 con el equipo del cliente,
   para que el Sheet ofrezca lo mismo que el tablero. */
var RESPONSABLES_SHEET = ["Vero", "Eve"];
function ponerResponsables() {
  var L = procesoLayout_();
  if (L.cols.responsable === undefined) throw new Error("No encuentro la columna RESPONSABLE");
  var regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(RESPONSABLES_SHEET, true).setAllowInvalid(true).build();
  var desde = L.C.header + 2; // primera fila de tareas
  L.sheet.getRange(desde, L.cols.responsable, 200, 1).setDataValidation(regla);
  Logger.log(JSON.stringify({ opciones: RESPONSABLES_SHEET, desdeFila: desde }));
}

/* Diagnóstico: muestra cómo vienen los mails de Fathom para escribir el
   lector automático. Solo lee, no toca nada. */
function verFathom() {
  var hilos = GmailApp.search('from:(fathom.video OR fathom.ai) newer_than:120d', 0, 5);
  if (!hilos.length) {
    Logger.log("No encontré mails de Fathom en los últimos 120 días. Probá buscar en Gmail: from:fathom.video");
    return;
  }
  var out = hilos.map(function (h) {
    var m = h.getMessages()[0];
    var cuerpo = String(m.getPlainBody() || "");
    return {
      fecha: Utilities.formatDate(m.getDate(), "America/Argentina/Buenos_Aires", "yyyy-MM-dd HH:mm"),
      de: m.getFrom(),
      asunto: m.getSubject(),
      links: (cuerpo.match(/https?:\/\/[^\s)>\]]+/g) || []).slice(0, 8),
      arranque: cuerpo.slice(0, 700)
    };
  });
  Logger.log(JSON.stringify(out, null, 2));
}

var ACCESS_TOKEN = "LiLjpLsOWEB9vj3FI2cMIfbrOqxEuRho";

/* Id del Sheet del cliente. Vacío hasta correr crearSheet(). */
var SHEET_ID = "1BIQZUKiUkIdIHQCdbHi2ci6cPsoOjjxCi95fYblbJ4g";
var CARPETA_DRIVE = "1dPqob8cO7hIGef07WFiTeTHIg4B7BBbC";
var NOMBRE_SHEET = "Evelyna I CONSULTORÍA: JoieM";

var SS_ = null;
function ss_() {
  if (SS_) return SS_;
  if (!SHEET_ID) throw new Error("Todavía no hay Sheet: ejecutá crearSheet() y pegá el id en SHEET_ID.");
  SS_ = SpreadsheetApp.openById(SHEET_ID);
  return SS_;
}

var HOJA_ETAPA1 = "01 I Plan de trabajo";
var HOJA_PROCESO = "02 I Proceso de trabajo";
var SHEET_OVERRIDES = "WebApp - Overrides";
var SHEET_NOTES = "WebApp - Notas";

var SHEET_SCHEMAS = {};
SHEET_SCHEMAS[SHEET_OVERRIDES] = ["task_id", "estado", "link", "inicio", "cierre", "hidden", "updated_at"];
SHEET_SCHEMAS[SHEET_NOTES] = ["id", "text", "author", "createdAt"];

/* Vocabulario del tablero. Los mismos estados que usa la web. */
var ESTADOS_SHEET = ["Pendiente", "En proceso", "Revisar", "Testear", "Finalizada"];
/* La web maneja estados internos; la hoja usa estas palabras. */
var ESTADO_A_SHEET = {
  "Por hacer": "Pendiente", "En proceso": "En proceso", "En revisión": "Revisar",
  "Testear": "Testear", "Completado": "Finalizada"
};
var PRIORIDADES_SHEET = ["P1", "P2", "P3"];

/* Encabezados de la hoja 02 (en este orden). */
var COLUMNAS_02 = ["PRIORIDAD", "TEMA", "TAREA", "RESPONSABLE", "INICIO", "CIERRE", "ESTADO", "OBSERVACIONES"];
/* Encabezados de la hoja 01. */
var COLUMNAS_01 = ["Fecha", "Hs", "Tarea", "Responsable", "Estado", "Resultado", "Observaciones"];

/* =====================================================================
   CREACIÓN DEL SHEET (una sola vez)
   ===================================================================== */

function crearSheet_() {
  if (SHEET_ID) {
    return { yaExistia: true, id: SHEET_ID, url: SpreadsheetApp.openById(SHEET_ID).getUrl() };
  }
  var ss = SpreadsheetApp.create(NOMBRE_SHEET);
  var archivo = DriveApp.getFileById(ss.getId());
  try {
    DriveApp.getFolderById(CARPETA_DRIVE).addFile(archivo);
    DriveApp.getRootFolder().removeFile(archivo);
  } catch (err) {
    // Si no puede moverlo, queda en "Mi unidad": no es motivo para fallar.
  }
  armarEtapa1_(ss);
  armarProceso_(ss);
  var primera = ss.getSheets()[0];
  if (primera.getName() === "Hoja 1" || primera.getName() === "Sheet1") ss.deleteSheet(primera);
  ensureSheets_(ss);
  SpreadsheetApp.flush();
  return { creado: true, id: ss.getId(), url: ss.getUrl(),
    aviso: "Pegá este id en SHEET_ID, guardá y publicá la versión nueva." };
}

/* Hoja 01: la tabla de reuniones. */
function armarEtapa1_(ss) {
  var sh = ss.insertSheet(HOJA_ETAPA1);
  sh.getRange(1, 1).setValue("JoieM I Plan de trabajo").setFontSize(16).setFontWeight("bold");
  sh.getRange(3, 1, 1, COLUMNAS_01.length).setValues([COLUMNAS_01])
    .setFontWeight("bold").setBackground("#212121").setFontColor("#FFFFFF");
  sh.setFrozenRows(3);
  sh.getRange(4, 5, 200, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Pendiente", "Finalizado"], true).setAllowInvalid(true).build());
  [90, 50, 320, 140, 110, 320, 260].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  return sh;
}

/* Hoja 02: objetivo, prioridades, iniciativas y finalizadas. */
function armarProceso_(ss) {
  var sh = ss.insertSheet(HOJA_PROCESO);
  sh.getRange(2, 2).setValue("OBJETIVO 1").setFontWeight("bold").setBackground("#212121").setFontColor("#FFFFFF");
  sh.getRange(3, 2).setValue("¿Qué estamos tratando de conseguir?").setBackground("#FEF200");
  sh.getRange(4, 2).setValue("");
  sh.getRange(6, 2).setValue("PRIORIDADES").setFontWeight("bold").setBackground("#212121").setFontColor("#FFFFFF");
  sh.getRange(7, 2).setValue("¿Cuáles son las acciones más importantes que queremos lograr?").setBackground("#FEF200");
  [1, 2, 3].forEach(function (n, i) { sh.getRange(8 + i, 2).setValue(n); });
  sh.getRange(11, 2).setValue("INICIATIVAS").setFontWeight("bold").setBackground("#212121").setFontColor("#FFFFFF");
  sh.getRange(12, 3, 1, COLUMNAS_02.length).setValues([COLUMNAS_02])
    .setFontWeight("bold").setBackground("#212121").setFontColor("#FFFFFF");
  sh.getRange(30, 2).setValue("FINALIZADAS").setFontWeight("bold").setBackground("#23914A").setFontColor("#FFFFFF");
  sh.setFrozenRows(12);
  sh.setColumnWidth(1, 30); sh.setColumnWidth(2, 110);
  [90, 150, 380, 130, 100, 100, 120, 320].forEach(function (w, i) { sh.setColumnWidth(i + 3, w); });
  validacionesProceso_(sh, 13, 60);
  return sh;
}

/* Desplegables de PRIORIDAD y ESTADO en las filas de tareas. */
function validacionesProceso_(sh, desde, filas) {
  var colPrio = 3, colEstado = 3 + COLUMNAS_02.indexOf("ESTADO");
  sh.getRange(desde, colPrio, filas, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PRIORIDADES_SHEET, true).setAllowInvalid(true).build());
  sh.getRange(desde, colEstado, filas, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS_SHEET, true).setAllowInvalid(true).build());
}

/* =====================================================================
   ENTRY POINTS
   ===================================================================== */

function doGet(e) {
  try {
    checkToken_(e.parameter.token);
    if (e.parameter.action !== "read") {
      return jsonOut_({ ok: false, error: "acción GET no soportada: " + e.parameter.action });
    }
    ensureSheets_(ss_());
    return jsonOut_({
      ok: true,
      seed: readSeed_(),
      overrides: readOverrides_(),
      notes: readSimpleRows_(SHEET_NOTES, SHEET_SCHEMAS[SHEET_NOTES])
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var body = JSON.parse(e.postData.contents || "{}");
    checkToken_(body.token);
    ensureSheets_(ss_());
    return jsonOut_({ ok: true, result: handleAction_(body.action, body.payload || {}) });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function handleAction_(action, p) {
  switch (action) {
    case "updateTask": return updateTask_(p.id, p.fields || {});
    case "addTask": return addTask_(p.fields || {});
    case "deleteTask": return deleteTask_(p.id);
    case "setObjetivo": return setObjetivo_(p.objetivo);
    case "setPrioridad": return setPrioridad_(p.n, p.texto);
    case "setOverride": return setOverride_(p.taskId, p.patch || {});
    case "addNote": return addRow_(SHEET_NOTES, SHEET_SCHEMAS[SHEET_NOTES],
      Object.assign({ id: "n" + Date.now(), createdAt: nowIso_() }, p));
    case "deleteNote": return deleteRow_(SHEET_NOTES, p.id);
    case "addMeetingSheet": return addMeetingSheet_(p);
    case "deleteMeetingSheet": return deleteMeetingSheet_(p);
    default: throw new Error("acción desconocida: " + action);
  }
}

function checkToken_(token) {
  if (!token || token !== ACCESS_TOKEN) throw new Error("token inválido");
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function nowIso_() { return new Date().toISOString(); }

/* =====================================================================
   UTILIDADES
   ===================================================================== */

function cell_(row, idx) {
  if (idx === undefined || idx === null) return null;
  var v = row[idx];
  if (v === null || v === undefined) return null;
  if (typeof v === "string") { v = v.trim(); return v === "" ? null : v; }
  return v;
}

function norm_(s) {
  return String(s === null || s === undefined ? "" : s)
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ").trim().toUpperCase();
}

function toIsoDate_(value) {
  if (value === null || value === undefined || value === "") return null;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "America/Argentina/Buenos_Aires", "yyyy-MM-dd");
  }
  var s = String(value).trim();
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    var d = m[1], mo = m[2], y = m[3];
    if (y.length === 2) y = "20" + y;
    return y + "-" + ("0" + mo).slice(-2) + "-" + ("0" + d).slice(-2);
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}

function toSheetDate_(iso) {
  if (!iso) return "";
  var m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return iso;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function listOptions_(cell) {
  var dv = cell.getDataValidation();
  if (!dv || dv.getAllowInvalid()) return null;
  var type = dv.getCriteriaType();
  var crit = dv.getCriteriaValues();
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return crit[0];
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    return crit[0].getValues().map(function (r) { return r[0]; }).filter(function (v) { return v !== ""; });
  }
  return null;
}

function dropdownValues_(cell) {
  var dv = cell.getDataValidation();
  if (!dv) return [];
  var type = dv.getCriteriaType();
  var crit = dv.getCriteriaValues();
  var vals = [];
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) vals = crit[0];
  else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    vals = crit[0].getValues().map(function (r) { return r[0]; });
  }
  return vals.map(function (v) { return String(v).trim(); }).filter(Boolean);
}

function setSafe_(cell, candidates) {
  var options = listOptions_(cell);
  var value = candidates[0];
  if (options) {
    var hit = null;
    candidates.forEach(function (c) {
      if (hit !== null) return;
      options.forEach(function (o) { if (hit === null && norm_(o) === norm_(c)) hit = o; });
    });
    if (hit === null) return false;
    value = hit;
  }
  try { cell.setValue(value); return true; } catch (err) { return false; }
}

/* =====================================================================
   LECTURA
   ===================================================================== */

var NOMBRES_COL = {
  "PRIORIDAD": "prioridad", "TEMA": "tema", "TAREA": "tarea", "RESPONSABLE": "responsable",
  "INICIO": "inicio", "CIERRE": "cierre", "DEADLINE": "cierre", "FIN": "cierre",
  "ESTADO": "estado", "OBSERVACIONES": "obs"
};

function colsProceso_(values) {
  for (var r = 0; r < values.length; r++) {
    var fila = values[r], map = null;
    for (var c = 0; c < fila.length; c++) {
      var k = NOMBRES_COL[norm_(fila[c])];
      if (!k) continue;
      map = map || {};
      if (map[k] === undefined) map[k] = c;
    }
    if (map && map.tarea !== undefined && map.estado !== undefined) { map.header = r; return map; }
  }
  throw new Error("No encuentro el encabezado (TAREA / ESTADO) en '" + HOJA_PROCESO + "'");
}

function colsEscritura_(C) {
  var out = {};
  ["prioridad", "tema", "tarea", "responsable", "inicio", "cierre", "estado", "obs"].forEach(function (k) {
    if (C[k] !== undefined) out[k] = C[k] + 1;
  });
  return out;
}

function findLabelRow_(values, texto, desde) {
  var target = norm_(texto);
  for (var i = desde || 0; i < values.length; i++) {
    for (var c = 0; c < Math.min(values[i].length, 4); c++) {
      if (norm_(values[i][c]) === target) return i;
    }
  }
  return -1;
}

function tieneTarea_(row, C) {
  return !!(cell_(row, C.tarea) || (C.tema !== undefined && cell_(row, C.tema)));
}

function leerTarea_(row, C, richObs) {
  function v(k) { return C[k] === undefined ? null : cell_(row, C[k]); }
  var obsText = v("obs");
  var links = obsLinks_(richObs, obsText);
  return {
    link: links.length ? links.join("\n") : null,
    prioridad: v("prioridad"), tema: v("tema"), tarea: v("tarea"), responsable: v("responsable"),
    inicio: toIsoDate_(v("inicio")), cierre: toIsoDate_(v("cierre")),
    estado: v("estado"), obs: obsSinLinks_(obsText, links)
  };
}

function readSeed_() {
  var sheet = ss_().getSheetByName(HOJA_PROCESO);
  if (!sheet) throw new Error("No encuentro la hoja '" + HOJA_PROCESO + "'");
  var values = sheet.getDataRange().getValues();
  var C = colsProceso_(values);
  var richObs = C.obs === undefined ? null : sheet.getRange(1, C.obs + 1, values.length, 1).getRichTextValues();

  var objetivoRow = findLabelRow_(values, "OBJETIVO 1");
  var prioridadesRow = findLabelRow_(values, "PRIORIDADES");
  var iniciativasRow = findLabelRow_(values, "INICIATIVAS");
  var finalizadasRow = findLabelRow_(values, "FINALIZADAS");

  var objetivo = objetivoRow === -1 ? "" : String(cell_(values[objetivoRow + 2], 1) || "");

  var prioridades = [];
  if (prioridadesRow !== -1) {
    for (var i = 0; i < 3; i++) {
      var fila = values[prioridadesRow + 2 + i] || [];
      prioridades.push({ n: i + 1, tt: String(cell_(fila, 2) || ""), desc: "" });
    }
  }

  var iniciativas = [], finalizados = [];
  var desde = Math.max(iniciativasRow === -1 ? C.header : iniciativasRow, C.header) + 1;
  var hasta = finalizadasRow === -1 ? values.length : finalizadasRow;
  for (var ir = desde; ir < hasta; ir++) {
    if (!tieneTarea_(values[ir], C)) continue;
    iniciativas.push(leerTarea_(values[ir], C, richObs ? richObs[ir][0] : null));
  }
  if (finalizadasRow !== -1) {
    for (var fr = finalizadasRow + 1; fr < values.length; fr++) {
      if (!tieneTarea_(values[fr], C)) continue;
      finalizados.push(leerTarea_(values[fr], C, richObs ? richObs[fr][0] : null));
    }
  }

  return {
    objetivo: objetivo, prioridades: prioridades,
    iniciativas: iniciativas, finalizados: finalizados,
    etapa1: readEtapa1_(), opciones: opciones_(sheet, C, desde)
  };
}

function opciones_(sheet, C, filaEjemplo) {
  var out = { responsable: [], tema: [], estado: [] };
  try {
    ["responsable", "tema", "estado"].forEach(function (k) {
      if (C[k] !== undefined) out[k] = dropdownValues_(sheet.getRange(filaEjemplo + 1, C[k] + 1));
    });
  } catch (err) {}
  return out;
}

/* =====================================================================
   REUNIONES (hoja 01)
   ===================================================================== */

function etapa1Table_() {
  var sheet = ss_().getSheetByName(HOJA_ETAPA1);
  if (!sheet) throw new Error("No encuentro la hoja '" + HOJA_ETAPA1 + "'");
  var values = sheet.getDataRange().getValues();
  var header = -1;
  for (var i = 0; i < values.length; i++) {
    if (norm_(values[i][0]) === "FECHA") { header = i; break; }
  }
  if (header === -1) throw new Error("No encuentro la tabla de reuniones en '" + HOJA_ETAPA1 + "'");
  var last = header;
  while (last + 1 < values.length && cell_(values[last + 1], 0)) last++;
  return { sheet: sheet, values: values, header: header, last: last };
}

function readEtapa1_() {
  var T = etapa1Table_();
  var out = [];
  for (var r = T.header + 1; r <= T.last; r++) {
    var row = T.values[r];
    out.push({
      fecha: toIsoDate_(cell_(row, 0)), hs: cell_(row, 1), tarea: cell_(row, 2),
      responsable: cell_(row, 3), estado: cell_(row, 4), resultado: cell_(row, 5), obs: cell_(row, 6)
    });
  }
  return out;
}

function horas_(duracion) {
  var s = String(duracion || "").trim().toLowerCase().replace("hs", "").replace("h", "").trim();
  var n = Number(s.replace(",", "."));
  return isNaN(n) ? (duracion || "") : n;
}

function addMeetingSheet_(p) {
  var T = etapa1Table_();
  var row = T.last + 2;
  var valores = [toSheetDate_(p.fecha), horas_(p.duracion), p.tarea || "Encuentro de trabajo",
    p.responsable || "", p.estado || "Finalizado", p.resumen || "", ""];
  T.sheet.getRange(row, 1, 1, valores.length).setValues([valores]);
  SpreadsheetApp.flush();
  return { row: row, guardado: T.sheet.getRange(row, 1, 1, 7).getDisplayValues()[0] };
}

function deleteMeetingSheet_(p) {
  var T = etapa1Table_();
  for (var r = T.header + 1; r <= T.last; r++) {
    var row = T.values[r];
    if (toIsoDate_(row[0]) === p.fecha &&
        String(cell_(row, 2) || "") === String(p.tarea || "") &&
        String(cell_(row, 5) || "") === String(p.resultado || "")) {
      T.sheet.deleteRow(r + 1);
      return { deleted: true };
    }
  }
  throw new Error("No encuentro esa reunión en '" + HOJA_ETAPA1 + "'. Recargá la página.");
}

/* =====================================================================
   OBJETIVO Y PRIORIDADES (se editan desde la web)
   ===================================================================== */

function setObjetivo_(texto) {
  var sheet = ss_().getSheetByName(HOJA_PROCESO);
  var values = sheet.getDataRange().getValues();
  var fila = findLabelRow_(values, "OBJETIVO 1");
  if (fila === -1) throw new Error("No encuentro el bloque OBJETIVO 1 en la hoja 02");
  sheet.getRange(fila + 3, 2).setValue(texto || "");
  return { objetivo: texto || "" };
}

function setPrioridad_(n, texto) {
  var num = Number(n);
  if (!(num >= 1 && num <= 3)) throw new Error("La prioridad tiene que ser 1, 2 o 3");
  var sheet = ss_().getSheetByName(HOJA_PROCESO);
  var values = sheet.getDataRange().getValues();
  var fila = findLabelRow_(values, "PRIORIDADES");
  if (fila === -1) throw new Error("No encuentro el bloque PRIORIDADES en la hoja 02");
  sheet.getRange(fila + 2 + num, 2).setValue(num);
  sheet.getRange(fila + 2 + num, 3).setValue(texto || "");
  return { n: num, texto: texto || "" };
}

/* =====================================================================
   LINKS EN OBSERVACIONES
   ===================================================================== */

var MAX_LINKS = 3;

function obsLinks_(richValue, text) {
  var out = [];
  function add(u) { if (u && out.indexOf(u) === -1 && out.length < MAX_LINKS) out.push(u); }
  if (richValue) {
    add(richValue.getLinkUrl());
    richValue.getRuns().forEach(function (run) { add(run.getLinkUrl()); });
  }
  (String(text || "").match(/https?:\/\/\S+/g) || []).forEach(add);
  return out;
}

function obsSinLinks_(text, links) {
  if (!text) return null;
  var lines = String(text).split(/\n/).filter(function (line) {
    var l = line.trim();
    return l && links.indexOf(l) === -1 && !/^https?:\/\/\S+$/.test(l);
  });
  var t = lines.join("\n").trim();
  return t || null;
}

function escribirObsYLinks_(cell, fields, perdidos) {
  var currentText = String(cell.getValue() || "");
  var currentLinks = obsLinks_(cell.getRichTextValue(), currentText);
  var partes = fields.link !== undefined
    ? String(fields.link || "").split(/\s*\n\s*|\s+\|\s+/).map(function (x) { return x.trim(); }).filter(Boolean)
    : currentLinks;
  var links = [], conTexto = [];
  partes.forEach(function (x) {
    if (/^https?:\/\/\S+$/.test(x)) { if (links.length < MAX_LINKS) links.push(x); }
    else conTexto.push(x);
  });
  var base = fields.obs !== undefined ? (fields.obs || "") : (obsSinLinks_(currentText, currentLinks) || "");
  base = [base].concat(conTexto).concat(perdidos || []).filter(Boolean).join("\n");
  var text = [base].concat(links).filter(Boolean).join("\n");
  if (!text) { cell.setValue(""); return; }
  var builder = SpreadsheetApp.newRichTextValue().setText(text);
  var re = /https?:\/\/[^\s]+/g, m;
  while ((m = re.exec(text)) !== null) {
    try { builder = builder.setLinkUrl(m.index, m.index + m[0].length, m[0]); } catch (err) {}
  }
  cell.setRichTextValue(builder.build());
}

/* =====================================================================
   ESCRITURA DE TAREAS
   ===================================================================== */

function hashId_(str) {
  var h = 5381;
  for (var i = 0; i < str.length; i++) h = (((h << 5) + h) ^ str.charCodeAt(i)) >>> 0;
  return "t" + h.toString(36);
}

function procesoLayout_() {
  var sheet = ss_().getSheetByName(HOJA_PROCESO);
  if (!sheet) throw new Error("No encuentro la hoja '" + HOJA_PROCESO + "'");
  var values = sheet.getDataRange().getValues();
  var C = colsProceso_(values);
  var iniciativasRow = findLabelRow_(values, "INICIATIVAS");
  var finalizadasRow = findLabelRow_(values, "FINALIZADAS");

  var tasks = [], seen = {};
  function scan(desde, hasta, source) {
    var last = desde - 1;
    for (var r = desde; r < hasta; r++) {
      if (!tieneTarea_(values[r], C)) continue;
      var key = [source, cell_(values[r], C.tema) || "", cell_(values[r], C.tarea) || ""].join("|");
      seen[key] = (seen[key] || 0) + 1;
      tasks.push({ id: hashId_(key + "#" + seen[key]), row: r + 1, source: source });
      last = r;
    }
    return last + 1;
  }
  var desde = Math.max(iniciativasRow === -1 ? C.header : iniciativasRow, C.header) + 1;
  var hasta = finalizadasRow === -1 ? values.length : finalizadasRow;
  var lastIniciativaRow = scan(desde, hasta, "iniciativa");
  if (lastIniciativaRow <= desde) lastIniciativaRow = desde;
  var lastFinalizadoRow = finalizadasRow === -1 ? values.length : scan(finalizadasRow + 1, values.length, "finalizado");
  if (finalizadasRow !== -1 && lastFinalizadoRow <= finalizadasRow + 1) lastFinalizadoRow = finalizadasRow + 1;

  return { sheet: sheet, values: values, C: C, cols: colsEscritura_(C), tasks: tasks,
    lastIniciativaRow: lastIniciativaRow, lastFinalizadoRow: lastFinalizadoRow };
}

function findTask_(L, id) {
  for (var i = 0; i < L.tasks.length; i++) if (L.tasks[i].id === id) return L.tasks[i];
  return null;
}

function idAtRow_(row) {
  var L = procesoLayout_();
  for (var i = 0; i < L.tasks.length; i++) if (L.tasks[i].row === row) return L.tasks[i].id;
  return null;
}

function writeTaskCells_(L, row, fields) {
  var perdidos = [];
  Object.keys(L.cols).forEach(function (k) {
    if (k === "obs" || fields[k] === undefined) return;
    var cell = L.sheet.getRange(row, L.cols[k]);
    var v = fields[k];
    if (v === null || v === "") { try { cell.setValue(""); } catch (err) {} return; }
    var candidatos = (k === "inicio" || k === "cierre") ? [toSheetDate_(v)]
      : k === "estado" ? [ESTADO_A_SHEET[v] || v, v]
      : [v];
    if (!setSafe_(cell, candidatos)) perdidos.push(k.charAt(0).toUpperCase() + k.slice(1) + ": " + v);
  });
  if (L.cols.obs !== undefined && (fields.obs !== undefined || fields.link !== undefined || perdidos.length)) {
    escribirObsYLinks_(L.sheet.getRange(row, L.cols.obs), fields, perdidos);
  }
  return perdidos;
}

function moveRow_(sheet, fromRow, afterRow) {
  sheet.insertRowAfter(afterRow);
  var to = afterRow + 1;
  if (fromRow > afterRow) fromRow++;
  var width = sheet.getMaxColumns();
  sheet.getRange(fromRow, 1, 1, width).copyTo(sheet.getRange(to, 1, 1, width));
  sheet.deleteRow(fromRow);
  return fromRow < to ? to - 1 : to;
}

function esFinalizada_(estado) {
  var e = String(estado || "");
  return /^final/i.test(e) || /^completad/i.test(e);
}

function updateTask_(id, fields) {
  var L = procesoLayout_();
  var t = findTask_(L, id);
  if (!t) throw new Error("No encuentro esa tarea en la hoja. Recargá la página.");
  var perdidos = writeTaskCells_(L, t.row, fields);
  var row = t.row;
  if (fields.estado !== undefined) {
    var done = esFinalizada_(fields.estado);
    if (done && t.source === "iniciativa") row = moveRow_(L.sheet, row, L.lastFinalizadoRow);
    else if (!done && t.source === "finalizado") row = moveRow_(L.sheet, row, L.lastIniciativaRow);
  }
  SpreadsheetApp.flush();
  var newId = idAtRow_(row);
  migrateOverride_(id, newId);
  return { id: newId, enObservaciones: perdidos };
}

function addTask_(fields) {
  var L = procesoLayout_();
  var after = esFinalizada_(fields.estado) ? L.lastFinalizadoRow : L.lastIniciativaRow;
  L.sheet.insertRowAfter(after);
  var row = after + 1;
  L.sheet.getRange(row, 1, 1, L.sheet.getMaxColumns()).clearContent();
  validacionesProceso_(L.sheet, row, 1);
  var perdidos = writeTaskCells_(L, row, Object.assign({ estado: "Pendiente" }, fields));
  SpreadsheetApp.flush();
  return { id: idAtRow_(row), enObservaciones: perdidos };
}

function deleteTask_(id) {
  var L = procesoLayout_();
  var t = findTask_(L, id);
  if (!t) throw new Error("No encuentro esa tarea en la hoja. Recargá la página.");
  L.sheet.deleteRow(t.row);
  deleteRow_(SHEET_OVERRIDES, id);
  return { deleted: id };
}

/* =====================================================================
   PESTAÑAS DE LA WEB
   ===================================================================== */

function ensureSheets_(ss) {
  ss = ss || ss_();
  Object.keys(SHEET_SCHEMAS).forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      var sheet = ss.insertSheet(name);
      sheet.appendRow(SHEET_SCHEMAS[name]);
      sheet.setFrozenRows(1);
    }
  });
}

function readSimpleRows_(sheetName, cols) {
  var sheet = ss_().getSheetByName(sheetName);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (!cell_(values[r], 0)) continue;
    var obj = {};
    cols.forEach(function (c, i) { obj[c] = cell_(values[r], i); });
    out.push(obj);
  }
  return out;
}

function readOverrides_() {
  var rows = readSimpleRows_(SHEET_OVERRIDES, SHEET_SCHEMAS[SHEET_OVERRIDES]);
  var out = {};
  rows.forEach(function (r) {
    out[r.task_id] = {
      estado: r.estado || null, link: r.link || null,
      inicio: toIsoDate_(r.inicio), cierre: toIsoDate_(r.cierre),
      hidden: r.hidden === true || String(r.hidden).toUpperCase() === "TRUE",
      updatedAt: r.updated_at || null
    };
  });
  return out;
}

function findRowIndexById_(sheet, id) {
  var values = sheet.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) if (String(values[r][0]) === String(id)) return r + 1;
  return -1;
}

function addRow_(sheetName, cols, obj) {
  var sheet = ss_().getSheetByName(sheetName);
  sheet.appendRow(cols.map(function (c) { return obj[c] === undefined || obj[c] === null ? "" : obj[c]; }));
  return obj;
}

function deleteRow_(sheetName, id) {
  var sheet = ss_().getSheetByName(sheetName);
  if (!sheet) return { deleted: false };
  var rowIdx = findRowIndexById_(sheet, id);
  if (rowIdx === -1) return { deleted: false };
  sheet.deleteRow(rowIdx);
  return { deleted: true };
}

function setOverride_(taskId, patch) {
  var sheet = ss_().getSheetByName(SHEET_OVERRIDES);
  var rowIdx = findRowIndexById_(sheet, taskId);
  var updatedAt = nowIso_();
  if (rowIdx === -1) {
    var row = { task_id: taskId, estado: "", link: "", inicio: "", cierre: "", hidden: false, updated_at: updatedAt };
    Object.keys(patch).forEach(function (k) { row[k] = patch[k] === null ? "" : patch[k]; });
    addRow_(SHEET_OVERRIDES, SHEET_SCHEMAS[SHEET_OVERRIDES], row);
  } else {
    var cols = SHEET_SCHEMAS[SHEET_OVERRIDES];
    Object.keys(patch).forEach(function (key) {
      var colIdx = cols.indexOf(key);
      if (colIdx === -1) return;
      sheet.getRange(rowIdx, colIdx + 1).setValue(patch[key] === null ? "" : patch[key]);
    });
    sheet.getRange(rowIdx, cols.indexOf("updated_at") + 1).setValue(updatedAt);
  }
  return { taskId: taskId, patch: patch };
}

function migrateOverride_(oldId, newId) {
  if (!newId || oldId === newId) return;
  var sheet = ss_().getSheetByName(SHEET_OVERRIDES);
  var rowIdx = findRowIndexById_(sheet, oldId);
  if (rowIdx === -1) return;
  sheet.getRange(rowIdx, 1).setValue(newId);
  sheet.getRange(rowIdx, SHEET_SCHEMAS[SHEET_OVERRIDES].indexOf("updated_at") + 1).setValue(nowIso_());
}

function diagnostico_() {
  if (!SHEET_ID) return { sheet: "todavía no creado: ejecutá crearSheet()" };
  var s = readSeed_();
  return { id: SHEET_ID, url: ss_().getUrl(), objetivo: s.objetivo,
    prioridades: s.prioridades, tareas: s.iniciativas.length,
    finalizadas: s.finalizados.length, reuniones: s.etapa1.length };
}
