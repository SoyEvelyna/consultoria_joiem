/**
 * Fathom → Tableros.
 *
 * Un solo proyecto, en hola@soyevelyna.com, que lee los mails de Fathom y
 * carga cada reunión en el tablero del cliente que corresponde.
 *
 * CÓMO FUNCIONA:
 * - Busca en Gmail los recaps de Fathom que todavía no procesó.
 * - Reconoce el cliente por el título de la reunión o por el participante
 *   (ver CLIENTES abajo).
 * - Escribe una fila en la hoja "01 I Plan de trabajo" del Sheet de ese
 *   cliente: fecha, duración aproximada, título, responsable, estado,
 *   el propósito de la reunión como Resultado y el link de Fathom
 *   clickeable en Observaciones.
 * - Etiqueta el mail como procesado para no cargarlo dos veces.
 *
 * CÓMO SE USA:
 * 1. probarFathom()  → muestra qué haría, sin escribir nada.
 * 2. importarFathom() → carga de verdad lo que haya pendiente.
 * 3. activarFathom()  → deja el activador cada hora.
 *    (desactivarFathom() lo apaga.)
 */

function probarFathom() { Logger.log(JSON.stringify(correr_(true), null, 2)); }
function importarFathom() { Logger.log(JSON.stringify(correr_(false), null, 2)); }
function activarFathom() { Logger.log(JSON.stringify(activar_(), null, 2)); }
function desactivarFathom() { Logger.log(JSON.stringify(desactivar_(), null, 2)); }

/* Los cuatro tableros. "claves" son las palabras que delatan al cliente en
   el título de la reunión o en el participante, sin acentos y en minúscula. */
var CLIENTES = [
  { nombre: "Weber", sheetId: "1rD90bzuYCm4HBJNMYvJDe_0dEaY25ED4zIL4xYKDhd0",
    claves: ["weber", "webersi"] },
  { nombre: "Noctis", sheetId: "1ELzLOSN7QbpKgmJ64WXPWhf_Yc1qR1h0UVcutlurhTI",
    claves: ["noctis", "natalia pereira", "manataliapereira", "nataliapereira"] },
  { nombre: "Gulli", sheetId: "1fz9FefdaKjy3MRH371d7hjIdxiFryatsAIqwgAaVj4o",
    claves: ["gulli", "dra gulli", "dragulli", "daniela gulli"] },
  { nombre: "JoieM", sheetId: "1BIQZUKiUkIdIHQCdbHi2ci6cPsoOjjxCi95fYblbJ4g",
    claves: ["joiem", "joie m", "joa m", "maglioco", "mvmaglioco", "veronica maglioco"] }
];

var BUSQUEDA = 'from:(no-reply@fathom.video OR fathom.video) newer_than:30d';
var ETIQUETA = "Fathom cargado";
var ZONA = "America/Argentina/Buenos_Aires";
var MAX_MAILS = 20;

/* =====================================================================
   RECORRIDO PRINCIPAL
   ===================================================================== */

function correr_(soloProbar) {
  var etiqueta = GmailApp.getUserLabelByName(ETIQUETA) || GmailApp.createLabel(ETIQUETA);
  var hilos = GmailApp.search(BUSQUEDA, 0, MAX_MAILS);
  var resultado = { revisados: 0, cargados: [], sinCliente: [], yaEstaban: [], prueba: !!soloProbar };

  hilos.forEach(function (hilo) {
    if (tieneEtiqueta_(hilo, ETIQUETA)) return;
    var msg = hilo.getMessages()[0];
    resultado.revisados++;
    var datos = leerMail_(msg);
    if (!datos.cliente) {
      resultado.sinCliente.push({ asunto: datos.asunto, fecha: datos.fecha });
      return;
    }
    if (yaCargada_(datos)) {
      resultado.yaEstaban.push({ cliente: datos.cliente.nombre, titulo: datos.titulo, fecha: datos.fecha });
      if (!soloProbar) hilo.addLabel(etiqueta);
      return;
    }
    if (soloProbar) {
      resultado.cargados.push({ cliente: datos.cliente.nombre, fecha: datos.fecha, hs: datos.horas,
        titulo: datos.titulo, resumen: (datos.resumen || "").slice(0, 120), link: datos.link });
      return;
    }
    var fila = escribirReunion_(datos);
    hilo.addLabel(etiqueta);
    resultado.cargados.push({ cliente: datos.cliente.nombre, fila: fila, titulo: datos.titulo, link: datos.link });
  });

  return resultado;
}

function tieneEtiqueta_(hilo, nombre) {
  return hilo.getLabels().some(function (l) { return l.getName() === nombre; });
}

/* =====================================================================
   LECTURA DEL MAIL
   ===================================================================== */

function sinAcentos_(s) {
  return String(s === null || s === undefined ? "" : s)
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ").trim().toLowerCase();
}

function leerMail_(msg) {
  var asunto = String(msg.getSubject() || "");
  var cuerpo = String(msg.getPlainBody() || "");
  var fecha = Utilities.formatDate(msg.getDate(), ZONA, "yyyy-MM-dd");
  var titulo = tituloDe_(asunto);
  var aguja = sinAcentos_(asunto + " " + cuerpo.slice(0, 1500));
  var cliente = null;
  CLIENTES.forEach(function (c) {
    if (cliente) return;
    if (c.claves.some(function (k) { return aguja.indexOf(k) !== -1; })) cliente = c;
  });
  return {
    asunto: asunto, fecha: fecha, titulo: titulo, cliente: cliente,
    link: linkDe_(cuerpo), horas: horasDe_(cuerpo), resumen: resumenDe_(cuerpo)
  };
}

/* "Recap for "X"" -> X · "Recap of your meeting with Y" -> Encuentro con Y */
function tituloDe_(asunto) {
  var m = asunto.match(/Recap for\s+"(.+)"/i) || asunto.match(/Recap de\s+"(.+)"/i);
  if (m) return m[1].trim();
  m = asunto.match(/Recap of your meeting with\s+(.+)$/i) || asunto.match(/reuni[oó]n con\s+(.+)$/i);
  if (m) return "Encuentro con " + m[1].trim();
  return asunto.replace(/^Recap[:\s-]*/i, "").trim() || "Encuentro de trabajo";
}

/* El link de la grabación: /calls/<id>, sin el timestamp. */
function linkDe_(cuerpo) {
  var m = String(cuerpo).match(/https:\/\/fathom\.video\/calls\/\d+/);
  return m ? m[0] + "?tab=summary" : null;
}

/* Duración aproximada: el timestamp más alto que aparece en el recap. */
function horasDe_(cuerpo) {
  var marcas = String(cuerpo).match(/timestamp=(\d+(?:\.\d+)?)/g) || [];
  var max = 0;
  marcas.forEach(function (t) {
    var n = Number(String(t).split("=")[1]);
    if (!isNaN(n) && n > max) max = n;
  });
  if (!max) return "";
  return Math.round((max / 3600) * 10) / 10; // horas con un decimal
}

/* El propósito de la reunión, que Fathom pone al principio del recap. */
function resumenDe_(cuerpo) {
  var texto = String(cuerpo).replace(/\s+/g, " ").trim();
  var m = texto.match(/Prop[oó]sito de la reuni[oó]n\s+(.+?)\s+Puntos clave/i)
       || texto.match(/Meeting Purpose\s+(.+?)\s+Key Takeaways/i);
  if (m) return m[1].trim();
  m = texto.match(/FATHOM\s+(.{40,400}?)\s+(Puntos clave|Key Takeaways|Topics)/i);
  return m ? m[1].trim() : "";
}

/* =====================================================================
   ESCRITURA EN LA HOJA 01 DEL CLIENTE
   ===================================================================== */

function hoja01_(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var hojas = ss.getSheets();
  for (var i = 0; i < hojas.length; i++) {
    if (String(hojas[i].getName() || "").trim().indexOf("01") === 0) return hojas[i];
  }
  throw new Error("No encuentro la hoja que empieza con 01 en " + ss.getName());
}

function tablaReuniones_(sheet) {
  var values = sheet.getDataRange().getValues();
  var header = -1;
  for (var i = 0; i < values.length; i++) {
    if (sinAcentos_(values[i][0]) === "fecha") { header = i; break; }
  }
  if (header === -1) throw new Error("No encuentro el encabezado 'Fecha' en " + sheet.getName());
  var last = header;
  while (last + 1 < values.length && String(values[last + 1][0] || "").trim()) last++;
  return { values: values, header: header, last: last };
}

function yaCargada_(datos) {
  try {
    var sheet = hoja01_(datos.cliente.sheetId);
    var T = tablaReuniones_(sheet);
    for (var r = T.header + 1; r <= T.last; r++) {
      var fila = T.values[r];
      var textoFila = String(fila[2] || "") + " " + String(fila[5] || "") + " " + String(fila[6] || "");
      if (datos.link && textoFila.indexOf(datos.link.split("?")[0]) !== -1) return true;
      var mismaFecha = fechaIso_(fila[0]) === datos.fecha;
      if (mismaFecha && sinAcentos_(fila[2]) === sinAcentos_(datos.titulo)) return true;
    }
  } catch (err) {}
  return false;
}

function fechaIso_(valor) {
  if (!valor) return "";
  if (Object.prototype.toString.call(valor) === "[object Date]") {
    return Utilities.formatDate(valor, ZONA, "yyyy-MM-dd");
  }
  var s = String(valor).trim();
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    var y = m[3].length === 2 ? "20" + m[3] : m[3];
    return y + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  }
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : s;
}

/* Escribe la reunión. Si un desplegable rechaza un valor, lo deja vacío
   en vez de romper la carga. */
function escribirReunion_(datos) {
  var sheet = hoja01_(datos.cliente.sheetId);
  var T = tablaReuniones_(sheet);
  var fila = T.last + 2; // 1-based, debajo de la última
  var fecha = new Date(Number(datos.fecha.slice(0, 4)), Number(datos.fecha.slice(5, 7)) - 1,
    Number(datos.fecha.slice(8, 10)), 12, 0, 0);

  ponerSeguro_(sheet, fila, 1, fecha);
  ponerSeguro_(sheet, fila, 2, datos.horas);
  ponerSeguro_(sheet, fila, 3, datos.titulo);
  ponerSeguro_(sheet, fila, 4, "Equipo");
  ponerSeguro_(sheet, fila, 5, "Finalizado");
  ponerSeguro_(sheet, fila, 6, datos.resumen || "");
  escribirLink_(sheet.getRange(fila, 7), datos.link);
  SpreadsheetApp.flush();
  return fila;
}

function ponerSeguro_(sheet, fila, col, valor) {
  if (valor === "" || valor === null || valor === undefined) return;
  var cell = sheet.getRange(fila, col);
  var dv = cell.getDataValidation();
  if (dv && !dv.getAllowInvalid()) {
    var opciones = [];
    try {
      var tipo = dv.getCriteriaType(), crit = dv.getCriteriaValues();
      if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) opciones = crit[0];
      else if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
        opciones = crit[0].getValues().map(function (r) { return r[0]; });
      }
    } catch (err) {}
    var hit = null;
    opciones.forEach(function (o) { if (hit === null && sinAcentos_(o) === sinAcentos_(valor)) hit = o; });
    if (hit === null) return; // el desplegable no lo acepta: se deja vacío
    valor = hit;
  }
  try { cell.setValue(valor); } catch (err) {}
}

/* El link queda como texto clickeable ("Ver grabación"). */
function escribirLink_(cell, link) {
  if (!link) { cell.setValue(""); return; }
  var texto = "Ver grabación en Fathom";
  var rico = SpreadsheetApp.newRichTextValue().setText(texto).setLinkUrl(0, texto.length, link).build();
  try { cell.setRichTextValue(rico); } catch (err) { cell.setValue(link); }
}

/* =====================================================================
   ACTIVADOR
   ===================================================================== */

function activar_() {
  desactivar_();
  ScriptApp.newTrigger("importarFathom").timeBased().everyHours(1).create();
  return { activado: true, cada: "1 hora" };
}

function desactivar_() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "importarFathom") { ScriptApp.deleteTrigger(t); n++; }
  });
  return { desactivados: n };
}
