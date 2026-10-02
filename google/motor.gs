/**
 * CANCHAS BAHÍA · Motor de datos (Google Apps Script)
 *
 * Una sola hoja de Google guarda todo: canchas, turnos libres, Falta uno,
 * equipos que buscan rival y el registro de consultas.
 *
 * Cómo instalarlo (una sola vez):
 *  1. Crear una hoja de cálculo nueva en Google Sheets.
 *  2. Extensiones > Apps Script. Borrar lo que haya y pegar todo este archivo.
 *  3. Elegir la función "configurar" arriba y tocar "Ejecutar". Aceptar los permisos.
 *  4. Implementar > Nueva implementación > tipo "Aplicación web".
 *     Ejecutar como: Yo. Quién tiene acceso: Cualquier usuario. Implementar.
 *  5. Copiar la URL que termina en /exec y pasársela a Claude.
 */

const ZONA = "America/Argentina/Buenos_Aires";
const DIAS_ADELANTE = 7;

const HOJAS = {
  Canchas:   ["id", "nombre", "barrio", "direccion", "horario", "tipos", "techada", "whatsapp", "precio5", "precio6", "precio7", "precio9", "clave", "activa"],
  Turnos:    ["cargado", "complejo", "tipo", "fecha", "hora", "estado", "precio"],
  FaltaUno:  ["cargado", "cancha", "tipo", "fecha", "hora", "faltan", "nivel", "nombre", "whatsapp", "oculto"],
  Rivales:   ["cargado", "equipo", "tipo", "nivel", "fecha", "horario", "cancha", "zona", "nombre", "whatsapp", "nota", "oculto"],
  Consultas: ["cargado", "complejo", "origen", "tipo", "fecha", "hora", "codigo"]
};

/* ---------- Instalación ---------- */
function configurar() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setSpreadsheetTimeZone(ZONA);
  Object.keys(HOJAS).forEach(nombre => {
    let hoja = libro.getSheetByName(nombre);
    if (!hoja) hoja = libro.insertSheet(nombre);
    hoja.getRange("A:Z").setNumberFormat("@"); // todo como texto: Google no convierte fechas ni horas
    const enc = HOJAS[nombre];
    hoja.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight("bold").setBackground("#e4f24a");
    hoja.setFrozenRows(1);
  });
  const sobrante = libro.getSheetByName("Hoja 1") || libro.getSheetByName("Sheet1");
  if (sobrante && libro.getSheets().length > 1) libro.deleteSheet(sobrante);

  const canchas = libro.getSheetByName("Canchas");
  if (canchas.getLastRow() < 2) {
    canchas.appendRow(["ejemplo", "Complejo de Prueba", "Centro", "Calle Falsa 123", "Lun a dom 10 a 24", "5,7", "si", "2914123456", "50000", "", "84000", "", nuevaClave(), "no"]);
  }
  // Limpieza automática todas las madrugadas
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "limpiar") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("limpiar").timeBased().everyDays(1).atHour(4).inTimezone(ZONA).create();
  limpiar();
}

/** Genera una clave nueva para un dueño. Ejecutala y copiá el resultado del registro. */
function nuevaClave() {
  return Utilities.getUuid().replace(/-/g, "").slice(0, 12);
}

/** Borra todo lo que ya pasó (corre sola todas las madrugadas). Las consultas se guardan siempre. */
function limpiar() {
  const hoy = hoyAR(0);
  borrarFilas("Turnos", fila => fila[3] < hoy);
  borrarFilas("FaltaUno", fila => fila[3] < hoy);
  borrarFilas("Rivales", fila => fila[4] < hoy);
}

/* ---------- Lectura (la página pide los datos) ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.accion === "panel") return json(datosPanel(p.c, p.k));
  return json(datosPublicos());
}

function datosPublicos() {
  const hoy = hoyAR(0), tope = hoyAR(DIAS_ADELANTE);
  const canchas = filas("Canchas").filter(c => c.id && String(c.activa).toLowerCase() !== "no").map(c => ({
    id: c.id, nombre: c.nombre, barrio: c.barrio, direccion: c.direccion, horario: c.horario,
    tipos: String(c.tipos).split(/[,\s]+/).filter(Boolean).map(Number),
    techada: /^s/i.test(String(c.techada)), whatsapp: c.whatsapp,
    precios: { 5: c.precio5, 6: c.precio6, 7: c.precio7, 9: c.precio9 }
  }));
  const ids = new Set(canchas.map(c => c.id));
  // Turnos: vale el último estado cargado para cada horario
  const estado = {};
  filas("Turnos").forEach(t => {
    if (!ids.has(t.complejo) || t.fecha < hoy || t.fecha > tope) return;
    estado[[t.complejo, t.tipo, t.fecha, t.hora].join("|")] = t;
  });
  const turnos = Object.values(estado).filter(t => t.estado === "libre" || t.estado === "oferta")
    .map(t => ({ complejo: t.complejo, tipo: t.tipo, fecha: t.fecha, hora: t.hora, estado: t.estado, precio: t.precio }));
  const falta = filas("FaltaUno").filter(f => !f.oculto && f.fecha >= hoy)
    .map(f => ({ cancha: f.cancha, tipo: f.tipo, fecha: f.fecha, hora: f.hora, faltan: f.faltan, nivel: f.nivel, nombre: f.nombre, whatsapp: f.whatsapp }));
  const rivales = filas("Rivales").filter(r => !r.oculto && r.fecha >= hoy)
    .map(r => ({ equipo: r.equipo, tipo: r.tipo, nivel: r.nivel, fecha: r.fecha, horario: r.horario, cancha: r.cancha, zona: r.zona, nombre: r.nombre, whatsapp: r.whatsapp, nota: r.nota, creado: r.cargado }));
  return { ok: true, canchas, turnos, falta, rivales };
}

function datosPanel(id, clave) {
  if (!claveValida(id, clave)) return { ok: false, error: "clave" };
  const mes = Utilities.formatDate(new Date(), ZONA, "yyyy-MM");
  const consultas = filas("Consultas").filter(c => c.complejo === id && String(c.cargado).slice(0, 7) === mes).length;
  return { ok: true, consultasMes: consultas };
}

/* ---------- Escritura (la página manda datos) ---------- */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: "datos" }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    switch (d.accion) {
      case "turno":    return json(guardarTurno(d));
      case "falta":    return json(guardarFalta(d));
      case "rival":    return json(guardarRival(d));
      case "consulta": return json(guardarConsulta(d));
      default:         return json({ ok: false, error: "accion" });
    }
  } finally { lock.releaseLock(); }
}

function guardarTurno(d) {
  if (!claveValida(d.complejo, d.clave)) return { ok: false, error: "clave" };
  if (!fechaValida(d.fecha) || !/^\d{2}$/.test(d.hora)) return { ok: false, error: "fecha" };
  if (["libre", "oferta", "ocupado"].indexOf(d.estado) < 0) return { ok: false, error: "estado" };
  agregar("Turnos", [ahora(), d.complejo, corto(d.tipo, 2), d.fecha, d.hora, d.estado, soloNumeros(d.precio).slice(0, 8)]);
  return { ok: true };
}

function guardarFalta(d) {
  if (!fechaValida(d.fecha) || !/^\d{2}:\d{2}$/.test(d.hora)) return { ok: false, error: "fecha" };
  if (!telValido(d.whatsapp) || !d.cancha) return { ok: false, error: "datos" };
  agregar("FaltaUno", [ahora(), corto(d.cancha, 60), corto(d.tipo, 2), d.fecha, d.hora, corto(d.faltan, 2), corto(d.nivel, 50), corto(d.nombre, 30), d.whatsapp, ""]);
  return { ok: true };
}

function guardarRival(d) {
  if (!fechaValida(d.fecha)) return { ok: false, error: "fecha" };
  if (!telValido(d.whatsapp) || !d.equipo) return { ok: false, error: "datos" };
  agregar("Rivales", [ahora(), corto(d.equipo, 40), corto(d.tipo, 2), corto(d.nivel, 20), d.fecha, corto(d.horario, 20), d.cancha === "si" ? "si" : "no", corto(d.zona, 30), corto(d.nombre, 30), d.whatsapp, corto(d.nota, 80), ""]);
  return { ok: true };
}

function guardarConsulta(d) {
  agregar("Consultas", [ahora(), corto(d.complejo, 40), corto(d.origen, 12), corto(d.tipo, 2), corto(d.fecha, 10), corto(d.hora, 5), corto(d.codigo, 8)]);
  return { ok: true };
}

/* ---------- Ayudantes ---------- */
function filas(nombre) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const datos = hoja.getRange(1, 1, hoja.getLastRow(), hoja.getLastColumn()).getDisplayValues();
  const enc = datos.shift();
  return datos.map(f => { const o = {}; enc.forEach((k, i) => o[k] = String(f[i]).trim()); return o; });
}
function agregar(nombre, fila) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre).appendRow(fila.map(v => "'" + String(v == null ? "" : v)));
}
function borrarFilas(nombre, condicion) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!hoja || hoja.getLastRow() < 2) return;
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getDisplayValues();
  for (let i = datos.length - 1; i >= 0; i--) if (condicion(datos[i])) hoja.deleteRow(i + 2);
}
function claveValida(id, clave) {
  if (!id || !clave) return false;
  return filas("Canchas").some(c => c.id === id && c.clave && c.clave === clave);
}
function hoyAR(dias) { return Utilities.formatDate(new Date(Date.now() + dias * 864e5), ZONA, "yyyy-MM-dd"); }
function ahora() { return Utilities.formatDate(new Date(), ZONA, "yyyy-MM-dd HH:mm"); }
function fechaValida(f) { return /^\d{4}-\d{2}-\d{2}$/.test(f) && f >= hoyAR(0) && f <= hoyAR(DIAS_ADELANTE); }
function telValido(t) { return /^549\d{10}$/.test(String(t)); }
function soloNumeros(v) { return String(v || "").replace(/\D/g, ""); }
function corto(v, n) { return String(v == null ? "" : v).replace(/[\r\n]+/g, " ").replace(/^[=+\-@]+/, "").slice(0, n); }
function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
