/**
 * Google Apps Script: guarda las respuestas del Amigo Secreto en una hoja de cálculo.
 *
 * Instalación:
 * 1. Crea una hoja en https://sheets.google.com
 * 2. Menú Extensiones > Apps Script. Borra lo que haya y pega este código.
 * 3. Cambia CLAVE por una contraseña propia.
 * 4. Implementar > Nueva implementación > Tipo: Aplicación web
 *      Ejecutar como: Yo   |   Quién tiene acceso: Cualquier usuario
 * 5. Copia la URL que termina en /exec y pégala en "urlGoogleScript" dentro de config.json.
 *
 * Ver las respuestas en JSON:  <URL>/exec?clave=TU_CLAVE
 */

const CLAVE = 'CAMBIA_ESTA_CLAVE';
const HOJA = 'Respuestas';

function obtenerHoja_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(HOJA);
    hoja.appendRow(['Fecha', 'Nombres', 'Respuesta']);
  }
  return hoja;
}

function limpiar_(texto) {
  let t = String(texto || '').trim().slice(0, 100);
  // Evita que un nombre se interprete como fórmula en la hoja
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}

function json_(datos) {
  return ContentService.createTextOutput(JSON.stringify(datos))
    .setMimeType(ContentService.MimeType.JSON);
}

// Recibe cada respuesta desde la página
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const datos = JSON.parse(e.postData.contents);
    const nombre = limpiar_(datos.nombre);
    const respuesta = datos.respuesta === 'Sí' ? 'Sí' : 'No';
    if (!nombre) return json_({ ok: false, error: 'Nombre vacío' });
    obtenerHoja_().appendRow([new Date(), nombre, respuesta]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Devuelve todas las respuestas en JSON (solo con la clave correcta)
function doGet(e) {
  if (!e.parameter || e.parameter.clave !== CLAVE) {
    return json_({ ok: false, error: 'Clave incorrecta' });
  }
  const filas = obtenerHoja_().getDataRange().getValues().slice(1);
  return json_(filas.map(f => ({ fecha: f[0], nombres: f[1], respuesta: f[2] })));
}
