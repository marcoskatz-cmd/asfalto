/**
 * Reenvío: lo que queda en el Apps Script de un sistema DESPUÉS de pasarlo a la plataforma.
 *
 * La dirección vieja (…/exec) sigue existiendo y reenvía cada pedido a la plataforma, que es la única que
 * guarda datos. Sirve para todo lo que todavía apunta a la dirección vieja: celulares con la app guardada,
 * rutinas en la nube, otros sistemas que no se migraron.
 *
 * Volver atrás: volver a desplegar la versión anterior del script en el mismo despliegue
 * (clasp deploy -i <id del despliegue> -V <versión anterior>) y restaurar la URL en el frontend.
 *
 * Al desplegar se reemplaza https://ingeco-plataforma.vercel.app/api/asfalto por la URL del sistema en la plataforma.
 */
var DESTINO = 'https://ingeco-plataforma.vercel.app/api/asfalto';

function doGet(e) { return reenviar_('get', e); }
function doPost(e) { return reenviar_('post', e); }

function reenviar_(metodo, e) {
  var url = DESTINO + (e && e.queryString ? '?' + e.queryString : '');
  var opciones = { method: metodo, muteHttpExceptions: true, followRedirects: true };
  if (metodo === 'post') {
    opciones.contentType = (e.postData && e.postData.type) || 'text/plain;charset=utf-8';
    opciones.payload = (e.postData && e.postData.contents) || '';
  }
  var r;
  try { r = UrlFetchApp.fetch(url, opciones); }
  catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No se pudo llegar al servidor. Probá de nuevo en un momento.' })).setMimeType(ContentService.MimeType.JSON);
  }
  var tipo = String((r.getHeaders() || {})['Content-Type'] || '');
  if (tipo.indexOf('text/html') === 0) return HtmlService.createHtmlOutput(r.getContentText());
  return ContentService.createTextOutput(r.getContentText()).setMimeType(ContentService.MimeType.JSON);
}
