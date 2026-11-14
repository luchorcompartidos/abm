/**
 * ABM Registros – API de sincronización con Google Sheets
 *
 * INSTALACIÓN (una sola vez):
 * 1. Creá una Google Sheet nueva (ej.: "ABM Registros").
 * 2. Menú Extensiones > Apps Script. Borrá el contenido y pegá todo este archivo. Guardá.
 * 3. (Opcional) Poné una clave en CLAVE y cargá la misma en la app.
 * 4. Implementar > Nueva implementación > tipo "Aplicación web".
 *      Ejecutar como: Yo   ·   Quién tiene acceso: Cualquier usuario
 *    Autorizá los permisos y copiá la URL que termina en /exec.
 * 5. En la app: ⚙️ > pegá la URL > Guardar. Repetí en cada dispositivo.
 *
 * Si después modificás este script: Implementar > Gestionar implementaciones >
 * editar (lápiz) > Versión: "Nueva versión" (así la URL no cambia).
 *
 * HOJAS:
 *  - "_datos": base de sincronización (no la edites a mano).
 *  - "V · <pestaña>": vistas legibles que se regeneran en cada sincronización (solo lectura).
 */

const CLAVE = '';            // opcional: clave compartida con la app
const HOJA = '_datos';
const CAB = ['id', 'kind', 'parentId', 'entityId', 'updatedAt', 'deleted', 'rev', 'device', 'data'];
const PREF_VISTA = 'V · ';

function doGet() {
  const rev = Number(PropertiesService.getScriptProperties().getProperty('rev') || 0);
  return json_({ ok: true, app: 'abm-registros', rev: rev });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (CLAVE && req.token !== CLAVE) return json_({ ok: false, error: 'Clave inválida' });
    if (req.action !== 'sync') return json_({ ok: false, error: 'Acción desconocida' });

    const sh = hoja_();
    const last = sh.getLastRow();
    const vals = last > 1 ? sh.getRange(2, 1, last - 1, CAB.length).getValues() : [];
    const idx = {};
    vals.forEach(function (r, i) { idx[String(r[0])] = i; });

    const props = PropertiesService.getScriptProperties();
    let rev = Number(props.getProperty('rev') || 0);
    let cambios = 0;

    (req.changes || []).forEach(function (c) {
      if (!c || !c.id) return;
      const i = idx[c.id];
      // Gana la versión más reciente (última escritura)
      if (i !== undefined && Number(vals[i][4]) >= Number(c.updatedAt)) return;
      rev++; cambios++;
      const fila = fila_(c, rev);
      if (i !== undefined) vals[i] = fila;
      else { idx[c.id] = vals.length; vals.push(fila); }
    });

    if (cambios) {
      sh.getRange(2, 1, vals.length, CAB.length).setValues(vals);
      props.setProperty('rev', String(rev));
    }

    const desde = Number(req.lastRev || 0);
    const reset = desde > rev; // la planilla se reinició o se cambió
    const items = vals.filter(function (r) { return reset || Number(r[6]) > desde; }).map(item_);

    if (cambios) {
      try { vistas_(vals.map(item_)); } catch (err) { console.error(err); }
    }
    return json_({ ok: true, rev: rev, items: items, cambios: cambios, reset: reset });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function hoja_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA, 0);
    sh.getRange(1, 1, 1, CAB.length).setValues([CAB]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:D').setNumberFormat('@');
    sh.getRange('H:I').setNumberFormat('@');
  }
  return sh;
}

function fila_(c, rev) {
  return [String(c.id), c.kind || '', c.parentId || '', c.entityId || '',
    Number(c.updatedAt) || 0, c.deleted ? 1 : 0, rev, c.dev || '', JSON.stringify(c.d || {})];
}

function item_(r) {
  let d = {};
  try { d = JSON.parse(r[8] || '{}'); } catch (e) {}
  return { id: String(r[0]), kind: String(r[1]), parentId: String(r[2] || ''), entityId: String(r[3] || ''),
    updatedAt: Number(r[4]) || 0, deleted: Number(r[5]) ? 1 : 0, dev: String(r[7] || ''), d: d };
}

/** Regenera una hoja legible por pestaña (Backlog, etc.) con todos sus registros vigentes. */
function vistas_(items) {
  const ss = SpreadsheetApp.getActive();
  const vivos = items.filter(function (i) { return !i.deleted; });
  const ents = vivos.filter(function (i) { return i.kind === 'entity'; })
    .sort(function (a, b) { return (a.d.order || 0) - (b.d.order || 0); });
  const subN = {};
  vivos.filter(function (i) { return i.kind === 'subtab'; }).forEach(function (s) { subN[s.id] = s.d.name; });
  const usados = {};

  ents.forEach(function (e) {
    let n = (PREF_VISTA + String(e.d.name || 'Pestaña')).replace(/[\[\]\*\?\/\\:]/g, ' ').slice(0, 90);
    while (usados[n]) n += '_';
    usados[n] = 1;
    const F = e.d.fields || [];
    const filas = [[e.d.subLabel || 'Sub-pestaña'].concat(F.map(function (f) { return f.l; }), ['Actualizado', 'Dispositivo'])];
    vivos.filter(function (r) { return r.kind === 'record' && r.entityId === e.id && subN[r.parentId] != null; })
      .sort(function (a, b) { return String(subN[a.parentId]).localeCompare(String(subN[b.parentId])) || a.updatedAt - b.updatedAt; })
      .forEach(function (r) {
        filas.push([subN[r.parentId]].concat(F.map(function (f) {
          const v = r.d[f.k];
          if (f.t === 'check') return v ? 'Sí' : '';
          const s = v == null ? '' : v;
          return (typeof s === 'string' && /^[=+\-@]/.test(s)) ? "'" + s : s;
        }), [new Date(r.updatedAt), r.dev]));
      });
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    sh.clearContents();
    sh.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
    sh.getRange(1, 1, 1, filas[0].length).setFontWeight('bold');
    sh.setFrozenRows(1);
  });

  ss.getSheets().forEach(function (sh) {
    const n = sh.getName();
    if (n.indexOf(PREF_VISTA) === 0 && !usados[n]) ss.deleteSheet(sh);
  });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
