/********************************************************************
 * Boda Miguel & Susana — Backend en Google Sheets
 *
 * QUÉ HACE
 *   - Guarda invitados y mesas (lo que editas en admin.html).
 *   - Recibe los RSVP de la invitación (index.html) y los aplica
 *     automáticamente al estado del invitado.
 *   - Expone el registro de confirmaciones para verlo en el admin.
 *
 * INSTALACIÓN (5 min, una sola vez)
 *   1. Crea una hoja de cálculo nueva en Google Sheets (cualquier nombre).
 *   2. Extensiones → Apps Script. Borra el código de ejemplo y pega
 *      TODO este archivo. Guarda (Ctrl+S).
 *   3. Cambia SHARED_KEY por una clave propia y difícil.
 *      Esa MISMA clave va en:
 *        · admin.html → Configuración → Clave compartida
 *        · index.html → const RSVP_KEY
 *   4. Implementar → Nueva implementación → tipo "Aplicación web":
 *        Ejecutar como: Yo
 *        Quién tiene acceso: Cualquier persona
 *      (Necesario para que los invitados confirmen sin login.)
 *   5. Autoriza con tu cuenta cuando lo pida y copia el URL que
 *      termina en /exec.
 *   6. Pega ese URL en:
 *        · admin.html → Configuración → Google Apps Script URL
 *        · index.html → const RSVP_ENDPOINT
 *
 * NOTA: la clave es una medida básica anti-spam, no seguridad real
 * (la invitación es pública y el código es visible). Para una boda
 * es suficiente.
 ********************************************************************/

const SHARED_KEY = 'CAMBIA-ESTA-CLAVE';

const SHEET_INVITADOS = 'Invitados';
const SHEET_MESAS = 'Mesas';
const SHEET_RSVP = 'RSVP';

const HEAD_INV = ['id', 'nombre', 'telefono', 'pases', 'estado', 'mesaId', 'creado'];
const HEAD_MESA = ['id', 'nombre', 'capacidad', 'invitados'];
const HEAD_RSVP = ['timestamp', 'guestId', 'nombre', 'asiste', 'comida', 'pases'];

/* ── Utilidades ── */
function _sheet(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(head);
  }
  return sh;
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _checkKey(key) {
  return String(key || '') === SHARED_KEY;
}

/* ── Lecturas (GET) ── */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || 'ping');

  if (action === 'ping') {
    return _json({ ok: true, time: new Date().toISOString() });
  }
  if (!_checkKey(p.key)) {
    return _json({ ok: false, error: 'unauthorized' });
  }
  if (action === 'state') {
    return _json({ ok: true, guests: _readGuests(), mesas: _readMesas() });
  }
  if (action === 'rsvpLog') {
    return _json({ ok: true, rsvp: _readRsvp() });
  }
  return _json({ ok: false, error: 'unknown action' });
}

/* ── Escrituras (POST) ── */
function doPost(e) {
  let d = {};
  try {
    d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return _json({ ok: false, error: 'bad json' });
  }
  if (!_checkKey(d.key)) {
    return _json({ ok: false, error: 'unauthorized' });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return _json({ ok: false, error: 'busy, retry' });
  }

  try {
    if (d.action === 'saveState') {
      _writeGuests(d.guests || []);
      _writeMesas(d.mesas || []);
      return _json({ ok: true });
    }
    if (d.action === 'rsvp') {
      return _json(_saveRsvp(d));
    }
    return _json({ ok: false, error: 'unknown action' });
  } finally {
    lock.releaseLock();
  }
}

/* ── Invitados ── */
function _readGuests() {
  const sh = _sheet(SHEET_INVITADOS, HEAD_INV);
  const n = sh.getLastRow();
  if (n < 2) return [];
  const v = sh.getRange(2, 1, n - 1, 7).getValues();
  return v
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      id: Number(r[0]),
      name: String(r[1]),
      phone: String(r[2] || ''),
      seats: Number(r[3]) || 1,
      status: String(r[4]) || 'pendiente',
      mesaId: (r[5] === '' || r[5] === null) ? null : Number(r[5]),
      createdAt: String(r[6] || '')
    }));
}

function _writeGuests(guests) {
  const sh = _sheet(SHEET_INVITADOS, HEAD_INV);
  sh.clearContents();
  sh.appendRow(HEAD_INV);
  const rows = guests.map(g => [
    (g.id !== undefined && g.id !== null) ? g.id : '',
    g.name || '',
    g.phone || '',
    g.seats || 1,
    g.status || 'pendiente',
    (g.mesaId === null || g.mesaId === undefined) ? '' : g.mesaId,
    g.createdAt || ''
  ]);
  if (rows.length) sh.getRange(2, 1, rows.length, 7).setValues(rows);
}

/* ── Mesas ── */
function _readMesas() {
  const sh = _sheet(SHEET_MESAS, HEAD_MESA);
  const n = sh.getLastRow();
  if (n < 2) return [];
  const v = sh.getRange(2, 1, n - 1, 4).getValues();
  return v
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      id: Number(r[0]),
      nombre: String(r[1]),
      capacidad: Number(r[2]) || 6,
      invitados: String(r[3] || '')
        .split(',')
        .map(s => Number(s.trim()))
        .filter(x => !isNaN(x))
    }));
}

function _writeMesas(mesas) {
  const sh = _sheet(SHEET_MESAS, HEAD_MESA);
  sh.clearContents();
  sh.appendRow(HEAD_MESA);
  const rows = mesas.map(m => [
    (m.id !== undefined && m.id !== null) ? m.id : '',
    m.nombre || '',
    m.capacidad || 6,
    (m.invitados || []).join(',')
  ]);
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
}

function _removeFromMesas(guestId) {
  const sh = _sheet(SHEET_MESAS, HEAD_MESA);
  const n = sh.getLastRow();
  if (n < 2) return;
  const v = sh.getRange(2, 1, n - 1, 4).getValues();
  v.forEach((r, i) => {
    const ids = String(r[3] || '')
      .split(',')
      .map(s => Number(s.trim()))
      .filter(x => !isNaN(x) && x !== guestId);
    sh.getRange(i + 2, 4).setValue(ids.join(','));
  });
}

/* ── RSVP: registra y auto-aplica al invitado ── */
function _saveRsvp(d) {
  const sh = _sheet(SHEET_RSVP, HEAD_RSVP);
  sh.appendRow([
    new Date(),
    String(d.guestId || ''),
    String(d.nombre || ''),
    String(d.asiste || ''),
    String(d.comida || ''),
    Number(d.pases || 1)
  ]);

  const gid = Number(d.guestId) || 0;
  _applyEstadoAlInvitado(gid, d.nombre, d.asiste);
  return { ok: true };
}

/* Aplica el estado al invitado: primero por id, si no hay match
   por nombre exacto (por si el link no traía id). */
function _applyEstadoAlInvitado(gid, nombre, asiste) {
  const gsh = _sheet(SHEET_INVITADOS, HEAD_INV);
  const n = gsh.getLastRow();
  if (n < 2) return false;
  const ids = gsh.getRange(2, 1, n - 1, 1).getValues();
  const names = gsh.getRange(2, 2, n - 1, 1).getValues();
  const nm = String(nombre || '').trim().toLowerCase();
  for (let i = 0; i < ids.length; i++) {
    const idMatch = gid && Number(ids[i][0]) === gid;
    const nameMatch = !idMatch && !!nm && String(names[i][0]).trim().toLowerCase() === nm;
    if (idMatch || nameMatch) {
      const row = i + 2;
      const estado = asiste === 'si' ? 'si' : 'no';
      gsh.getRange(row, 5).setValue(estado);
      if (estado !== 'si') {
        gsh.getRange(row, 6).setValue('');
        _removeFromMesas(gid || Number(ids[i][0]));
      }
      return true;
    }
  }
  return false;
}

function _readRsvp() {
  const sh = _sheet(SHEET_RSVP, HEAD_RSVP);
  const n = sh.getLastRow();
  if (n < 2) return [];
  const v = sh.getRange(2, 1, n - 1, 6).getValues();
  return v
    .map(r => ({
      timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0] || ''),
      guestId: String(r[1] || ''),
      nombre: String(r[2] || ''),
      asiste: String(r[3] || ''),
      comida: String(r[4] || ''),
      pases: Number(r[5]) || 1
    }))
    .reverse()
    .slice(0, 200);
}
