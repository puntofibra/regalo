/*************************************************************
 *  REGALO PARTY  —  Backend (API para la web en GitHub Pages)
 *  Apps Script + Google Sheets
 *
 *  QUÉ CAMBIA RESPECTO A LA VERSIÓN ANTERIOR
 *  1) Nuevo router de API (?api=...) con respuesta JSONP y POST,
 *     para que la web alojada en GitHub Pages pueda hablar con
 *     este script desde otro dominio.
 *  2) BUG DE LOS ENLACES ARREGLADO: al borrar enlaces del panel
 *     ya no vuelven a aparecer (ver _parseEnlaces y
 *     adminGuardarConfig).
 *  3) Los correos y los enlaces generados apuntan a la web de
 *     GitHub Pages (WEB_BASE), no a la URL /exec.
 *************************************************************/

const ID_SHEET     = '';   // <-- déjalo vacío si el script está VINCULADO al Sheet
const APP_URL      = 'https://script.google.com/macros/s/AKfycbyIzpI5-PnW3lG1lYGwcZPnlzgs2lZ51tYNzO8ZrpLn2BbN-duhbKbKdR5gPy3xafBUkw/exec';

/* ▼▼▼ CAMBIA ESTO por la URL de tu GitHub Pages (con la barra final) ▼▼▼ */
const WEB_BASE     = 'https://puntofibra.github.io/regalo/';
/* ▲▲▲ ------------------------------------------------------------- ▲▲▲ */

const H_CONFIG     = 'CONFIG';
const H_APORTES    = 'APORTACIONES';
const CACHE_KEY    = 'REGALO_PUBLIC_V1';
const CACHE_SEG    = 20;
const H_EVENTOS    = 'EVENTOS';
const H_APORTES_EV = 'APORTES_EVENTOS';
const H_LEADS      = 'LEADS';
const ADMIN_EMAIL  = 'futurmovil.com@gmail.com';
const PRECIO_EVENTO = 3;
const DIAS_ACCESO_EVENTO = 30;

/* Enlaces públicos de la web (GitHub Pages) */
function urlPublica(evento){ return WEB_BASE + (evento ? ('?evento=' + encodeURIComponent(evento)) : ''); }
function urlAdmin(evento){   return WEB_BASE + 'admin.html' + (evento ? ('?evento=' + encodeURIComponent(evento)) : ''); }

/* ---------- ACCESO AL LIBRO ---------- */
function _ss() {
  return ID_SHEET ? SpreadsheetApp.openById(ID_SHEET) : SpreadsheetApp.getActiveSpreadsheet();
}

/* ---------- CONFIG POR DEFECTO ---------- */
const CONFIG_DEF = [
  ['titulo',            '🎁 Regalo sorpresa de cumpleaños'],
  ['subtitulo',         'Entre todos lo conseguimos. ¡Súmate!'],
  ['nombreHomenajeado', 'Manu'],
  ['objetivo',          '300'],
  ['importeMinimo',     '10'],
  ['bizumNumero',       '600000000'],
  ['bizumTitular',      'Organizador del regalo'],
  ['myposUsuario',      'futurmovil'],
  ['tieneMyPOS',        'SI'],
  ['regaloNombre',      'El regalo soñado'],
  ['regaloDescripcion', 'Describe aquí el regalo: modelo, color, por qué le va a encantar…'],
  ['regaloFoto',        ''],
  ['regaloEnlaces',     '[]'],
  ['regaloEnlace',      ''],
  ['regaloEnlace2',     ''],
  ['regaloEnlace3',     ''],
  ['regaloEnlace4',     ''],
  ['fechaLimite',       ''],
  ['pin',               '1234'],
  ['moderacion',        'NO'],
  ['mostrarImportes',   'SI'],
  ['permitirAnonimo',   'SI'],
  ['mensajeGracias',    '¡Gracias por participar! Tu nombre ya aparece en el muro 💜'],
  ['activo',            'SI'],
  ['mensajeCerrado',    'La recaudación está cerrada. ¡Gracias a todos!'],
  ['tema',              'default']
];

/* ---------- INICIALIZACIÓN ---------- */
function inicializar() {
  _ensureEventoSheets(_ss());
  const ss = _ss();

  let c = ss.getSheetByName(H_CONFIG);
  if (!c) {
    c = ss.insertSheet(H_CONFIG);
    c.getRange(1, 1, 1, 2).setValues([['CLAVE', 'VALOR']])
      .setFontWeight('bold').setBackground('#1a1030').setFontColor('#e9d5ff');
    c.getRange(2, 1, CONFIG_DEF.length, 2).setValues(CONFIG_DEF);
    c.setColumnWidth(1, 190); c.setColumnWidth(2, 460);
    c.setFrozenRows(1);
  } else {
    const act = c.getRange(2, 1, Math.max(c.getLastRow() - 1, 1), 1).getValues().flat();
    CONFIG_DEF.forEach(function (p) {
      if (act.indexOf(p[0]) === -1) c.appendRow(p);
    });
  }

  let a = ss.getSheetByName(H_APORTES);
  if (!a) {
    a = ss.insertSheet(H_APORTES);
    a.getRange(1, 1, 1, 7).setValues([['ID', 'FECHA', 'NOMBRE', 'MENSAJE', 'IMPORTE', 'METODO', 'ESTADO']])
      .setFontWeight('bold').setBackground('#1a1030').setFontColor('#e9d5ff');
    a.setColumnWidths(1, 7, 150);
    a.setColumnWidth(4, 320);
    a.setFrozenRows(1);
  }
  _limpiarCache();
  return 'OK';
}

/* ---------- MULTI-EVENTO ---------- */
function _ensureEventoSheets(ss) {
  let ev = ss.getSheetByName(H_EVENTOS);
  if (!ev) {
    ev = ss.insertSheet(H_EVENTOS);
    ev.getRange(1, 1, 1, 6).setValues([['ID', 'EMAIL', 'PIN', 'FECHA', 'CONFIG_JSON', 'PRIMER_USO']])
      .setFontWeight('bold').setBackground('#1a1030').setFontColor('#e9d5ff');
    ev.setColumnWidths(1, 6, 180);
    ev.setColumnWidth(5, 500);
    ev.setFrozenRows(1);
  } else if (ev.getLastColumn() < 6) {
    ev.getRange(1, 6).setValue('PRIMER_USO').setFontWeight('bold').setBackground('#1a1030').setFontColor('#e9d5ff');
  }
  let ae = ss.getSheetByName(H_APORTES_EV);
  if (!ae) {
    ae = ss.insertSheet(H_APORTES_EV);
    ae.getRange(1, 1, 1, 8).setValues([['ID', 'EVENTO', 'FECHA', 'NOMBRE', 'MENSAJE', 'IMPORTE', 'METODO', 'ESTADO']])
      .setFontWeight('bold').setBackground('#1a1030').setFontColor('#e9d5ff');
    ae.setColumnWidths(1, 8, 140);
    ae.setColumnWidth(5, 320);
    ae.setFrozenRows(1);
  }
  let le = ss.getSheetByName(H_LEADS);
  if (!le) {
    le = ss.insertSheet(H_LEADS);
    le.getRange(1, 1, 1, 5).setValues([['ID', 'FECHA', 'EMAIL', 'ESTADO', 'TOKEN']])
      .setFontWeight('bold').setBackground('#1a1030').setFontColor('#e9d5ff');
    le.setColumnWidths(1, 5, 180);
    le.setFrozenRows(1);
  }
  return { ev: ev, ae: ae, le: le };
}

function _defaultEventoConfig() {
  return {
    titulo: '🎁 Regalo sorpresa',
    subtitulo: 'Entre todos lo conseguimos. ¡Súmate!',
    nombreHomenajeado: '',
    objetivo: '0',
    importeMinimo: '0',
    bizumNumero: '',
    bizumTitular: '',
    regaloNombre: '',
    regaloDescripcion: '',
    regaloFoto: '',
    regaloEnlaces: '[]',
    regaloEnlace: '',
    fechaLimite: '',
    pin: '',
    moderacion: 'NO',
    activo: 'SI',
    permitirAnonimo: 'SI',
    mostrarImportes: 'SI',
    mensajeCerrado: '',
    tieneMyPOS: 'NO'
  };
}

function _findRowById(sheet, id) {
  const n = sheet.getLastRow() - 1;
  if (n <= 0) return -1;
  const ids = sheet.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function _findLeadByEmail(le, email) {
  const n = le.getLastRow() - 1;
  if (n <= 0) return -1;
  const vals = le.getRange(2, 1, n, 3).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][2]).trim().toLowerCase() === String(email).trim().toLowerCase()) return i + 2;
  }
  return -1;
}

function _leerConfigEvento(id) {
  const sh = _ensureEventoSheets(_ss()).ev;
  const row = _findRowById(sh, id);
  const base = _defaultEventoConfig();
  if (row < 0) return null;
  const vals = sh.getRange(row, 1, 1, 6).getValues()[0];
  let cfg = {};
  try { cfg = JSON.parse(vals[4] || '{}'); } catch (e) { cfg = {}; }
  const out = Object.assign({}, base, cfg);
  out.pin = vals[2] || '';
  out.tieneMyPOS = 'NO';
  out._email = vals[1];
  out._fila = row;
  out._primerUso = vals[5] || '';
  return out;
}

function _marcarPrimerUso(id) {
  const sh = _ensureEventoSheets(_ss()).ev;
  const row = _findRowById(sh, id);
  if (row < 0) return;
  const actual = sh.getRange(row, 6).getValue();
  if (!actual) sh.getRange(row, 6).setValue(new Date());
}

function _eventoCaducado(cfg) {
  if (!cfg || !cfg._primerUso) return false;
  const inicio = new Date(cfg._primerUso).getTime();
  if (isNaN(inicio)) return false;
  const limite = inicio + DIAS_ACCESO_EVENTO * 24 * 60 * 60 * 1000;
  return Date.now() > limite;
}

function _guardarConfigEvento(id, datos) {
  const sh = _ensureEventoSheets(_ss()).ev;
  const row = _findRowById(sh, id);
  if (row < 0) return false;
  const actual = _leerConfigEvento(id) || _defaultEventoConfig();
  const merged = Object.assign({}, actual, datos);

  // === FIX ENLACES (eventos) ===
  // Si el panel manda la lista de enlaces, esa lista MANDA: se vacían las
  // claves antiguas para que un enlace borrado no reaparezca.
  if (Object.prototype.hasOwnProperty.call(datos, 'regaloEnlaces')) {
    merged.regaloEnlace  = '';
    merged.regaloEnlace2 = '';
    merged.regaloEnlace3 = '';
    merged.regaloEnlace4 = '';
  }

  delete merged.pin; delete merged._email; delete merged._fila;
  delete merged.tieneMyPOS; delete merged._primerUso;
  sh.getRange(row, 5).setValue(JSON.stringify(merged));
  return true;
}

/* ---------- LEADS / VENTA ---------- */
function registrarInteresEvento(email) {
  email = String(email || '').trim();
  if (!email || email.indexOf('@') < 0) return { ok: false, msg: 'Correo no válido.' };
  const le = _ensureEventoSheets(_ss()).le;
  const fila = _findLeadByEmail(le, email);
  if (fila < 0) le.appendRow([Utilities.getUuid(), new Date(), email, 'INTERES', '']);
  try {
    const cfg = leerConfig();
    MailApp.sendEmail({
      to: email,
      subject: '🎁 Cómo crear tu propia recaudación de regalo',
      htmlBody:
        '<p>¡Hola!</p>' +
        '<p>Gracias por tu interés en crear tu propia página para recaudar un regalo entre varias personas, como esta.</p>' +
        '<p>Con tu panel de administrador podrás poner el nombre del homenajeado, la foto y descripción del regalo, el objetivo económico, y tu número de Bizum. Luego compartes el enlace de tu página por WhatsApp con los invitados y cada uno podrá aportar y aparecer en el muro.</p>' +
        '<p>El acceso al panel tiene un coste único de <b>' + PRECIO_EVENTO + '€</b>, que se paga por Bizum al ' + (cfg.bizumNumero || '') + ' (' + (cfg.bizumTitular || '') + ').</p>' +
        '<p>Cuando hayas hecho el pago, vuelve a la página y pulsa en "Ya pagué, solicitar PIN" con este mismo correo. En cuanto verifiquemos el pago te enviaremos tu PIN de acceso.</p>' +
        '<p>¡Gracias! 💜</p>'
    });
  } catch (e) {}
  return { ok: true, msg: 'Te hemos enviado la información a tu correo ✉️' };
}

function solicitarPinEvento(email) {
  email = String(email || '').trim();
  if (!email || email.indexOf('@') < 0) return { ok: false, msg: 'Correo no válido.' };
  const le = _ensureEventoSheets(_ss()).le;
  const token = Utilities.getUuid();
  let fila = _findLeadByEmail(le, email);
  const id = Utilities.getUuid();
  if (fila < 0) le.appendRow([id, new Date(), email, 'SOLICITADO', token]);
  else le.getRange(fila, 4, 1, 2).setValues([['SOLICITADO', token]]);
  const idFinal = fila < 0 ? id : le.getRange(fila, 1).getValue();
  const link = APP_URL + '?accion=aprobar&id=' + encodeURIComponent(idFinal) + '&token=' + encodeURIComponent(token);
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: '💰 Solicitud de PIN — ' + email,
      htmlBody:
        '<p>' + email + ' indica que ha pagado los ' + PRECIO_EVENTO + '€ y solicita su PIN de acceso.</p>' +
        '<p>Verifica que el pago haya llegado y, si es correcto, pulsa el botón para generarle y enviarle su PIN:</p>' +
        '<p><a href="' + link + '" style="background:#a855f7;color:#fff;padding:12px 20px;border-radius:10px;text-decoration:none;font-weight:bold">Enviar PIN a ' + email + '</a></p>' +
        '<p>Si el enlace no funciona, cópialo aquí: ' + link + '</p>'
    });
  } catch (e) {}
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Hemos recibido tu solicitud 🎁',
      htmlBody: '<p>Hemos recibido tu solicitud de PIN. En cuanto verifiquemos tu pago de ' + PRECIO_EVENTO + '€ recibirás un correo con tu PIN de acceso.</p>'
    });
  } catch (e) {}
  return { ok: true, msg: 'Solicitud enviada. Te avisaremos por correo en cuanto verifiquemos el pago 🙌' };
}

function _aprobarLead(id, token) {
  const le = _ensureEventoSheets(_ss()).le;
  const fila = _findRowById(le, id);
  const html = function (msg) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:system-ui;padding:30px;text-align:center">' + msg + '</div>');
  };
  if (fila < 0) return html('Solicitud no encontrada.');
  const vals = le.getRange(fila, 1, 1, 5).getValues()[0];
  const email = vals[2], estado = vals[3], tok = vals[4];
  if (estado === 'ACTIVO') return html('Esta solicitud ya fue aprobada anteriormente.');
  if (String(tok) !== String(token)) return html('Enlace no válido.');

  const ev = _ensureEventoSheets(_ss()).ev;
  const eventoId = Utilities.getUuid();
  const pin = String(Math.floor(1000 + Math.random() * 9000));
  ev.appendRow([eventoId, email, pin, new Date(), JSON.stringify(_defaultEventoConfig()), '']);
  le.getRange(fila, 4).setValue('ACTIVO');

  const linkAdmin_   = urlAdmin(eventoId);
  const linkPublico_ = urlPublica(eventoId);
  try {
    MailApp.sendEmail({
      to: email,
      subject: '🎉 ¡Tu panel ya está listo! Tu PIN de acceso',
      htmlBody:
        '<p>¡Ya puedes crear tu recaudación!</p>' +
        '<p>Tu PIN temporal es: <b style="font-size:20px">' + pin + '</b> (podrás cambiarlo dentro del panel).</p>' +
        '<p>Entra a tu panel de administrador aquí: <a href="' + linkAdmin_ + '">' + linkAdmin_ + '</a></p>' +
        '<p>Personaliza el nombre, la foto y descripción del regalo, el objetivo y tu número de Bizum. Cuando esté listo, comparte este enlace con los invitados por WhatsApp: <br>' + linkPublico_ + '</p>' +
        '<p>¡Gracias por confiar en nosotros! 💜</p>'
    });
  } catch (e) {}
  return html('✅ PIN enviado correctamente a ' + email);
}

/* ---------- MENÚ DEL SHEET ---------- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎁 Regalo')
    .addItem('Inicializar / reparar hojas', 'inicializar')
    .addItem('Ver enlaces', 'mostrarEnlaces')
    .addItem('Crear panel nuevo (venta)', 'crearPanelNuevo')
    .addToUi();
}

function mostrarEnlaces() {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui;padding:14px;line-height:1.6">' +
    '<b>Web pública (participantes):</b><br><a target="_blank" href="' + urlPublica('') + '">' + urlPublica('') + '</a><br><br>' +
    '<b>Panel de control:</b><br><a target="_blank" href="' + urlAdmin('') + '">' + urlAdmin('') + '</a><br><br>' +
    '<span style="color:#888;font-size:12px">API (no compartir): ' + APP_URL + '</span>' +
    '</div>').setWidth(560).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enlaces de la app');
}

function crearPanelWeb(pin) {
  if (!_pinOk(pin, '')) return { ok: false, msg: 'Sesión no válida' };
  const ev = _ensureEventoSheets(_ss()).ev;
  const eventoId = Utilities.getUuid();
  const pinNuevo = String(Math.floor(1000 + Math.random() * 9000));
  const cfg = _defaultEventoConfig();
  cfg.tieneMyPOS = 'NO'; cfg.myposUsuario = ''; cfg.pin = '';
  ev.appendRow([eventoId, '', pinNuevo, new Date(), JSON.stringify(cfg), '']);
  return { ok: true, evento: eventoId, pin: pinNuevo,
           linkPublico: urlPublica(eventoId), linkAdmin: urlAdmin(eventoId) };
}

function crearPanelNuevo() {
  const r = _crearPanelInterno();
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui;padding:16px;line-height:1.7;font-size:14px">' +
    '<p><b>Panel creado ✅</b></p>' +
    '<p><b>Enlace público</b>:<br><a target="_blank" href="' + r.linkPublico + '">' + r.linkPublico + '</a></p>' +
    '<p><b>Panel de administración</b>:<br><a target="_blank" href="' + r.linkAdmin + '">' + r.linkAdmin + '</a></p>' +
    '<p><b>PIN de acceso al admin:</b> ' + r.pin + '</p>' +
    '<p style="color:#888">El acceso caduca automáticamente 30 días después de que se introduzca el PIN por primera vez.</p>' +
    '</div>').setWidth(560).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Panel de venta creado');
}

function _crearPanelInterno() {
  const ev = _ensureEventoSheets(_ss()).ev;
  const eventoId = Utilities.getUuid();
  const pin = String(Math.floor(1000 + Math.random() * 9000));
  const cfg = _defaultEventoConfig();
  cfg.tieneMyPOS = 'NO'; cfg.myposUsuario = ''; cfg.pin = '';
  ev.appendRow([eventoId, '', pin, new Date(), JSON.stringify(cfg), '']);
  return { evento: eventoId, pin: pin, linkPublico: urlPublica(eventoId), linkAdmin: urlAdmin(eventoId) };
}

function enviarPanelPorEmail(pinSesion, evento, email, linkPub, linkAdm) {
  if (!_pinOk(pinSesion, '')) return { ok: false, msg: 'Sesión no válida' };
  evento = String(evento || '').trim();
  email  = String(email || '').trim();
  if (!evento) return { ok: false, msg: 'Panel no encontrado.' };
  if (!email || email.indexOf('@') < 0) return { ok: false, msg: 'Correo no válido.' };
  const cfg = _leerConfigEvento(evento);
  if (!cfg) return { ok: false, msg: 'Panel no encontrado.' };
  const linkPublico_ = linkPub || urlPublica(evento);
  const linkAdmin_   = linkAdm || urlAdmin(evento);
  try {
    MailApp.sendEmail({
      to: email,
      subject: '🎁 Tu panel de recaudación ya está listo',
      htmlBody:
        '<p>¡Hola!</p>' +
        '<p>Aquí tienes los datos de tu panel de recaudación para el regalo:</p>' +
        '<p><b>1) Enlace para compartir con los invitados</b>:<br><a href="' + linkPublico_ + '">' + linkPublico_ + '</a></p>' +
        '<p><b>2) Enlace de tu panel de administrador</b> (solo para ti):<br><a href="' + linkAdmin_ + '">' + linkAdmin_ + '</a></p>' +
        '<p><b>3) Tu PIN de acceso</b>: <b>' + cfg.pin + '</b></p>' +
        '<p>Desde tu panel podrás poner el nombre del homenajeado, la foto y descripción del regalo, el objetivo, tu número de Bizum, y validar las aportaciones. ¡Mucha suerte! 💜</p>'
    });
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: 'Error al enviar: ' + e.message };
  }
}

/* =========================================================
   ROUTER WEB  ·  doGet / doPost
   ========================================================= */
function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.accion === 'aprobar') return _aprobarLead(p.id || '', p.token || '');
  if (p.api) return _jsonpOut(p.callback || '', _apiExec(p));

  // Página de cortesía: la web real vive en GitHub Pages.
  return HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font-family:system-ui;padding:40px;text-align:center;line-height:1.6">' +
    '<h2>🎁 Regalo Party</h2>' +
    '<p>Esta dirección es solo el motor de la aplicación.</p>' +
    '<p><a href="' + urlPublica('') + '">Ir a la página del regalo →</a></p>' +
    '</div>')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { p = {}; }
  if (p.d && typeof p.d === 'object') p.d = JSON.stringify(p.d);
  const out = _apiExec(p);
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function _jsonpOut(cb, obj) {
  const body = JSON.stringify(obj);
  if (cb && /^[A-Za-z0-9_$.]{1,60}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function _apiExec(p) {
  const ev  = String(p.evento || '').trim();
  const pin = String(p.pin || '');
  let d = {};
  if (p.d) { try { d = JSON.parse(p.d); } catch (err) { d = {}; } }
  try {
    switch (String(p.api || '')) {
      case 'datos':        return getDatosPublicos(String(p.force) === '1', ev);
      case 'aportar':      d.evento = ev; return registrarAportacion(d);
      case 'login':        return adminLogin(pin, ev);
      case 'admindatos':   return adminDatos(pin, ev);
      case 'guardar':      return adminGuardarConfig(pin, d, ev);
      case 'estado':       return adminCambiarEstado(pin, p.id || '', p.estado || '', ev);
      case 'eliminar':     return adminEliminar(pin, p.id || '', ev);
      case 'manual':       return adminAportacionManual(pin, d, ev);
      case 'crearpanel':   return crearPanelWeb(pin);
      case 'enviarpanel':  return enviarPanelPorEmail(pin, p.ev2 || '', p.email || '', p.linkpub || '', p.linkadm || '');
      case 'interes':      return registrarInteresEvento(p.email || '');
      case 'solicitarpin': return solicitarPinEvento(p.email || '');
      default:             return { ok: false, msg: 'Acción no válida' };
    }
  } catch (err) {
    return { ok: false, msg: 'Error: ' + err.message };
  }
}

/* ---------- LECTURA CONFIG ---------- */
function leerConfig(evento) {
  if (evento) return _leerConfigEvento(evento);
  const ss = _ss();
  const sh = ss.getSheetByName(H_CONFIG);
  if (!sh) { inicializar(); return leerConfig(); }
  const n = sh.getLastRow() - 1;
  const obj = {};
  if (n > 0) {
    sh.getRange(2, 1, n, 2).getValues().forEach(function (r) {
      if (r[0]) obj[String(r[0]).trim()] = String(r[1]);
    });
  }
  return obj;
}

function _num(v, def) {
  const n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? def : n;
}

function _fotoUrl(u) {
  u = String(u || '').trim();
  if (!u) return '';
  const m = u.match(/[-\w]{25,}/);
  if (u.indexOf('drive.google.com') > -1 && m) {
    return 'https://drive.google.com/thumbnail?id=' + m[0] + '&sz=w1200';
  }
  return u;
}

/* =========================================================
   FIX DEL BUG DE LOS ENLACES
   ---------------------------------------------------------
   Antes: si la lista 'regaloEnlaces' quedaba vacía (porque el
   admin borró todos los enlaces), el código caía al "plan B" y
   releía regaloEnlace / regaloEnlace2 / 3 / 4, que seguían
   guardados en la hoja. Resultado: los enlaces borrados volvían
   a aparecer una y otra vez.

   Ahora: si la clave 'regaloEnlaces' existe, esa lista MANDA,
   aunque esté vacía. El "plan B" solo se usa en hojas antiguas
   que todavía no tienen esa clave.
   ========================================================= */
function _parseEnlaces(cfg) {
  const tieneLista = cfg && Object.prototype.hasOwnProperty.call(cfg, 'regaloEnlaces')
                     && String(cfg.regaloEnlaces).trim() !== '';
  if (tieneLista) {
    let p = [];
    try { p = JSON.parse(cfg.regaloEnlaces); } catch (e) { p = []; }
    if (!Array.isArray(p)) p = [];
    return p.filter(function (u) { return u && String(u).trim(); })
            .map(function (u) { return String(u).trim(); });
  }
  const arr = [];
  ['regaloEnlace', 'regaloEnlace2', 'regaloEnlace3', 'regaloEnlace4'].forEach(function (k) {
    if (cfg[k]) arr.push(String(cfg[k]).trim());
  });
  return arr.filter(function (u) { return u; });
}

/* ---------- DATOS PÚBLICOS ---------- */
function getDatosPublicos(force, evento) {
  const cache = CacheService.getScriptCache();
  const cacheKey = evento ? (CACHE_KEY + '_EV_' + evento) : CACHE_KEY;
  if (!force) {
    const hit = cache.get(cacheKey);
    if (hit) return JSON.parse(hit);
  }

  const cfg = leerConfig(evento);
  if (evento && !cfg) return { cfg: {}, total: 0, objetivo: 0, porcentaje: 0, participantes: 0, pendientes: 0, aportaciones: [], noExiste: true };

  const sh = evento ? _ensureEventoSheets(_ss()).ae : _ss().getSheetByName(H_APORTES);
  const n = sh ? sh.getLastRow() - 1 : 0;
  let total = 0, pendientes = 0;
  const lista = [];

  if (n > 0) {
    const anchoFila = evento ? 8 : 7;
    sh.getRange(2, 1, n, anchoFila).getValues().forEach(function (r) {
      if (evento && String(r[1]) !== String(evento)) return;
      const off = evento ? 1 : 0;
      const rFecha = r[1 + off], rNombre = r[2 + off], rMensaje = r[3 + off],
            rImporte = r[4 + off], rMetodo = r[5 + off], rEstado = r[6 + off];
      const estado = String(rEstado || '').toUpperCase();
      const imp = _num(rImporte, 0);
      if (estado === 'CONFIRMADO') {
        total += imp;
        lista.push({
          id: r[0],
          fecha: rFecha ? Utilities.formatDate(new Date(rFecha), Session.getScriptTimeZone(), 'dd/MM HH:mm') : '',
          nombre: rNombre, mensaje: rMensaje, importe: imp,
          metodo: String(rMetodo || '').toUpperCase()
        });
      } else if (estado === 'PENDIENTE') pendientes++;
    });
  }
  lista.reverse();

  const objetivo = _num(cfg.objetivo, 0);
  const out = {
    cfg: {
      titulo: cfg.titulo, subtitulo: cfg.subtitulo, nombreHomenajeado: cfg.nombreHomenajeado,
      objetivo: objetivo, importeMinimo: _num(cfg.importeMinimo, 5),
      bizumNumero: cfg.bizumNumero, bizumTitular: cfg.bizumTitular,
      myposUsuario: cfg.myposUsuario || 'futurmovil',
      tieneMyPOS: evento ? 'NO' : ((cfg.tieneMyPOS || 'SI').toUpperCase()),
      regaloNombre: cfg.regaloNombre, regaloDescripcion: cfg.regaloDescripcion,
      regaloFoto: _fotoUrl(cfg.regaloFoto), regaloEnlaces: _parseEnlaces(cfg),
      fechaLimite: cfg.fechaLimite, moderacion: (cfg.moderacion || 'NO').toUpperCase(),
      mostrarImportes: (cfg.mostrarImportes || 'SI').toUpperCase(),
      permitirAnonimo: (cfg.permitirAnonimo || 'SI').toUpperCase(),
      mensajeGracias: cfg.mensajeGracias,
      activo: (cfg.activo || 'SI').toUpperCase(),
      mensajeCerrado: cfg.mensajeCerrado,
      tema: cfg.tema || 'default'
    },
    total: Math.round(total * 100) / 100,
    objetivo: objetivo,
    porcentaje: objetivo > 0 ? Math.min(100, Math.round((total / objetivo) * 1000) / 10) : 0,
    participantes: lista.length,
    pendientes: pendientes,
    aportaciones: lista.slice(0, 200)
  };

  cache.put(cacheKey, JSON.stringify(out), CACHE_SEG);
  return out;
}

function _limpiarCache(evento) {
  const c = CacheService.getScriptCache();
  c.remove(CACHE_KEY);
  if (evento) c.remove(CACHE_KEY + '_EV_' + evento);
}

/* ---------- REGISTRAR APORTACIÓN ---------- */
function registrarAportacion(d) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok: false, msg: 'Servidor ocupado, inténtalo de nuevo.' }; }

  try {
    const evento = String(d.evento || '').trim();
    const cfg = leerConfig(evento || undefined);
    if (evento && !cfg) return { ok: false, msg: 'Evento no encontrado.' };
    if ((cfg.activo || 'SI').toUpperCase() !== 'SI') return { ok: false, msg: 'La recaudación está cerrada.' };

    const min = _num(cfg.importeMinimo, 5);
    const imp = _num(d.importe, 0);
    if (imp < min) return { ok: false, msg: 'El importe mínimo es ' + min + ' €.' };
    if (imp > 5000) return { ok: false, msg: 'Importe demasiado alto.' };

    let nombre = String(d.nombre || '').trim().slice(0, 40);
    if (!nombre) {
      if ((cfg.permitirAnonimo || 'SI').toUpperCase() === 'SI') nombre = 'Anónimo';
      else return { ok: false, msg: 'Escribe tu nombre.' };
    }
    const mensaje = String(d.mensaje || '').trim();
    let metodo = (String(d.metodo || '').toUpperCase() === 'TARJETA') ? 'TARJETA' : 'BIZUM';
    if (evento) metodo = (metodo === 'TARJETA') ? 'BIZUM' : metodo;
    const estado = 'PENDIENTE';

    const id = Utilities.getUuid();
    if (evento) {
      _ensureEventoSheets(_ss()).ae.appendRow([id, evento, new Date(), nombre, mensaje, imp, metodo, estado]);
    } else {
      _ss().getSheetByName(H_APORTES).appendRow([id, new Date(), nombre, mensaje, imp, metodo, estado]);
    }

    _limpiarCache(evento);
    try {
      const emailAdmin = evento ? (cfg._email || '') : ADMIN_EMAIL;
      if (emailAdmin) {
        const linkAdminNotif = urlAdmin(evento);
        MailApp.sendEmail({
          to: emailAdmin,
          subject: '🔔 Nueva aportación pendiente de validar (' + imp + ' €)',
          htmlBody:
            '<p>¡Hola!</p>' +
            '<p><b>' + nombre + '</b> ha hecho una aportación de <b>' + imp + ' €</b> por ' + metodo + ' que está pendiente de que la valides.</p>' +
            (mensaje ? ('<p>Mensaje: "' + mensaje + '"</p>') : '') +
            '<p>Entra en tu panel para aceptarla y que se publique en el muro:</p>' +
            '<p><a href="' + linkAdminNotif + '">' + linkAdminNotif + '</a></p>'
        });
      }
    } catch (eMail) { Logger.log('ERROR_NOTIF_ADMIN ' + String(eMail)); }

    return {
      ok: true, estado: estado,
      msg: 'Pendiente de verificar tu pago por el administrador 🔎. En cuanto se compruebe, tu aportación se publicará en el muro con tu nombre. ¡Gracias por tu generosidad! 💜'
    };
  } catch (err) {
    return { ok: false, msg: 'Error: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- ADMIN ---------- */
function _pinOk(pin, evento) {
  const cfg = leerConfig(evento || undefined);
  if (!cfg) return false;
  const coincide = String(pin || '').trim() === String(cfg.pin || '').trim() && String(cfg.pin || '').trim() !== '';
  if (!coincide) return false;
  if (evento && _eventoCaducado(cfg)) return false;
  return true;
}

function adminLogin(pin, evento) {
  evento = String(evento || '').trim();
  if (evento) {
    const cfg = leerConfig(evento);
    const coincide = cfg && String(pin || '').trim() === String(cfg.pin || '').trim() && String(cfg.pin || '').trim() !== '';
    if (!coincide) { Utilities.sleep(700); return { ok: false, msg: 'PIN incorrecto' }; }
    if (_eventoCaducado(cfg)) return { ok: false, msg: 'Tu acceso caducó (30 días desde tu primer uso).' };
    _marcarPrimerUso(evento);
    return { ok: true };
  }
  if (!_pinOk(pin, evento)) { Utilities.sleep(700); return { ok: false, msg: 'PIN incorrecto' }; }
  return { ok: true };
}

function adminDatos(pin, evento) {
  if (!_pinOk(pin, evento)) return { ok: false, msg: 'Sesión no válida' };
  evento = String(evento || '').trim();

  const cfg = leerConfig(evento || undefined);
  cfg.objetivo = _num(cfg.objetivo, 0);
  cfg.importeMinimo = _num(cfg.importeMinimo, 5);
  cfg.tieneMyPOS = evento ? 'NO' : ((cfg.tieneMyPOS || 'SI').toUpperCase());
  cfg.regaloEnlacesArr = _parseEnlaces(cfg);

  const sh = evento ? _ensureEventoSheets(_ss()).ae : _ss().getSheetByName(H_APORTES);
  const n = sh ? sh.getLastRow() - 1 : 0;
  const filas = [];
  let confirmado = 0, pendiente = 0, bizum = 0, tarjeta = 0;
  const anchoFila = evento ? 8 : 7;
  const off = evento ? 1 : 0;

  if (n > 0) {
    sh.getRange(2, 1, n, anchoFila).getValues().forEach(function (r) {
      if (evento && String(r[1]) !== String(evento)) return;
      const rFecha = r[1 + off], rNombre = r[2 + off], rMensaje = r[3 + off],
            rImporte = r[4 + off], rMetodo = r[5 + off], rEstado = r[6 + off];
      const est = String(rEstado || '').toUpperCase();
      const imp = _num(rImporte, 0);
      if (est === 'CONFIRMADO') {
        confirmado += imp;
        if (String(rMetodo).toUpperCase() === 'BIZUM') bizum += imp;
        else if (String(rMetodo).toUpperCase() === 'TARJETA') tarjeta += imp;
      }
      if (est === 'PENDIENTE') pendiente += imp;
      filas.push({
        id: r[0],
        fecha: rFecha ? Utilities.formatDate(new Date(rFecha), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
        nombre: rNombre, mensaje: rMensaje, importe: imp,
        metodo: String(rMetodo || '').toUpperCase(), estado: est
      });
    });
  }
  filas.reverse();

  return {
    ok: true, cfg: cfg, filas: filas,
    stats: {
      confirmado: Math.round(confirmado * 100) / 100,
      pendiente: Math.round(pendiente * 100) / 100,
      bizum: Math.round(bizum * 100) / 100,
      tarjeta: Math.round(tarjeta * 100) / 100,
      num: filas.length,
      porcentaje: _num(cfg.objetivo, 0) > 0 ? Math.min(100, Math.round(confirmado / _num(cfg.objetivo, 1) * 1000) / 10) : 0
    },
    url: urlPublica(evento)
  };
}

function adminGuardarConfig(pin, datos, evento) {
  if (!_pinOk(pin, evento)) return { ok: false, msg: 'Sesión no válida' };
  evento = String(evento || '').trim();

  if (evento) {
    datos = Object.assign({}, datos);
    const nuevoPin = String(datos.pin || '').trim();
    delete datos.pin;
    const nuevoEmail = String(datos._email || '').trim();
    delete datos._email;
    _guardarConfigEvento(evento, datos);
    if (nuevoPin.length >= 4) {
      const shEv = _ensureEventoSheets(_ss()).ev;
      const row = _findRowById(shEv, evento);
      if (row > 0) shEv.getRange(row, 3).setValue(nuevoPin);
    }
    if (nuevoEmail) {
      const shEv2 = _ensureEventoSheets(_ss()).ev;
      const row2 = _findRowById(shEv2, evento);
      if (row2 > 0) shEv2.getRange(row2, 2).setValue(nuevoEmail);
    }
    _limpiarCache(evento);
    return { ok: true, msg: 'Configuración guardada ✓' };
  }

  datos = Object.assign({}, datos);

  // === FIX ENLACES (campaña principal) ===
  // Si llega la lista nueva, vaciamos las claves antiguas para que
  // un enlace borrado no resucite al recargar el panel.
  if (Object.prototype.hasOwnProperty.call(datos, 'regaloEnlaces')) {
    datos.regaloEnlace  = '';
    datos.regaloEnlace2 = '';
    datos.regaloEnlace3 = '';
    datos.regaloEnlace4 = '';
  }

  const sh = _ss().getSheetByName(H_CONFIG);
  const n = sh.getLastRow() - 1;
  const rango = sh.getRange(2, 1, n, 2);
  const vals = rango.getValues();

  Object.keys(datos).forEach(function (k) {
    let encontrado = false;
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === k) { vals[i][1] = datos[k]; encontrado = true; break; }
    }
    if (!encontrado) sh.appendRow([k, datos[k]]);
  });

  rango.setValues(vals);
  _limpiarCache();
  return { ok: true, msg: 'Configuración guardada ✓' };
}

function adminCambiarPinEvento(pin, evento, nuevoPin) {
  if (!_pinOk(pin, evento)) return { ok: false, msg: 'Sesión no válida' };
  evento = String(evento || '').trim();
  nuevoPin = String(nuevoPin || '').trim();
  if (!evento) return { ok: false, msg: 'Esta acción solo está disponible en tu evento.' };
  if (nuevoPin.length < 4) return { ok: false, msg: 'El PIN debe tener al menos 4 dígitos.' };
  const sh = _ensureEventoSheets(_ss()).ev;
  const row = _findRowById(sh, evento);
  if (row < 0) return { ok: false, msg: 'Evento no encontrado' };
  sh.getRange(row, 3).setValue(nuevoPin);
  return { ok: true, msg: 'PIN actualizado ✓' };
}

function adminCambiarEstado(pin, id, estado, evento) {
  if (!_pinOk(pin, evento)) return { ok: false, msg: 'Sesión no válida' };
  evento = String(evento || '').trim();
  estado = String(estado).toUpperCase();
  const sh = evento ? _ensureEventoSheets(_ss()).ae : _ss().getSheetByName(H_APORTES);
  const n = sh.getLastRow() - 1;
  if (n < 1) return { ok: false, msg: 'Sin datos' };
  const ids = sh.getRange(2, 1, n, 1).getValues();
  const colEstado = evento ? 8 : 7;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sh.getRange(i + 2, colEstado).setValue(estado);
      _limpiarCache(evento);
      return { ok: true, msg: 'Estado actualizado' };
    }
  }
  return { ok: false, msg: 'No encontrado' };
}

function adminEliminar(pin, id, evento) {
  if (!_pinOk(pin, evento)) return { ok: false, msg: 'Sesión no válida' };
  evento = String(evento || '').trim();
  const sh = evento ? _ensureEventoSheets(_ss()).ae : _ss().getSheetByName(H_APORTES);
  const n = sh.getLastRow() - 1;
  if (n < 1) return { ok: false, msg: 'Sin datos' };
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sh.deleteRow(i + 2);
      _limpiarCache(evento);
      return { ok: true, msg: 'Aportación eliminada' };
    }
  }
  return { ok: false, msg: 'No encontrado' };
}

function adminAportacionManual(pin, d, evento) {
  if (!_pinOk(pin, evento)) return { ok: false, msg: 'Sesión no válida' };
  evento = String(evento || '').trim();
  let metodo = (String(d.metodo).toUpperCase() === 'TARJETA' ? 'TARJETA'
              : (String(d.metodo).toUpperCase() === 'EFECTIVO' ? 'EFECTIVO' : 'BIZUM'));
  if (evento && metodo === 'TARJETA') metodo = 'BIZUM';
  if (evento) {
    _ensureEventoSheets(_ss()).ae.appendRow([
      Utilities.getUuid(), evento, new Date(),
      String(d.nombre || 'Anónimo').slice(0, 40), String(d.mensaje || ''),
      _num(d.importe, 0), metodo, 'CONFIRMADO'
    ]);
    _limpiarCache(evento);
  } else {
    _ss().getSheetByName(H_APORTES).appendRow([
      Utilities.getUuid(), new Date(),
      String(d.nombre || 'Anónimo').slice(0, 40), String(d.mensaje || ''),
      _num(d.importe, 0), metodo, 'CONFIRMADO'
    ]);
    _limpiarCache();
  }
  return { ok: true, msg: 'Aportación añadida ✓' };
}

/* ---------- UTILIDAD DE LIMPIEZA (opcional, ejecutar 1 vez) ----------
   Borra de la hoja CONFIG los enlaces antiguos que causaban el bug.
   No es obligatorio: el código nuevo ya los ignora.
--------------------------------------------------------------------- */
function limpiarEnlacesAntiguos() {
  const sh = _ss().getSheetByName(H_CONFIG);
  const n = sh.getLastRow() - 1;
  const vals = sh.getRange(2, 1, n, 2).getValues();
  const viejas = ['regaloEnlace', 'regaloEnlace2', 'regaloEnlace3', 'regaloEnlace4'];
  vals.forEach(function (r, i) {
    if (viejas.indexOf(String(r[0]).trim()) > -1) vals[i][1] = '';
  });
  sh.getRange(2, 1, n, 2).setValues(vals);
  _limpiarCache();
  return 'Enlaces antiguos limpiados';
}

function _forceAuth() {
  const dest = Session.getActiveUser().getEmail() || ADMIN_EMAIL;
  MailApp.sendEmail(dest, 'Prueba autorizacion REGALO PARTY', 'Correo de prueba para activar el permiso de envio de correos.');
  return 'sent to ' + dest;
}
