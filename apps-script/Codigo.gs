// ============================================================
// DEUDAS PESOS (COP) — con interes mensual opcional
// ============================================================
// Modelo de interes: "por mes cumplido", interes simple.
//  - El mismo dia de la deuda ya se suma el X% del capital (el primer
//    mes de interes se cobra por adelantado, aunque pague ese dia).
//  - Cada vez que se cumple un mes desde la fecha de la deuda, se suma
//    otro X% del CAPITAL que quede pendiente en ese momento.
//  - Los pagos (parciales o totales) cubren primero los intereses
//    pendientes y despues el capital.
//  - El interes no genera interes. Todo se redondea a pesos enteros.
//  - Todo se recalcula desde la fecha de la deuda y la lista de pagos,
//    asi que el total siempre cuadra aunque se toque la hoja a mano.
// ============================================================
const MASTER_USER = 'Manu';
const CORREO_MASTER = 'futurmovil.com@gmail.com';
const NOMBRE_APP = 'DeudasPesos';
// Enlace publico de la app instalable (GitHub Pages). Es el que va en los correos
// para que al anadirla a la pantalla de inicio (sobre todo en iPhone) salga el icono.
const URL_APP = 'https://puntofibra.github.io/deudaspesos/';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle(NOMBRE_APP);
}

// ============================================================
// INIT
// ============================================================
const CAB_USUARIOS = ['id', 'usuario', 'contrasena', 'creadoPor', 'fechaCreacion', 'correo', 'ocultoPorCreador', 'aliasCreador', 'licenciaHasta'];
const CAB_DEUDAS = ['id', 'acreedorId', 'deudorId', 'descripcion', 'importe', 'fecha', 'estado', 'fechaPago', 'nota', 'interesMensual'];
const CAB_PAGOS = ['id', 'idDeuda', 'monto', 'fecha', 'aInteres', 'aCapital'];

function initApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['usuarios', 'deudas', 'pagos_deuda', 'omitidos', 'vinculos', 'solicitudes_pago', 'claves master'].forEach(nombre => {
    if (!ss.getSheetByName(nombre)) ss.insertSheet(nombre);
  });

  const sheetUsuarios = ss.getSheetByName('usuarios');
  if (sheetUsuarios.getLastRow() === 0) {
    sheetUsuarios.appendRow(CAB_USUARIOS);
    sheetUsuarios.appendRow([Utilities.getUuid(), MASTER_USER, Utilities.getUuid().toLowerCase().trim(), '', new Date().toISOString(), CORREO_MASTER, false, MASTER_USER, '']);
  }

  const sheetDeudas = ss.getSheetByName('deudas');
  if (sheetDeudas.getLastRow() === 0) sheetDeudas.appendRow(CAB_DEUDAS);
  else asegurarCabecera(sheetDeudas, CAB_DEUDAS);

  const sheetPagos = ss.getSheetByName('pagos_deuda');
  if (sheetPagos.getLastRow() === 0) sheetPagos.appendRow(CAB_PAGOS);
  else asegurarCabecera(sheetPagos, CAB_PAGOS);

  const sheetOmitidos = ss.getSheetByName('omitidos');
  if (sheetOmitidos.getLastRow() === 0) sheetOmitidos.appendRow(['usuarioId', 'omitidoId', 'fecha']);

  const sheetVinculos = ss.getSheetByName('vinculos');
  if (sheetVinculos.getLastRow() === 0) {
    sheetVinculos.appendRow(['usuarioId1', 'usuarioId2', 'fecha', 'alias1', 'alias2', 'oculto1', 'oculto2']);
  }

  const sheetSolicitudes = ss.getSheetByName('solicitudes_pago');
  if (sheetSolicitudes.getLastRow() === 0) {
    sheetSolicitudes.appendRow(['id', 'idDeuda', 'deudorId', 'monto', 'tipo', 'estado', 'fecha', 'fechaRespuesta']);
  }
}

// Si faltan columnas al final de la cabecera, las anade (no toca datos).
function asegurarCabecera(sheet, cab) {
  const n = sheet.getLastColumn();
  if (n >= cab.length) return;
  sheet.getRange(1, 1, 1, cab.length).setValues([cab]);
}

// ============================================================
// UTILIDADES DE PESOS / FECHAS
// ============================================================
function pesos(x) { return Math.round(Number(x) || 0); }

function aFecha(v) {
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function esVerdad(v) { return v === true || String(v).toLowerCase().trim() === 'true'; }

// Suma n meses a una fecha manteniendo el dia (31 ene + 1 mes = 28/29 feb).
function sumarMeses(base, n) {
  const d = new Date(base.getTime());
  const dia = d.getDate();
  d.setDate(1);
  d.setMonth(d.getMonth() + n);
  const ultimo = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  d.setDate(Math.min(dia, ultimo));
  return d;
}

function fmtCOP(n) {
  const v = pesos(n);
  return '$' + String(Math.abs(v)).replace(/\B(?=(\d{3})+(?!\d))/g, '.') ;
}

// ============================================================
// MOTOR DE INTERESES
// ============================================================
// importe: capital inicial, tasaPct: % mensual (0 = sin interes),
// fechaInicio: fecha de la deuda, pagos: [{id, monto, fecha}],
// hasta: fecha hasta la que se devengan intereses (null = ahora).
function calcularDeuda(importe, tasaPct, fechaInicio, pagos, hasta) {
  const capitalInicial = pesos(importe);
  const tasaNum = Math.max(0, Number(tasaPct) || 0);
  const tasa = tasaNum / 100;
  const inicio = aFecha(fechaInicio) || new Date();
  const limite = aFecha(hasta) || new Date();

  let capPend = capitalInicial, intPend = 0, intGen = 0, intPag = 0, capPag = 0;
  // El primer interes se cobra el mismo dia de la deuda (mes 1),
  // y despues uno mas cada vez que se cumple un mes.
  let n = 0;
  let corte = inicio;
  const movs = [];

  function devengarHasta(f) {
    if (tasa <= 0) return;
    let guard = 0;
    while (capPend > 0 && (n === 0 || corte.getTime() <= f.getTime()) && guard < 1200) {
      const i = pesos(capPend * tasa);
      if (i > 0) {
        intPend += i; intGen += i;
        movs.push({ tipo: 'interes', fecha: corte.toISOString(), monto: i, mes: n + 1, sobre: capPend });
      }
      n++; guard++;
      corte = sumarMeses(inicio, n);
    }
  }

  const ordenados = (pagos || []).slice().sort((a, b) => aFecha(a.fecha) - aFecha(b.fecha));
  ordenados.forEach(p => {
    const f = aFecha(p.fecha) || new Date();
    devengarHasta(f);
    const m = pesos(p.monto);
    const aInt = Math.min(m, intPend);
    const aCap = Math.min(m - aInt, capPend);
    intPend -= aInt; intPag += aInt;
    capPend -= aCap; capPag += aCap;
    movs.push({ tipo: 'pago', id: p.id, fecha: f.toISOString(), monto: m, aInteres: aInt, aCapital: aCap });
  });

  devengarHasta(limite);

  return {
    capitalInicial: capitalInicial,
    tasa: tasaNum,
    capitalPendiente: capPend,
    interesPendiente: intPend,
    interesGenerado: intGen,
    interesPagado: intPag,
    capitalPagado: capPag,
    totalPagado: intPag + capPag,
    restante: capPend + intPend,
    mesesCumplidos: Math.max(0, n - 1),
    proximoCorte: (tasa > 0 && capPend > 0) ? corte.toISOString() : null,
    interesProximo: (tasa > 0 && capPend > 0) ? pesos(capPend * tasa) : 0,
    movimientos: movs
  };
}

// Estado calculado de una fila de la hoja 'deudas'
function estadoDeFila(fila, pagosData) {
  const id = String(fila[0]);
  const estado = String(fila[6]);
  const hasta = estado === 'activo' ? null : (aFecha(fila[7]) || null);
  return calcularDeuda(fila[4], fila[9], fila[5], pagosDeDeuda(id, pagosData), hasta);
}

function buscarFilaDeuda(idDeuda) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('deudas');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idDeuda)) return { sheet: sheet, index: i, fila: data[i] };
  }
  return null;
}

// ============================================================
// AUTH
// ============================================================
function usuariosSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('usuarios');
}
function usuariosData() {
  return usuariosSheet().getDataRange().getValues();
}

function obtenerUsuarioPorId(id, data) {
  const rows = data || usuariosData();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      return { id: String(rows[i][0]), usuario: String(rows[i][1]), contrasena: String(rows[i][2]), creadoPor: String(rows[i][3]), correo: String(rows[i][5] || '') };
    }
  }
  return null;
}

function esCuentaMaster(usuarioPublico, correo) {
  const u = String(usuarioPublico || '').toLowerCase().trim();
  const c = String(correo || '').toLowerCase().trim();
  return u === String(MASTER_USER).toLowerCase().trim() || (c !== '' && c === CORREO_MASTER);
}

function sesionDeFila(fila) {
  return { ok: true, id: String(fila[0]), usuario: String(fila[1]), correo: String(fila[5] || ''), esMaster: esCuentaMaster(String(fila[1]), String(fila[5] || '')) };
}

function _loginInterno(usuario, contrasena) {
  initApp();
  const data = usuariosData();
  const u = String(usuario).toLowerCase().trim();
  const p = String(contrasena).toLowerCase().trim();
  for (let i = 1; i < data.length; i++) {
    const pass = String(data[i][2]).toLowerCase().trim();
    if (pass !== p) continue;
    const nombrePublico = String(data[i][1]).toLowerCase().trim();
    const aliasCreador = String(data[i][7] || '').toLowerCase().trim();
    const id = String(data[i][0]).toLowerCase().trim();
    const correo = String(data[i][5] || '').toLowerCase().trim();
    if (u === nombrePublico || (aliasCreador && u === aliasCreador) || u === id || (correo && u === correo)) {
      return sesionDeFila(data[i]);
    }
  }
  return { ok: false };
}

// Login directo por nombre de usuario (solo lo usa la llave NFC, en servidor)
function _loginPorNombre(usuario) {
  initApp();
  const data = usuariosData();
  const u = String(usuario).toLowerCase().trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase().trim() === u) return sesionDeFila(data[i]);
  }
  return { ok: false };
}

function cambiarContrasena(id, actual, nueva) {
  const sheet = usuariosSheet();
  const data = sheet.getDataRange().getValues();
  const a = String(actual).toLowerCase().trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      if (String(data[i][2]).toLowerCase().trim() !== a) return { ok: false, msg: 'Contrasena actual incorrecta' };
      sheet.getRange(i + 1, 3).setValue(String(nueva).toLowerCase().trim());
      return { ok: true };
    }
  }
  return { ok: false, msg: 'Usuario no encontrado' };
}

// ============================================================
// CLAVES MAESTRAS (pestana "claves master", columna A)
// ============================================================
function clavesMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('claves master');
  if (!sh) sh = ss.insertSheet('claves master');
  return sh;
}

function listaClavesMaster() {
  const sh = clavesMasterSheet();
  const last = sh.getLastRow();
  if (last < 1) return [];
  return sh.getRange(1, 1, last, 1).getValues()
    .map(r => r[0])
    .filter(v => v !== null && v !== undefined && String(v).trim() !== '')
    .map(v => String(v).trim());
}

function esClaveMasterValida(clave) {
  if (!clave) return false;
  const c = String(clave).trim();
  return listaClavesMaster().indexOf(c) !== -1;
}

function _passwordMasterAlmacenada() {
  const data = usuariosData();
  const mu = String(MASTER_USER).toLowerCase().trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase().trim() === mu) return String(data[i][2]);
  }
  return null;
}

// El master entra con cualquiera de las claves de la pestana "claves master".
function login(usuario, clave) {
  initApp();
  const u = String(usuario == null ? '' : usuario).trim();
  if (u.toLowerCase() === String(MASTER_USER).toLowerCase().trim()) {
    if (esClaveMasterValida(clave)) {
      const real = _passwordMasterAlmacenada();
      if (real !== null) return _loginInterno(MASTER_USER, real);
    }
    return { ok: false, mensaje: 'Usuario o contrasena incorrectos' };
  }
  return _loginInterno(usuario, clave);
}

// ============================================================
// USUARIOS
// ============================================================
function crearUsuario(usuarioNuevo, correoNuevo, creadorId) {
  const nombre = String(usuarioNuevo || '').trim();
  const correo = String(correoNuevo || '').trim().toLowerCase();
  if (!nombre) return { ok: false, msg: 'Escribe un nombre de usuario' };
  if (!correo || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) {
    return { ok: false, msg: 'Escribe un correo electronico valido' };
  }
  const creador = obtenerUsuarioPorId(creadorId);
  if (!creador) return { ok: false, msg: 'Tu sesion no es valida, vuelve a iniciar sesion' };

  const data = usuariosData();
  for (let i = 1; i < data.length; i++) {
    const idFila = String(data[i][0]).toLowerCase().trim();
    const correoFila = String(data[i][5] || '').toLowerCase().trim();
    if (idFila === correo || correoFila === correo) {
      return { ok: false, msg: 'Ya existe un usuario con ese correo electronico' };
    }
  }

  const passInicial = String(Math.floor(100000 + Math.random() * 900000));
  usuariosSheet().appendRow([correo, nombre, passInicial, String(creadorId), new Date().toISOString(), correo, false, nombre, '']);

  let avisoCorreo = '';
  try {
    enviarInvitacion(correo, nombre, creador.usuario, passInicial);
  } catch (err) {
    avisoCorreo = 'Usuario creado, pero no se pudo enviar el correo de invitacion.';
  }
  return { ok: true, msg: avisoCorreo || ('Invitacion enviada a ' + correo) };
}

function plantillaCorreo(titulo, cuerpo, url) {
  return '<div style="font-family:Arial,sans-serif;max-width:480px;margin:auto;border:1px solid #e3ddd0;border-radius:14px;overflow:hidden">' +
    '<div style="background:#0e5a47;color:#fff;padding:16px 20px;font-size:18px;font-weight:bold">&#128181; ' + NOMBRE_APP + '</div>' +
    '<div style="padding:18px 20px;color:#1c2a24">' +
    '<h2 style="color:#0e7a5f;margin-top:0">' + titulo + '</h2>' + cuerpo +
    (url ? '<p><a href="' + url + '" style="background:#0e7a5f;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none">Entrar a ' + NOMBRE_APP + '</a></p>' : '') +
    '</div></div>';
}

function enviarInvitacion(correo, nombre, nombreCreador, passInicial) {
  const url = URL_APP; // la PWA de GitHub Pages (con icono), no la /exec
  const cuerpo =
    '<p><b>' + escaparHtml(nombreCreador) + '</b> te ha invitado a <b>' + NOMBRE_APP + '</b>, ' +
    'una forma sencilla de llevar el control de las deudas en pesos entre ambos: quien le debe a quien, ' +
    'los intereses (si los hay) y los pagos totales o por cuotas.</p>' +
    '<div style="background:#f4f1ea;border-radius:10px;padding:14px;margin:14px 0">' +
    '<p style="margin:0"><b>Usuario:</b> ' + escaparHtml(nombre) + '</p>' +
    '<p style="margin:0"><b>Contrasena inicial:</b> ' + passInicial + '</p>' +
    '</div>' +
    '<p>Por seguridad, cambia la contrasena despues de tu primer ingreso desde "Mi perfil".</p>';
  MailApp.sendEmail({
    to: correo,
    subject: nombreCreador + ' te ha invitado a ' + NOMBRE_APP,
    htmlBody: plantillaCorreo('Hola, ' + escaparHtml(nombre) + '!', cuerpo, url),
    body: nombreCreador + ' te ha invitado a ' + NOMBRE_APP + '. Usuario: ' + nombre + ' / Contrasena inicial: ' + passInicial + (url ? (' / Enlace: ' + url) : '')
  });
}

function escaparHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function listarMisUsuarios(idSolicitante) {
  const data = usuariosData();
  const dataV = vinculosSheet();
  const yo = obtenerUsuarioPorId(idSolicitante, data);
  if (!yo) return { creador: null, creados: [] };
  const v = String(idSolicitante);

  let creador = null;
  let miFilaOculta = false;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === v) { miFilaOculta = esVerdad(data[i][6]); break; }
  }
  if (yo.creadoPor) {
    if (miFilaOculta) {
      creador = { id: null, usuario: 'xxx', oculto: true };
    } else {
      const c = obtenerUsuarioPorId(yo.creadoPor, data);
      if (c) creador = { id: c.id, usuario: c.usuario };
    }
  }

  const creados = [];
  const yaAgregados = {};
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]) === v) {
      if (esVerdad(data[i][6])) continue;
      const id = String(data[i][0]);
      creados.push({ id: id, usuario: String(data[i][7] || data[i][1]), correo: String(data[i][5] || ''), tipo: 'creado' });
      yaAgregados[id] = true;
    }
  }

  for (let i = 1; i < dataV.length; i++) {
    const x = String(dataV[i][0]), y = String(dataV[i][1]);
    const oculto = esVerdad(dataV[i][5]) || esVerdad(dataV[i][6]);
    let otroId = null, alias = '';
    if (x === v) { otroId = y; alias = String(dataV[i][3] || ''); }
    else if (y === v) { otroId = x; alias = String(dataV[i][4] || ''); }
    if (!otroId || oculto || yaAgregados[otroId]) continue;
    const infoUsr = obtenerUsuarioPorId(otroId, data);
    if (!infoUsr) continue;
    if (creador && String(creador.id) === otroId) continue;
    creados.push({ id: otroId, usuario: alias.trim() !== '' ? alias : infoUsr.usuario, correo: '', tipo: 'vinculado' });
    yaAgregados[otroId] = true;
  }
  return { creador: creador, creados: creados };
}

function eliminarUsuario(idUsuario, idSolicitante) {
  const sheet = usuariosSheet();
  const data = sheet.getDataRange().getValues();
  let filaIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idUsuario)) { filaIndex = i; break; }
  }
  if (filaIndex === -1) return { ok: false, msg: 'Usuario no encontrado' };
  if (String(data[filaIndex][3]) !== String(idSolicitante)) {
    return { ok: false, msg: 'Solo puedes eliminar usuarios que hayas creado tu' };
  }
  if (existeDeudaActivaEntre(idSolicitante, idUsuario)) {
    return { ok: false, msg: 'No se puede eliminar: hay deudas pendientes entre ustedes' };
  }
  sheet.getRange(filaIndex + 1, 7).setValue(true);
  return { ok: true, msg: 'Usuario quitado de tu lista. Podras recuperarlo despues con su correo.' };
}

function recuperarUsuarioEliminado(correoBuscado, idSolicitante) {
  const correo = String(correoBuscado || '').trim().toLowerCase();
  if (!correo) return { ok: false, msg: 'Escribe el correo del usuario a recuperar' };
  const sheet = usuariosSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const idFila = String(data[i][0]).toLowerCase().trim();
    const correoFila = String(data[i][5] || '').toLowerCase().trim();
    if ((idFila === correo || correoFila === correo) && String(data[i][3]) === String(idSolicitante)) {
      if (!esVerdad(data[i][6])) return { ok: false, msg: 'Ese usuario ya esta en tu lista' };
      sheet.getRange(i + 1, 7).setValue(false);
      return { ok: true, msg: 'Usuario recuperado: ' + data[i][1] };
    }
  }
  return { ok: false, msg: 'No se encontro ningun usuario eliminado con ese correo entre los que has creado' };
}

// ============================================================
// VINCULOS
// ============================================================
function vinculosSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('vinculos').getDataRange().getValues();
}

function filaVinculo(viewerId, otroId, data) {
  const rows = data || vinculosSheet();
  const v = String(viewerId), o = String(otroId);
  for (let i = 1; i < rows.length; i++) {
    const x = String(rows[i][0]), y = String(rows[i][1]);
    if (x === v && y === o) return { filaIndex: i, lado: 1, aliasParaOtro: String(rows[i][3] || ''), ocultoParaMi: esVerdad(rows[i][5]) };
    if (x === o && y === v) return { filaIndex: i, lado: 2, aliasParaOtro: String(rows[i][4] || ''), ocultoParaMi: esVerdad(rows[i][6]) };
  }
  return null;
}

function aliasDeVinculo(viewerId, otroId, data) {
  const fv = filaVinculo(viewerId, otroId, data);
  return fv ? fv.aliasParaOtro : '';
}

function editarAliasVinculo(otroId, nuevoNombre, viewerId) {
  const nombre = String(nuevoNombre || '').trim();
  if (!nombre) return { ok: false, msg: 'Escribe un nombre' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('vinculos');
  const data = sheet.getDataRange().getValues();
  const fv = filaVinculo(viewerId, otroId, data);
  if (!fv) return { ok: false, msg: 'No tienes un vinculo con ese usuario' };
  sheet.getRange(fv.filaIndex + 1, fv.lado === 1 ? 4 : 5).setValue(nombre);
  return { ok: true, usuario: nombre };
}

function ocultarVinculo(otroId, viewerId) {
  if (existeDeudaActivaEntre(viewerId, otroId)) {
    return { ok: false, msg: 'No se puede eliminar: hay deudas pendientes entre ustedes' };
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('vinculos');
  const data = sheet.getDataRange().getValues();
  const fv = filaVinculo(viewerId, otroId, data);
  if (!fv) return { ok: false, msg: 'No tienes un vinculo con ese usuario' };
  sheet.getRange(fv.filaIndex + 1, 6).setValue(true);
  sheet.getRange(fv.filaIndex + 1, 7).setValue(true);
  return { ok: true, msg: 'Usuario eliminado. Para volver a verse, alguno tendra que agregarlo de nuevo por correo.' };
}

function idsVinculados(usuarioId) {
  const u = String(usuarioId);
  const set = {};
  const dataU = usuariosData();
  const yo = obtenerUsuarioPorId(u, dataU);
  if (yo && yo.creadoPor) set[String(yo.creadoPor)] = true;
  for (let i = 1; i < dataU.length; i++) {
    if (String(dataU[i][3]) === u && !esVerdad(dataU[i][6])) set[String(dataU[i][0])] = true;
  }
  const dataV = vinculosSheet();
  for (let i = 1; i < dataV.length; i++) {
    const x = String(dataV[i][0]), y = String(dataV[i][1]);
    if (esVerdad(dataV[i][5]) || esVerdad(dataV[i][6])) continue;
    if (x === u) set[y] = true;
    else if (y === u) set[x] = true;
  }
  delete set[u];
  return set;
}

function listarUsuariosVinculados(usuarioId) {
  const vinc = idsVinculados(usuarioId);
  const nombres = mapaNombresParaVisor(usuarioId);
  const out = [];
  const dataU = usuariosData();
  for (let i = 1; i < dataU.length; i++) {
    const id = String(dataU[i][0]);
    if (vinc[id]) out.push({ id: id, usuario: nombres[id] || String(dataU[i][1]) });
  }
  return out;
}

function agregarUsuarioPorCorreo(correoBuscado, idSolicitante) {
  const correo = String(correoBuscado || '').trim().toLowerCase();
  if (!correo) return { ok: false, msg: 'Escribe el correo del usuario a agregar' };
  const yo = obtenerUsuarioPorId(idSolicitante);
  if (!yo) return { ok: false, msg: 'Tu sesion no es valida, vuelve a iniciar sesion' };
  const data = usuariosData();
  let objetivoId = null, objetivoNombre = '';
  for (let i = 1; i < data.length; i++) {
    const idFila = String(data[i][0]).toLowerCase().trim();
    const correoFila = String(data[i][5] || '').toLowerCase().trim();
    if (idFila === correo || correoFila === correo) {
      objetivoId = String(data[i][0]);
      objetivoNombre = String(data[i][1]);
      break;
    }
  }
  if (!objetivoId) return { ok: false, msg: 'No existe ningun usuario con ese correo electronico' };
  if (String(objetivoId) === String(idSolicitante)) return { ok: false, msg: 'Ese es tu propio correo' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('vinculos');
  const dataV = sheet.getDataRange().getValues();
  const fv = filaVinculo(idSolicitante, objetivoId, dataV);
  if (fv) {
    const oculto = esVerdad(dataV[fv.filaIndex][5]) || esVerdad(dataV[fv.filaIndex][6]);
    if (!oculto) return { ok: false, msg: 'Ya tienes vinculado a ' + objetivoNombre };
    sheet.getRange(fv.filaIndex + 1, 6).setValue(false);
    sheet.getRange(fv.filaIndex + 1, 7).setValue(false);
    return { ok: true, msg: 'Ahora estas vinculado con ' + objetivoNombre };
  }
  const vinc = idsVinculados(idSolicitante);
  if (vinc[objetivoId]) return { ok: false, msg: 'Ya tienes vinculado a ' + objetivoNombre };
  sheet.appendRow([String(idSolicitante), String(objetivoId), new Date().toISOString(), '', '', false, false]);
  return { ok: true, msg: 'Ahora estas vinculado con ' + objetivoNombre };
}

function recuperarContrasenaPorCorreo(correoBuscado) {
  const correo = String(correoBuscado || '').trim().toLowerCase();
  if (!correo || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) {
    return { ok: false, msg: 'Escribe un correo electronico valido' };
  }
  initApp();
  const sheet = usuariosSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const idFila = String(data[i][0]).toLowerCase().trim();
    const correoFila = String(data[i][5] || '').toLowerCase().trim();
    if (idFila === correo || correoFila === correo) {
      const nuevaPass = String(Math.floor(100000 + Math.random() * 900000));
      sheet.getRange(i + 1, 3).setValue(nuevaPass);
      try {
        enviarRecuperacion(correo, String(data[i][1]), nuevaPass);
      } catch (err) {
        return { ok: false, msg: 'No se pudo enviar el correo de recuperacion.' };
      }
      return { ok: true, msg: 'Te hemos enviado una nueva contrasena a ' + correo + '.' };
    }
  }
  return { ok: false, msg: 'Ese correo no existe en nuestra aplicacion.' };
}

function enviarRecuperacion(correo, nombre, nuevaPass) {
  const url = URL_APP; // la PWA de GitHub Pages (con icono), no la /exec
  const cuerpo =
    '<p>Has solicitado recuperar tu contrasena de <b>' + NOMBRE_APP + '</b>. Esta es tu nueva contrasena temporal:</p>' +
    '<div style="background:#f4f1ea;border-radius:10px;padding:14px;margin:14px 0">' +
    '<p style="margin:0"><b>Nueva contrasena:</b> ' + nuevaPass + '</p></div>' +
    '<p>Por seguridad, cambiala despues de iniciar sesion desde "Mi perfil".</p>';
  MailApp.sendEmail({
    to: correo,
    subject: 'Recuperacion de contrasena - ' + NOMBRE_APP,
    htmlBody: plantillaCorreo('Hola, ' + escaparHtml(nombre) + '!', cuerpo, url),
    body: 'Tu nueva contrasena de ' + NOMBRE_APP + ' es: ' + nuevaPass + (url ? (' / Enlace: ' + url) : '')
  });
}

// ============================================================
// NOMBRES POR VISOR + EDICION DE NOMBRE
// ============================================================
function mapaNombresParaVisor(viewerId) {
  const data = usuariosData();
  const dataV = vinculosSheet();
  const v = String(viewerId);
  const mapa = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]);
    const publico = String(data[i][1]);
    const aliasCreador = String(data[i][7] || '');
    if (String(data[i][3]) === v && aliasCreador.trim() !== '') {
      mapa[id] = aliasCreador;
    } else {
      const aliasVinc = aliasDeVinculo(v, id, dataV);
      mapa[id] = aliasVinc.trim() !== '' ? aliasVinc : publico;
    }
  }
  return mapa;
}
function nombreUsuario(id, mapaNombres) {
  return mapaNombres[String(id)] || '(usuario eliminado)';
}

function editarNombreUsuarioCreado(idUsuario, nuevoNombre, idSolicitante) {
  const nombre = String(nuevoNombre || '').trim();
  if (!nombre) return { ok: false, msg: 'Escribe un nombre' };
  const sheet = usuariosSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idUsuario)) {
      if (String(data[i][3]) !== String(idSolicitante)) return { ok: false, msg: 'Solo puedes renombrar usuarios que hayas creado tu' };
      sheet.getRange(i + 1, 8).setValue(nombre);
      return { ok: true, usuario: nombre };
    }
  }
  return { ok: false, msg: 'Usuario no encontrado' };
}

function editarMiNombre(nuevoNombre, id) {
  const nombre = String(nuevoNombre || '').trim();
  if (!nombre) return { ok: false, msg: 'Escribe un nombre' };
  const sheet = usuariosSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 2).setValue(nombre);
      return { ok: true, usuario: nombre };
    }
  }
  return { ok: false, msg: 'Usuario no encontrado' };
}

// ============================================================
// DEUDAS
// ============================================================
function crearDeuda(acreedorId, deudorId, descripcion, importe, interesMensual) {
  if (String(acreedorId) === String(deudorId)) return { ok: false, msg: 'No puedes crear una deuda contigo mismo' };
  const capital = pesos(importe);
  if (!capital || capital <= 0) return { ok: false, msg: 'Importe no valido' };
  let tasa = Number(interesMensual) || 0;
  if (tasa < 0 || tasa > 100) return { ok: false, msg: 'El interes mensual debe estar entre 0 y 100%' };
  tasa = Math.round(tasa * 100) / 100;

  const _plan = estadoPlan(acreedorId);
  if (!_plan.puedeCrear) return { ok: false, limite: true, plan: _plan, msg: 'Has alcanzado el limite de deudas gratuitas.' };

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('deudas').appendRow([
    Utilities.getUuid(), String(acreedorId), String(deudorId), String(descripcion || ''), capital,
    new Date().toISOString(), 'activo', '', '', tasa
  ]);
  return { ok: true };
}

function itemDeuda(fila, pagosData, solicitudesData, nombres) {
  const id = String(fila[0]);
  const acreedorId = String(fila[1]);
  const deudorId = String(fila[2]);
  const c = estadoDeFila(fila, pagosData);
  return {
    id: id, acreedorId: acreedorId, deudorId: deudorId,
    acreedor: nombreUsuario(acreedorId, nombres), deudor: nombreUsuario(deudorId, nombres),
    descripcion: fila[3], importe: c.capitalInicial, fecha: fila[5], fechaPago: fila[7],
    estado: String(fila[6]), nota: fila[8] || '',
    tasa: c.tasa,
    capitalPendiente: c.capitalPendiente, interesPendiente: c.interesPendiente,
    interesGenerado: c.interesGenerado, interesPagado: c.interesPagado,
    capitalPagado: c.capitalPagado, pagado: c.totalPagado, restante: c.restante,
    mesesCumplidos: c.mesesCumplidos, proximoCorte: c.proximoCorte, interesProximo: c.interesProximo,
    movimientos: c.movimientos,
    solicitud: solicitudesData ? solicitudPendienteDeDeuda(id, solicitudesData) : null,
    rechazos: solicitudesData ? solicitudesRechazadasDeDeuda(id, solicitudesData) : []
  };
}

function listarDeudas(usuarioId) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('deudas').getDataRange().getValues();
  const pagosData = pagosDeSheet();
  const solicitudesData = solicitudesSheet();
  const nombres = mapaNombresParaVisor(usuarioId);
  const u = String(usuarioId);
  const meDeben = [], debo = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][6]) !== 'activo') continue;
    const acreedorId = String(data[i][1]);
    const deudorId = String(data[i][2]);
    if (acreedorId !== u && deudorId !== u) continue;
    const item = itemDeuda(data[i], pagosData, solicitudesData, nombres);
    if (acreedorId === u) meDeben.push(item); else debo.push(item);
  }
  const porFecha = (a, b) => new Date(b.fecha) - new Date(a.fecha);
  meDeben.sort(porFecha); debo.sort(porFecha);
  return { meDeben: meDeben, debo: debo };
}

function listarHistorial(usuarioId) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('deudas').getDataRange().getValues();
  const pagosData = pagosDeSheet();
  const nombres = mapaNombresParaVisor(usuarioId);
  const u = String(usuarioId);
  const historial = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][6]) === 'activo') continue;
    const acreedorId = String(data[i][1]);
    const deudorId = String(data[i][2]);
    if (acreedorId !== u && deudorId !== u) continue;
    const item = itemDeuda(data[i], pagosData, null, nombres);
    item.rol = acreedorId === u ? 'acreedor' : 'deudor';
    historial.push(item);
  }
  historial.sort((a, b) => new Date(b.fechaPago) - new Date(a.fechaPago));
  return historial;
}

// Las deudas no se borran: se archivan (los intereses dejan de correr ese dia).
function eliminarDeuda(idDeuda, usuarioId) {
  const r = buscarFilaDeuda(idDeuda);
  if (!r) return { ok: false, msg: 'Deuda no encontrada' };
  if (String(r.fila[1]) !== String(usuarioId)) return { ok: false, msg: 'Solo quien creo la deuda puede archivarla' };
  if (String(r.fila[6]) !== 'activo') return { ok: false, msg: 'Esta deuda ya no esta activa' };
  r.sheet.getRange(r.index + 1, 7).setValue('archivada');
  r.sheet.getRange(r.index + 1, 8).setValue(new Date().toISOString());
  return { ok: true };
}

// ============================================================
// PAGOS (parciales o totales) — primero a intereses, luego a capital
// ============================================================
function pagosDeSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pagos_deuda').getDataRange().getValues();
}

function pagosDeDeuda(idDeuda, data) {
  const rows = data || pagosDeSheet();
  const pagos = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(idDeuda)) pagos.push({ id: rows[i][0], monto: pesos(rows[i][2]), fecha: rows[i][3] });
  }
  pagos.sort((a, b) => new Date(a.fecha) - new Date(b.fecha));
  return pagos;
}

// esTotal: si es true, se paga exactamente lo que quede (capital + intereses) en este momento.
function registrarPago(idDeuda, monto, usuarioId, esTotal) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = buscarFilaDeuda(idDeuda);
    if (!r) return { ok: false, msg: 'Deuda no encontrada' };
    if (String(r.fila[1]) !== String(usuarioId)) return { ok: false, msg: 'Solo quien creo la deuda puede registrar pagos' };
    if (String(r.fila[6]) !== 'activo') return { ok: false, msg: 'Esta deuda ya no esta activa' };

    const est = estadoDeFila(r.fila, null);
    const restante = est.restante;
    let m = esTotal ? restante : pesos(monto);
    if (!m || m <= 0) return { ok: false, msg: 'Importe no valido' };
    if (m > restante) return { ok: false, msg: 'El importe supera lo que queda pendiente (' + fmtCOP(restante) + ')' };

    const aInt = Math.min(m, est.interesPendiente);
    const aCap = m - aInt;
    const ahora = new Date().toISOString();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pagos_deuda').appendRow([
      Utilities.getUuid(), String(idDeuda), m, ahora, aInt, aCap
    ]);

    const nuevoRestante = restante - m;
    const completado = nuevoRestante <= 0;
    if (completado) {
      r.sheet.getRange(r.index + 1, 7).setValue('pagado');
      r.sheet.getRange(r.index + 1, 8).setValue(ahora);
    }
    return { ok: true, completado: completado, restante: completado ? 0 : nuevoRestante, monto: m, aInteres: aInt, aCapital: aCap };
  } finally {
    lock.releaseLock();
  }
}

function obtenerNota(idDeuda, usuarioId) {
  const r = buscarFilaDeuda(idDeuda);
  if (!r) return { ok: false, msg: 'Deuda no encontrada' };
  const u = String(usuarioId);
  if (String(r.fila[1]) !== u && String(r.fila[2]) !== u) return { ok: false, msg: 'No tienes acceso a esta deuda' };
  return { ok: true, nota: r.fila[8] || '', esCreador: String(r.fila[1]) === u };
}

function guardarNota(idDeuda, nota, usuarioId) {
  const r = buscarFilaDeuda(idDeuda);
  if (!r) return { ok: false, msg: 'Deuda no encontrada' };
  if (String(r.fila[1]) !== String(usuarioId)) return { ok: false, msg: 'Solo quien creo la deuda puede editar la nota' };
  r.sheet.getRange(r.index + 1, 9).setValue(nota || '');
  return { ok: true };
}

// ============================================================
// CONTACTOS OMITIDOS
// ============================================================
function omitidosSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('omitidos').getDataRange().getValues();
}

function existeDeudaActivaEntre(id1, id2) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('deudas').getDataRange().getValues();
  const a = String(id1), b = String(id2);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][6]) !== 'activo') continue;
    const x = String(data[i][1]), y = String(data[i][2]);
    if ((x === a && y === b) || (x === b && y === a)) return true;
  }
  return false;
}

function omitirUsuario(usuarioId, omitidoId) {
  if (String(usuarioId) === String(omitidoId)) return { ok: false, msg: 'No puedes omitirte a ti mismo' };
  if (existeDeudaActivaEntre(usuarioId, omitidoId)) return { ok: false, msg: 'No puedes omitir a alguien con deudas pendientes entre ustedes' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('omitidos');
  const data = sheet.getDataRange().getValues();
  const u = String(usuarioId), o = String(omitidoId);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === u && String(data[i][1]) === o) return { ok: true };
  }
  sheet.appendRow([u, o, new Date().toISOString()]);
  return { ok: true };
}

function recuperarUsuario(usuarioId, omitidoId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('omitidos');
  const data = sheet.getDataRange().getValues();
  const u = String(usuarioId), o = String(omitidoId);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === u && String(data[i][1]) === o) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, msg: 'Ese usuario no estaba omitido por ti' };
}

function hayOmisionMutua(id1, id2, data) {
  const rows = data || omitidosSheet();
  const a = String(id1), b = String(id2);
  for (let i = 1; i < rows.length; i++) {
    const x = String(rows[i][0]), y = String(rows[i][1]);
    if ((x === a && y === b) || (x === b && y === a)) return true;
  }
  return false;
}

function listarUsuariosParaDeuda(usuarioId) {
  const omitidosData = omitidosSheet();
  const u = String(usuarioId);
  return listarUsuariosVinculados(usuarioId).filter(us => String(us.id) !== u && !hayOmisionMutua(usuarioId, us.id, omitidosData));
}

function listarContactos(usuarioId) {
  const data = usuariosData();
  const omitidosData = omitidosSheet();
  const u = String(usuarioId);
  const misOmitidos = {};
  for (let i = 1; i < omitidosData.length; i++) {
    if (String(omitidosData[i][0]) === u) misOmitidos[String(omitidosData[i][1])] = true;
  }
  const vinc = idsVinculados(usuarioId);
  const nombres = mapaNombresParaVisor(usuarioId);
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]);
    if (id === u || !vinc[id]) continue;
    out.push({ id: id, usuario: nombres[id] || String(data[i][1]), omitido: !!misOmitidos[id] });
  }
  return out;
}

// ============================================================
// SOLICITUDES DE PAGO (el deudor solicita, el acreedor acepta/rechaza)
// ============================================================
function solicitudesSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('solicitudes_pago').getDataRange().getValues();
}

function solicitudPendienteDeDeuda(idDeuda, data) {
  const rows = data || solicitudesSheet();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(idDeuda) && String(rows[i][5]) === 'pendiente') {
      return { id: String(rows[i][0]), idDeuda: String(rows[i][1]), deudorId: String(rows[i][2]), monto: pesos(rows[i][3]), tipo: String(rows[i][4]), fecha: rows[i][6] };
    }
  }
  return null;
}

function solicitudesRechazadasDeDeuda(idDeuda, data) {
  const rows = data || solicitudesSheet();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const estado = String(rows[i][5]);
    if (String(rows[i][1]) === String(idDeuda) && (estado === 'rechazada' || estado === 'cancelada')) {
      out.push({ monto: pesos(rows[i][3]), tipo: String(rows[i][4]), estado: estado, fecha: rows[i][7] || rows[i][6] });
    }
  }
  out.sort((a, b) => new Date(a.fecha) - new Date(b.fecha));
  return out;
}

function crearSolicitudPago(idDeuda, tipo, monto, usuarioId) {
  const r = buscarFilaDeuda(idDeuda);
  if (!r) return { ok: false, msg: 'Deuda no encontrada' };
  if (String(r.fila[6]) !== 'activo') return { ok: false, msg: 'Esta deuda ya no esta activa' };
  if (String(r.fila[2]) !== String(usuarioId)) return { ok: false, msg: 'Solo el deudor puede solicitar un pago' };
  if (solicitudPendienteDeDeuda(idDeuda)) return { ok: false, msg: 'Ya hay una solicitud pendiente para esta deuda' };

  const restante = estadoDeFila(r.fila, null).restante;
  const t = String(tipo) === 'total' ? 'total' : 'parcial';
  let montoSolicitado;
  if (t === 'total') {
    montoSolicitado = restante;
  } else {
    montoSolicitado = pesos(monto);
    if (!montoSolicitado || montoSolicitado <= 0) return { ok: false, msg: 'Importe no valido' };
    if (montoSolicitado > restante) return { ok: false, msg: 'El importe supera lo que queda pendiente (' + fmtCOP(restante) + ')' };
  }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('solicitudes_pago').appendRow([
    Utilities.getUuid(), String(idDeuda), String(usuarioId), montoSolicitado, t, 'pendiente', new Date().toISOString(), ''
  ]);
  return { ok: true, msg: 'Solicitud de pago enviada' };
}

// Al aceptar un pago TOTAL se cobra lo que quede en ese momento (por si
// entre la solicitud y la aceptacion se cumplio otro mes de interes).
function responderSolicitud(idSolicitud, aceptar, usuarioId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSol = ss.getSheetByName('solicitudes_pago');
  const dataSol = sheetSol.getDataRange().getValues();
  let filaSol = -1;
  for (let i = 1; i < dataSol.length; i++) {
    if (String(dataSol[i][0]) === String(idSolicitud)) { filaSol = i; break; }
  }
  if (filaSol === -1) return { ok: false, msg: 'Solicitud no encontrada' };
  if (String(dataSol[filaSol][5]) !== 'pendiente') return { ok: false, msg: 'Esta solicitud ya fue respondida' };

  const idDeuda = String(dataSol[filaSol][1]);
  const monto = pesos(dataSol[filaSol][3]);
  const tipo = String(dataSol[filaSol][4]);
  const r = buscarFilaDeuda(idDeuda);
  if (!r) return { ok: false, msg: 'Deuda no encontrada' };
  const acreedorId = String(r.fila[1]);
  const deudorId = String(r.fila[2]);

  if (aceptar) {
    if (String(usuarioId) !== acreedorId) return { ok: false, msg: 'Solo el acreedor puede aceptar la solicitud' };
    const res = registrarPago(idDeuda, monto, acreedorId, tipo === 'total');
    if (!res.ok) return res;
    sheetSol.getRange(filaSol + 1, 4).setValue(res.monto);
    sheetSol.getRange(filaSol + 1, 6).setValue('aceptada');
    sheetSol.getRange(filaSol + 1, 8).setValue(new Date().toISOString());
    return { ok: true, completado: res.completado, restante: res.restante, msg: res.completado ? 'Pago aceptado. Deuda saldada.' : 'Pago aceptado.' };
  }
  if (String(usuarioId) !== acreedorId && String(usuarioId) !== deudorId) return { ok: false, msg: 'No tienes permiso sobre esta solicitud' };
  const nuevoEstado = String(usuarioId) === deudorId ? 'cancelada' : 'rechazada';
  sheetSol.getRange(filaSol + 1, 6).setValue(nuevoEstado);
  sheetSol.getRange(filaSol + 1, 8).setValue(new Date().toISOString());
  return { ok: true, msg: nuevoEstado === 'cancelada' ? 'Solicitud cancelada' : 'Solicitud rechazada' };
}

// ============================================================
// LICENCIAS / PLAN
// ============================================================
const LIMITE_GRATIS_TOTAL = 10;
const LIMITE_MENSUAL = 2;

function _asegurarColLicencia_() {
  const sh = usuariosSheet();
  const lastCol = sh.getLastColumn();
  const cab = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let idx = cab.indexOf('licenciaHasta');
  if (idx === -1) {
    sh.getRange(1, lastCol + 1).setValue('licenciaHasta');
    idx = lastCol;
  }
  return idx;
}

function _licenciaHastaDe_(userId) {
  const idx = _asegurarColLicencia_();
  const data = usuariosData();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      const v = data[i][idx];
      if (!v) return null;
      const d = aFecha(v);
      return d || null;
    }
  }
  return null;
}

function _conteoCreadas_(userId) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('deudas').getDataRange().getValues();
  const ahora = new Date();
  const mes = ahora.getMonth(), anio = ahora.getFullYear();
  let total = 0, esteMes = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(userId)) {
      total++;
      const d = aFecha(data[i][5]);
      if (d && d.getMonth() === mes && d.getFullYear() === anio) esteMes++;
    }
  }
  return { total: total, esteMes: esteMes };
}

function estadoPlan(userId) {
  const u = obtenerUsuarioPorId(userId);
  if (u && esCuentaMaster(u.usuario, u.correo)) {
    return {
      plan: 'licencia', licenciaActiva: true, licenciaHasta: new Date(9999, 11, 31).toISOString(), vitalicia: true,
      creadasTotal: 0, creadasMes: 0, restantes: -1, limiteTotal: LIMITE_GRATIS_TOTAL, limiteMensual: LIMITE_MENSUAL,
      puedeCrear: true, motivo: ''
    };
  }
  const lic = _licenciaHastaDe_(userId);
  const licenciaActiva = !!(lic && lic.getTime() > Date.now());
  const c = _conteoCreadas_(userId);
  let plan, puedeCrear, restantes, motivo = '';
  if (licenciaActiva) { plan = 'licencia'; puedeCrear = true; restantes = -1; }
  else if (c.total < LIMITE_GRATIS_TOTAL) { plan = 'gratis_inicial'; puedeCrear = true; restantes = LIMITE_GRATIS_TOTAL - c.total; }
  else {
    plan = 'gratis_mensual';
    restantes = Math.max(0, LIMITE_MENSUAL - c.esteMes);
    puedeCrear = restantes > 0;
    if (!puedeCrear) motivo = 'limite_mensual';
  }
  return {
    plan: plan, licenciaActiva: licenciaActiva, licenciaHasta: lic ? lic.toISOString() : null,
    vitalicia: !!(lic && lic.getFullYear() >= 9999), creadasTotal: c.total, creadasMes: c.esteMes,
    restantes: restantes, limiteTotal: LIMITE_GRATIS_TOTAL, limiteMensual: LIMITE_MENSUAL,
    puedeCrear: puedeCrear, motivo: motivo
  };
}

function _esMasterPorId_(id) {
  const u = obtenerUsuarioPorId(id);
  return !!(u && esCuentaMaster(u.usuario, u.correo));
}

function activarLicenciaMaster(solicitanteId, correoCliente, tipo) {
  if (!_esMasterPorId_(solicitanteId)) return { ok: false, msg: 'No autorizado' };
  const correo = String(correoCliente || '').trim().toLowerCase();
  if (!correo) return { ok: false, msg: 'Indica un correo' };
  const sh = usuariosSheet();
  const idxLic = _asegurarColLicencia_();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][5] || '').trim().toLowerCase() === correo) {
      let hasta;
      if (tipo === 'vitalicia') hasta = new Date(9999, 11, 31);
      else { hasta = new Date(); hasta.setFullYear(hasta.getFullYear() + 1); }
      sh.getRange(i + 1, idxLic + 1).setValue(hasta.toISOString());
      return { ok: true, correo: correo, hasta: hasta.toISOString(), vitalicia: tipo === 'vitalicia' };
    }
  }
  return { ok: false, msg: 'No hay ningun usuario con ese correo' };
}

function quitarLicenciaMaster(solicitanteId, correoCliente) {
  if (!_esMasterPorId_(solicitanteId)) return { ok: false, msg: 'No autorizado' };
  const correo = String(correoCliente || '').trim().toLowerCase();
  if (!correo) return { ok: false, msg: 'Indica un correo' };
  const sh = usuariosSheet();
  const idxLic = _asegurarColLicencia_();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][5] || '').trim().toLowerCase() === correo) {
      sh.getRange(i + 1, idxLic + 1).setValue('');
      return { ok: true, correo: correo };
    }
  }
  return { ok: false, msg: 'No hay ningun usuario con ese correo' };
}

// ============================================================
// ROUTER
// ============================================================
function ejecutar(accion, params) {
  params = params || {};
  switch (accion) {
    case 'estadoPlan': return estadoPlan(params.userId);
    case 'activarLicenciaMaster': return activarLicenciaMaster(params.id, params.correoCliente, params.tipo);
    case 'quitarLicenciaMaster': return quitarLicenciaMaster(params.id, params.correoCliente);
    case 'login': return login(params.usuario, params.contrasena);
    case 'cambiarContrasena': return cambiarContrasena(params.id, params.actual, params.nueva);
    case 'crearUsuario': return crearUsuario(params.usuarioNuevo, params.correoNuevo, params.creadorId);
    case 'listarMisUsuarios': return listarMisUsuarios(params.id);
    case 'eliminarUsuario': return eliminarUsuario(params.idUsuario, params.idSolicitante);
    case 'recuperarUsuarioEliminado': return recuperarUsuarioEliminado(params.correo, params.id);
    case 'agregarUsuarioPorCorreo': return agregarUsuarioPorCorreo(params.correo, params.id);
    case 'recuperarContrasena': return recuperarContrasenaPorCorreo(params.correo);
    case 'editarNombreUsuarioCreado': return editarNombreUsuarioCreado(params.idUsuario, params.nuevoNombre, params.id);
    case 'editarAliasVinculo': return editarAliasVinculo(params.idUsuario, params.nuevoNombre, params.id);
    case 'ocultarVinculo': return ocultarVinculo(params.idUsuario, params.id);
    case 'editarMiNombre': return editarMiNombre(params.nuevoNombre, params.id);
    case 'crearDeuda': return crearDeuda(params.acreedorId, params.deudorId, params.descripcion, params.importe, params.interesMensual);
    case 'listarDeudas': return listarDeudas(params.id);
    case 'listarHistorial': return listarHistorial(params.id);
    case 'eliminarDeuda': return eliminarDeuda(params.idDeuda, params.id);
    case 'registrarPago': return registrarPago(params.idDeuda, params.monto, params.id, !!params.total);
    case 'crearSolicitudPago': return crearSolicitudPago(params.idDeuda, params.tipo, params.monto, params.id);
    case 'responderSolicitud': return responderSolicitud(params.idSolicitud, params.aceptar, params.id);
    case 'guardarNota': return guardarNota(params.idDeuda, params.nota, params.id);
    case 'obtenerNota': return obtenerNota(params.idDeuda, params.id);
    case 'listarUsuariosParaDeuda': return listarUsuariosParaDeuda(params.id);
    case 'listarContactos': return listarContactos(params.id);
    case 'omitirUsuario': return omitirUsuario(params.id, params.omitidoId);
    case 'recuperarUsuario': return recuperarUsuario(params.id, params.omitidoId);
    default: return { ok: false, msg: 'Accion desconocida' };
  }
}

// ============================================================
// PRUEBA DEL MOTOR (ejecutar desde el editor y mirar el registro)
// ============================================================
function probarMotorIntereses() {
  const inicio = new Date(2026, 0, 15);
  const pagos = [
    { id: 'p1', monto: 150000, fecha: new Date(2026, 2, 20) },
    { id: 'p2', monto: 500000, fecha: new Date(2026, 4, 2) }
  ];
  const r = calcularDeuda(1000000, 5, inicio, pagos, new Date(2026, 6, 1));
  Logger.log(JSON.stringify(r, null, 2));
}
