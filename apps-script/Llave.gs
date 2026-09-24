/**
 * LLAVE NFC para DeudasPesos.
 * Habla con la Central de Llaves. El secreto nunca sale al navegador.
 * (Usa las mismas credenciales de app que DeudasApp en euros.)
 */
var LLAVE_CENTRAL_URL = 'https://script.google.com/macros/s/AKfycbxqb6AB1Tg5xgygsU2XzDSw8fqfk9wHPOJFTEJ3LHBNebhOL72FSfM_EvVcQs6_HIFW/exec';
var LLAVE_APP_ID = 'deudas';
var LLAVE_APP_SECRETO = 'PON_AQUI_EL_SECRETO';

// titular de la llave (en minusculas) -> nombre de usuario en esta app
var LLAVE_CUENTAS = {
  'manu': { usuario: 'Manu' }
};

function llaveLlamar(datos) {
  var carga = Object.assign({ app: LLAVE_APP_ID, secreto: LLAVE_APP_SECRETO }, datos);
  var r = UrlFetchApp.fetch(LLAVE_CENTRAL_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(carga),
    followRedirects: true,
    muteHttpExceptions: true
  });
  var d;
  try { d = JSON.parse(r.getContentText()); }
  catch (e) { throw new Error('La central de llaves no responde'); }
  if (!d.ok) throw new Error(d.error || 'Error en la central de llaves');
  return d;
}

function llavePedirSesion() {
  try {
    var d = llaveLlamar({ accion: 'sesion' });
    return { ok: true, codigo: d.codigo, url: d.url };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function llaveComprobar(codigo) {
  try {
    var e = llaveLlamar({ accion: 'estado', codigo: codigo });
    if (e.estado === 'esperando') return { estado: 'esperando' };
    if (e.estado === 'caducado') return { estado: 'caducado' };
    if (e.estado === 'denegado') return { estado: 'denegado', motivo: e.motivo || 'Llave no autorizada' };

    var id = llaveLlamar({ accion: 'canjear', codigo: codigo });
    var titular = String(id.usuario || '').toLowerCase();
    var cuenta = LLAVE_CUENTAS[titular];
    if (!cuenta) return { estado: 'denegado', motivo: 'El titular "' + titular + '" no tiene cuenta en DeudasPesos' };
    var res = _loginPorNombre(cuenta.usuario);
    if (!res.ok) return { estado: 'denegado', motivo: 'No se pudo iniciar sesion con esa cuenta' };
    return { estado: 'ok', usuario: res };
  } catch (err) {
    return { estado: 'denegado', motivo: err.message };
  }
}
