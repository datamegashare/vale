// ═══════════════════════════════════════════════════════════════
// Vale Digital — Backend GAS v2.0  (Etapa 1 — fix redirect)
// Endpoints: login, guardarVale, getMisVales, eliminarVale
// ═══════════════════════════════════════════════════════════════

// ── CONFIGURACIÓN ───────────────────────────────────────────────
const SPREADSHEET_ID = '10XkTAarQdgucz8WIwNh5FhV2vyW9qdoRhRCKnQ7wb4k'; // ← ID del Sheet (URL entre /d/ y /edit)

const HOJA = {
  VALES    : 'VALES',
  USUARIOS : 'USUARIOS',
  OBRAS    : 'OBRAS'
};

const COL_VALES = {
  id_vale        : 1,
  fecha_hora     : 2,
  usuario_email  : 3,
  usuario_nombre : 4,
  obra_codigo    : 5,
  destino        : 6,
  titulo         : 7,
  contenido_html : 8,
  estado         : 9,
  aprobado_por   : 10,
  fecha_aprobacion: 11,
  gestionado_por : 12,
  fecha_cierre   : 13,
  nota_cierre    : 14,
  eliminado      : 15
};

const COL_USUARIOS = {
  email      : 1,
  nombre     : 2,
  rol        : 3,
  obra_codigo: 4,
  activo     : 5
};

const COL_OBRAS = {
  codigo      : 1,
  descripcion : 2,
  email_pañol : 3,
  email_almacen: 4,
  activa      : 5
};

// ── HELPERS ─────────────────────────────────────────────────────

function getSheet(nombre) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(nombre);
}

function respOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function respError(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetToObjects(hoja, colMap) {
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return [];
  const headers = Object.keys(colMap);
  return datos.slice(1).map(fila => {
    const obj = {};
    headers.forEach(key => { obj[key] = fila[colMap[key] - 1]; });
    return obj;
  });
}

function parseBody(e) {
  try {
    return JSON.parse(e.postData.contents);
  } catch (_) {
    return {};
  }
}

// ── ROUTER ──────────────────────────────────────────────────────

// Todo el tráfico va por GET para evitar pérdida de body en el redirect 302 de GAS
function doGet(e) {
  const accion = e.parameter.accion || '';
  try {
    switch (accion) {
      case 'login'        : return acLogin(e);
      case 'getMisVales'  : return acGetMisVales(e);
      case 'guardarVale'  : return acGuardarVale(e.parameter);
      case 'eliminarVale' : return acEliminarVale(e.parameter);
      default             : return respError('Acción desconocida: ' + accion);
    }
  } catch (err) {
    return respError('Error interno [' + accion + ']: ' + err.message);
  }
}

// doPost ya no se usa — se mantiene por compatibilidad
function doPost(e) {
  return doGet(e);
}

// ── ACCIÓN: login ────────────────────────────────────────────────
// GET ?accion=login&email=xxx
function acLogin(e) {
  const email = (e.parameter.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  const hUsuarios = getSheet(HOJA.USUARIOS);
  const usuarios  = sheetToObjects(hUsuarios, COL_USUARIOS);
  const usuario   = usuarios.find(u => u.email.toLowerCase() === email);

  if (!usuario)         return respError('Usuario no registrado.');
  if (!usuario.activo)  return respError('Usuario inactivo. Contactá al administrador.');

  const hObras = getSheet(HOJA.OBRAS);
  const obras  = sheetToObjects(hObras, COL_OBRAS);
  const obra   = obras.find(o => o.codigo === usuario.obra_codigo && o.activa);

  if (!obra) return respError('Obra asignada no encontrada o inactiva.');

  return respOk({
    email      : usuario.email,
    nombre     : usuario.nombre,
    rol        : usuario.rol,
    obra_codigo: usuario.obra_codigo,
    obra_desc  : obra.descripcion
  });
}

// ── ACCIÓN: getMisVales ──────────────────────────────────────────
// GET ?accion=getMisVales&email=xxx
// Retorna vales del usuario de los últimos 30 días, excluyendo eliminados.
function acGetMisVales(e) {
  const email = (e.parameter.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  const hVales = getSheet(HOJA.VALES);
  const vales  = sheetToObjects(hVales, COL_VALES);

  const hace30 = new Date();
  hace30.setDate(hace30.getDate() - 30);

  const misvales = vales.filter(v => {
    if (v.usuario_email.toLowerCase() !== email) return false;
    if (v.eliminado === true || v.eliminado === 'TRUE') return false;
    const fecha = new Date(v.fecha_hora);
    return fecha >= hace30;
  });

  // Ordenar descendente por fecha
  misvales.sort((a, b) => new Date(b.fecha_hora) - new Date(a.fecha_hora));

  return respOk(misvales);
}

// ── ACCIÓN: guardarVale ──────────────────────────────────────────
// POST { accion, id_vale, usuario_email, usuario_nombre, obra_codigo,
//        destino, titulo, contenido_html, estado, fecha_hora }
function acGuardarVale(params) {
  // Los parámetros llegan como strings desde GET
  const contenido_raw = params.contenido_html || '';
  const body = {
    id_vale        : params.id_vale,
    usuario_email  : params.usuario_email,
    usuario_nombre : params.usuario_nombre,
    obra_codigo    : params.obra_codigo,
    destino        : params.destino,
    titulo         : params.titulo,
    contenido_html : contenido_raw,
    estado         : params.estado,
    fecha_hora     : params.fecha_hora
  };
  const { id_vale, usuario_email, usuario_nombre, obra_codigo,
          destino, titulo, contenido_html, estado, fecha_hora } = body;

  if (!id_vale)        return respError('id_vale requerido.');
  if (!usuario_email)  return respError('usuario_email requerido.');
  if (!destino)        return respError('destino requerido (ALMACEN o PAÑOL).');
  if (!contenido_html) return respError('contenido_html requerido.');

  const estadoFinal = estado || 'BORRADOR';

  const hVales = getSheet(HOJA.VALES);
  const datos  = hVales.getDataRange().getValues();

  // Buscar si existe el vale por id (upsert)
  let filaExistente = -1;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][COL_VALES.id_vale - 1] === id_vale) {
      filaExistente = i + 1; // índice base 1 para GAS
      break;
    }
  }

  const ahora     = new Date();
  const fechaHora = fecha_hora || ahora.toISOString();

  if (filaExistente > 0) {
    // UPDATE — solo actualiza campos editables
    hVales.getRange(filaExistente, COL_VALES.titulo, 1, 1).setValue(titulo || '');
    hVales.getRange(filaExistente, COL_VALES.contenido_html, 1, 1).setValue(contenido_html);
    hVales.getRange(filaExistente, COL_VALES.destino, 1, 1).setValue(destino);
    hVales.getRange(filaExistente, COL_VALES.estado, 1, 1).setValue(estadoFinal);
  } else {
    // INSERT — fila nueva
    const nuevaFila = [
      id_vale,
      fechaHora,
      usuario_email,
      usuario_nombre  || '',
      obra_codigo     || '',
      destino,
      titulo          || '',
      contenido_html,
      estadoFinal,
      '', // aprobado_por
      '', // fecha_aprobacion
      '', // gestionado_por
      '', // fecha_cierre
      '', // nota_cierre
      false // eliminado
    ];
    hVales.appendRow(nuevaFila);
  }

  SpreadsheetApp.flush();
  return respOk({ id_vale, estado: estadoFinal, accion: filaExistente > 0 ? 'updated' : 'inserted' });
}

// ── ACCIÓN: eliminarVale ─────────────────────────────────────────
// POST { accion, id_vale, usuario_email }
// Solo permite eliminar vales en estado BORRADOR del propio usuario.
function acEliminarVale(params) {
  const { id_vale, usuario_email } = params;
  if (!id_vale)       return respError('id_vale requerido.');
  if (!usuario_email) return respError('usuario_email requerido.');

  const hVales = getSheet(HOJA.VALES);
  const datos  = hVales.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (fila[COL_VALES.id_vale - 1] !== id_vale) continue;

    // Validar dueño
    if (fila[COL_VALES.usuario_email - 1].toLowerCase() !== usuario_email.toLowerCase()) {
      return respError('No podés eliminar un vale que no es tuyo.');
    }
    // Validar estado
    const estadoActual = fila[COL_VALES.estado - 1];
    if (estadoActual !== 'BORRADOR') {
      return respError('Solo se pueden eliminar vales en estado BORRADOR. Estado actual: ' + estadoActual);
    }

    // Borrado lógico
    const numFila = i + 1;
    hVales.getRange(numFila, COL_VALES.eliminado, 1, 1).setValue(true);
    hVales.getRange(numFila, COL_VALES.estado, 1, 1).setValue('ELIMINADO');
    hVales.getRange(numFila, COL_VALES.fecha_cierre, 1, 1).setValue(new Date().toISOString());

    SpreadsheetApp.flush();
    return respOk({ id_vale, eliminado: true });
  }

  return respError('Vale no encontrado: ' + id_vale);
}
