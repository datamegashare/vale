// ═══════════════════════════════════════════════════════════════
// Vale Digital — Backend GAS v3.0  (Etapa 2 — Panel de Aprobación)
// Endpoints v2: login, guardarVale (+ auto-aprobación JEFE_OBRA),
//               getMisVales, eliminarVale
// Endpoints v3: getPendientesAprobacion, aprobarVale,
//               rechazarVale, getValesAprobados
// ═══════════════════════════════════════════════════════════════

// ── CONFIGURACIÓN ───────────────────────────────────────────────
const SPREADSHEET_ID = '10XkTAarQdgucz8WIwNh5FhV2vyW9qdoRhRCKnQ7wb4k'; // ← ID del Sheet

const HOJA = {
  VALES    : 'VALES',
  USUARIOS : 'USUARIOS',
  OBRAS    : 'OBRAS'
};

const COL_VALES = {
  id_vale         : 1,
  fecha_hora      : 2,
  usuario_email   : 3,
  usuario_nombre  : 4,
  obra_codigo     : 5,
  destino         : 6,
  titulo          : 7,
  contenido_html  : 8,
  estado          : 9,
  aprobado_por    : 10,
  fecha_aprobacion: 11,
  gestionado_por  : 12,
  fecha_cierre    : 13,
  nota_cierre     : 14,
  eliminado       : 15
};

const COL_USUARIOS = {
  email      : 1,
  nombre     : 2,
  rol        : 3,
  obra_codigo: 4,
  activo     : 5
};

const COL_OBRAS = {
  codigo       : 1,
  descripcion  : 2,
  email_pañol  : 3,
  email_almacen: 4,
  activa       : 5
};

// Jerarquía de aprobación: quién puede aprobar a quién
// SUPERVISOR aprueba vales de CAPATAZ
// JEFE_OBRA  aprueba vales de SUPERVISOR
const PUEDE_APROBAR = {
  'SUPERVISOR': 'CAPATAZ',
  'JEFE_OBRA' : 'SUPERVISOR'
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

// Devuelve el objeto usuario o null
function buscarUsuario(email) {
  const hUsuarios = getSheet(HOJA.USUARIOS);
  const usuarios  = sheetToObjects(hUsuarios, COL_USUARIOS);
  return usuarios.find(u => u.email.toLowerCase() === email.toLowerCase()) || null;
}

// Encuentra la fila (base 1) de un vale por id_vale. Retorna -1 si no existe.
function buscarFilaVale(datos, id_vale) {
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][COL_VALES.id_vale - 1] === id_vale) return i + 1;
  }
  return -1;
}

// ── ROUTER ──────────────────────────────────────────────────────

// Todo el tráfico va por GET para evitar pérdida de body en el redirect 302 de GAS
function doGet(e) {
  const accion = e.parameter.accion || '';
  try {
    switch (accion) {
      // ── Etapa 1 ──
      case 'login'                    : return acLogin(e);
      case 'getMisVales'              : return acGetMisVales(e);
      case 'guardarVale'              : return acGuardarVale(e.parameter);
      case 'eliminarVale'             : return acEliminarVale(e.parameter);
      // ── Etapa 2 ──
      case 'getPendientesAprobacion'  : return acGetPendientesAprobacion(e.parameter);
      case 'aprobarVale'              : return acAprobarVale(e.parameter);
      case 'rechazarVale'             : return acRechazarVale(e.parameter);
      case 'getValesAprobados'        : return acGetValesAprobados(e.parameter);
      default                         : return respError('Acción desconocida: ' + accion);
    }
  } catch (err) {
    return respError('Error interno [' + accion + ']: ' + err.message);
  }
}

// doPost ya no se usa — se mantiene por compatibilidad
function doPost(e) {
  return doGet(e);
}

// ════════════════════════════════════════════════════════════════
// ETAPA 1 — ENDPOINTS
// ════════════════════════════════════════════════════════════════

// ── ACCIÓN: login ────────────────────────────────────────────────
// GET ?accion=login&email=xxx
function acLogin(e) {
  const email = (e.parameter.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  const usuario = buscarUsuario(email);
  if (!usuario)        return respError('Usuario no registrado.');
  if (!usuario.activo) return respError('Usuario inactivo. Contactá al administrador.');

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

  const misVales = vales.filter(v => {
    if (v.usuario_email.toLowerCase() !== email) return false;
    if (v.eliminado === true || v.eliminado === 'TRUE') return false;
    const fecha = new Date(v.fecha_hora);
    return fecha >= hace30;
  });

  misVales.sort((a, b) => new Date(b.fecha_hora) - new Date(a.fecha_hora));

  return respOk(misVales);
}

// ── ACCIÓN: guardarVale ──────────────────────────────────────────
// GET ?accion=guardarVale&id_vale=xxx&...
// ► ETAPA 2: si el solicitante es JEFE_OBRA y estado = ENVIADO → auto-aprobación
function acGuardarVale(params) {
  const body = {
    id_vale        : params.id_vale,
    usuario_email  : params.usuario_email,
    usuario_nombre : params.usuario_nombre,
    obra_codigo    : params.obra_codigo,
    destino        : params.destino,
    titulo         : params.titulo,
    contenido_html : params.contenido_html || '',
    estado         : params.estado,
    fecha_hora     : params.fecha_hora
  };

  const { id_vale, usuario_email, usuario_nombre, obra_codigo,
          destino, titulo, contenido_html, estado, fecha_hora } = body;

  if (!id_vale)        return respError('id_vale requerido.');
  if (!usuario_email)  return respError('usuario_email requerido.');
  if (!destino)        return respError('destino requerido (ALMACEN o PAÑOL).');
  if (!contenido_html) return respError('contenido_html requerido.');

  // ── Etapa 2: auto-aprobación JEFE_OBRA ──────────────────────
  let estadoFinal    = estado || 'BORRADOR';
  let aprobado_por   = '';
  let fecha_aprobacion = '';

  if (estadoFinal === 'ENVIADO') {
    const solicitante = buscarUsuario(usuario_email);
    if (solicitante && solicitante.rol === 'JEFE_OBRA') {
      estadoFinal      = 'APROBADO';
      aprobado_por     = usuario_email;
      fecha_aprobacion = new Date().toISOString();
    }
  }
  // ────────────────────────────────────────────────────────────

  const hVales = getSheet(HOJA.VALES);
  const datos  = hVales.getDataRange().getValues();

  let filaExistente = buscarFilaVale(datos, id_vale);

  const ahora     = new Date();
  const fechaHora = fecha_hora || ahora.toISOString();

  if (filaExistente > 0) {
    // UPDATE — actualiza campos editables
    hVales.getRange(filaExistente, COL_VALES.titulo,          1, 1).setValue(titulo || '');
    hVales.getRange(filaExistente, COL_VALES.contenido_html,  1, 1).setValue(contenido_html);
    hVales.getRange(filaExistente, COL_VALES.destino,         1, 1).setValue(destino);
    hVales.getRange(filaExistente, COL_VALES.estado,          1, 1).setValue(estadoFinal);
    if (aprobado_por) {
      hVales.getRange(filaExistente, COL_VALES.aprobado_por,     1, 1).setValue(aprobado_por);
      hVales.getRange(filaExistente, COL_VALES.fecha_aprobacion, 1, 1).setValue(fecha_aprobacion);
    }
  } else {
    // INSERT — fila nueva
    const nuevaFila = [
      id_vale,
      fechaHora,
      usuario_email,
      usuario_nombre   || '',
      obra_codigo      || '',
      destino,
      titulo           || '',
      contenido_html,
      estadoFinal,
      aprobado_por,      // aprobado_por   (vacío si no hay auto-aprobación)
      fecha_aprobacion,  // fecha_aprobacion
      '',                // gestionado_por
      '',                // fecha_cierre
      '',                // nota_cierre
      false              // eliminado
    ];
    hVales.appendRow(nuevaFila);
  }

  SpreadsheetApp.flush();
  return respOk({
    id_vale,
    estado : estadoFinal,
    accion : filaExistente > 0 ? 'updated' : 'inserted',
    auto_aprobado: estadoFinal === 'APROBADO' && aprobado_por !== ''
  });
}

// ── ACCIÓN: eliminarVale ─────────────────────────────────────────
// GET ?accion=eliminarVale&id_vale=xxx&usuario_email=xxx
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

    if (fila[COL_VALES.usuario_email - 1].toLowerCase() !== usuario_email.toLowerCase()) {
      return respError('No podés eliminar un vale que no es tuyo.');
    }

    const estadoActual = fila[COL_VALES.estado - 1];
    if (estadoActual !== 'BORRADOR') {
      return respError('Solo se pueden eliminar vales en estado BORRADOR. Estado actual: ' + estadoActual);
    }

    const numFila = i + 1;
    hVales.getRange(numFila, COL_VALES.eliminado,    1, 1).setValue(true);
    hVales.getRange(numFila, COL_VALES.estado,       1, 1).setValue('ELIMINADO');
    hVales.getRange(numFila, COL_VALES.fecha_cierre, 1, 1).setValue(new Date().toISOString());

    SpreadsheetApp.flush();
    return respOk({ id_vale, eliminado: true });
  }

  return respError('Vale no encontrado: ' + id_vale);
}

// ════════════════════════════════════════════════════════════════
// ETAPA 2 — ENDPOINTS DE APROBACIÓN
// ════════════════════════════════════════════════════════════════

// ── ACCIÓN: getPendientesAprobacion ─────────────────────────────
// GET ?accion=getPendientesAprobacion&email=xxx
// Determina el rol del aprobador y devuelve los vales ENVIADOS que le corresponde aprobar.
// SUPERVISOR → vales ENVIADOS de CAPATAZ de su obra
// JEFE_OBRA  → vales ENVIADOS de SUPERVISOR de su obra
function acGetPendientesAprobacion(params) {
  const email = (params.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  const aprobador = buscarUsuario(email);
  if (!aprobador)        return respError('Usuario no registrado.');
  if (!aprobador.activo) return respError('Usuario inactivo.');

  const rolSolicitante = PUEDE_APROBAR[aprobador.rol];
  if (!rolSolicitante) {
    return respError('El rol ' + aprobador.rol + ' no tiene permisos para aprobar vales.');
  }

  // Necesitamos los emails de los usuarios con rol correcto y misma obra
  const hUsuarios = getSheet(HOJA.USUARIOS);
  const usuarios  = sheetToObjects(hUsuarios, COL_USUARIOS);

  const emailsSolicitantes = usuarios
    .filter(u => u.rol === rolSolicitante &&
                 u.obra_codigo === aprobador.obra_codigo &&
                 u.activo)
    .map(u => u.email.toLowerCase());

  if (emailsSolicitantes.length === 0) return respOk([]);

  const hVales = getSheet(HOJA.VALES);
  const vales  = sheetToObjects(hVales, COL_VALES);

  const pendientes = vales.filter(v => {
    if (v.estado !== 'ENVIADO') return false;
    if (v.eliminado === true || v.eliminado === 'TRUE') return false;
    return emailsSolicitantes.includes(v.usuario_email.toLowerCase());
  });

  // Más antiguo primero (ASC)
  pendientes.sort((a, b) => new Date(a.fecha_hora) - new Date(b.fecha_hora));

  // Enriquecer con nombre del solicitante
  const pendientesEnriquecidos = pendientes.map(v => {
    const sol = usuarios.find(u => u.email.toLowerCase() === v.usuario_email.toLowerCase());
    return {
      ...v,
      solicitante_nombre : sol ? sol.nombre : v.usuario_email,
      solicitante_rol    : sol ? sol.rol    : ''
    };
  });

  return respOk(pendientesEnriquecidos);
}

// ── ACCIÓN: aprobarVale ──────────────────────────────────────────
// GET ?accion=aprobarVale&id_vale=xxx&aprobado_por=xxx
function acAprobarVale(params) {
  const { id_vale, aprobado_por } = params;
  if (!id_vale)      return respError('id_vale requerido.');
  if (!aprobado_por) return respError('aprobado_por requerido.');

  const aprobador = buscarUsuario(aprobado_por);
  if (!aprobador)        return respError('Aprobador no registrado.');
  if (!aprobador.activo) return respError('Aprobador inactivo.');

  const rolSolicitante = PUEDE_APROBAR[aprobador.rol];
  if (!rolSolicitante) {
    return respError('El rol ' + aprobador.rol + ' no tiene permisos para aprobar vales.');
  }

  const hVales = getSheet(HOJA.VALES);
  const datos  = hVales.getDataRange().getValues();
  const fila   = buscarFilaVale(datos, id_vale);

  if (fila < 0) return respError('Vale no encontrado: ' + id_vale);

  const filaData = datos[fila - 1];

  // Validar estado
  const estadoActual = filaData[COL_VALES.estado - 1];
  if (estadoActual !== 'ENVIADO') {
    return respError('El vale no está en estado ENVIADO. Estado actual: ' + estadoActual);
  }

  // Validar rol del solicitante
  const emailSolicitante = filaData[COL_VALES.usuario_email - 1].toLowerCase();
  const solicitante      = buscarUsuario(emailSolicitante);
  if (!solicitante || solicitante.rol !== rolSolicitante) {
    return respError('No tenés permiso para aprobar este vale (rol del solicitante: ' +
                     (solicitante ? solicitante.rol : 'desconocido') + ').');
  }

  // Validar misma obra
  const obraVale = filaData[COL_VALES.obra_codigo - 1];
  if (obraVale !== aprobador.obra_codigo) {
    return respError('No podés aprobar vales de otra obra. Tu obra: ' +
                     aprobador.obra_codigo + ' | Obra del vale: ' + obraVale);
  }

  // Actualizar
  const ahora = new Date().toISOString();
  hVales.getRange(fila, COL_VALES.estado,           1, 1).setValue('APROBADO');
  hVales.getRange(fila, COL_VALES.aprobado_por,     1, 1).setValue(aprobado_por);
  hVales.getRange(fila, COL_VALES.fecha_aprobacion, 1, 1).setValue(ahora);

  SpreadsheetApp.flush();
  return respOk({ id_vale, estado: 'APROBADO', aprobado_por, fecha_aprobacion: ahora });
}

// ── ACCIÓN: rechazarVale ─────────────────────────────────────────
// GET ?accion=rechazarVale&id_vale=xxx&rechazado_por=xxx&nota_rechazo=xxx
function acRechazarVale(params) {
  const { id_vale, rechazado_por, nota_rechazo } = params;
  if (!id_vale)       return respError('id_vale requerido.');
  if (!rechazado_por) return respError('rechazado_por requerido.');

  const nota = (nota_rechazo || '').trim();
  if (nota.length < 5) return respError('nota_rechazo obligatoria (mínimo 5 caracteres).');

  const aprobador = buscarUsuario(rechazado_por);
  if (!aprobador)        return respError('Usuario no registrado.');
  if (!aprobador.activo) return respError('Usuario inactivo.');

  const rolSolicitante = PUEDE_APROBAR[aprobador.rol];
  if (!rolSolicitante) {
    return respError('El rol ' + aprobador.rol + ' no tiene permisos para rechazar vales.');
  }

  const hVales = getSheet(HOJA.VALES);
  const datos  = hVales.getDataRange().getValues();
  const fila   = buscarFilaVale(datos, id_vale);

  if (fila < 0) return respError('Vale no encontrado: ' + id_vale);

  const filaData = datos[fila - 1];

  // Validar estado
  const estadoActual = filaData[COL_VALES.estado - 1];
  if (estadoActual !== 'ENVIADO') {
    return respError('El vale no está en estado ENVIADO. Estado actual: ' + estadoActual);
  }

  // Validar rol del solicitante
  const emailSolicitante = filaData[COL_VALES.usuario_email - 1].toLowerCase();
  const solicitante      = buscarUsuario(emailSolicitante);
  if (!solicitante || solicitante.rol !== rolSolicitante) {
    return respError('No tenés permiso para rechazar este vale (rol del solicitante: ' +
                     (solicitante ? solicitante.rol : 'desconocido') + ').');
  }

  // Validar misma obra
  const obraVale = filaData[COL_VALES.obra_codigo - 1];
  if (obraVale !== aprobador.obra_codigo) {
    return respError('No podés rechazar vales de otra obra. Tu obra: ' +
                     aprobador.obra_codigo + ' | Obra del vale: ' + obraVale);
  }

  // Actualizar
  const ahora = new Date().toISOString();
  hVales.getRange(fila, COL_VALES.estado,           1, 1).setValue('RECHAZADO');
  hVales.getRange(fila, COL_VALES.aprobado_por,     1, 1).setValue(rechazado_por); // quien rechaza
  hVales.getRange(fila, COL_VALES.fecha_aprobacion, 1, 1).setValue(ahora);
  hVales.getRange(fila, COL_VALES.nota_cierre,      1, 1).setValue(nota);

  SpreadsheetApp.flush();
  return respOk({ id_vale, estado: 'RECHAZADO', rechazado_por, fecha: ahora, nota_cierre: nota });
}

// ── ACCIÓN: getValesAprobados ────────────────────────────────────
// GET ?accion=getValesAprobados&email=xxx
// Historial de vales aprobados/rechazados por este usuario — últimos 30 días.
function acGetValesAprobados(params) {
  const email = (params.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  const aprobador = buscarUsuario(email);
  if (!aprobador) return respError('Usuario no registrado.');

  const hVales = getSheet(HOJA.VALES);
  const vales  = sheetToObjects(hVales, COL_VALES);

  const hace30 = new Date();
  hace30.setDate(hace30.getDate() - 30);

  const historial = vales.filter(v => {
    if (!['APROBADO', 'RECHAZADO'].includes(v.estado)) return false;
    if (v.eliminado === true || v.eliminado === 'TRUE') return false;
    if ((v.aprobado_por || '').toLowerCase() !== email) return false;
    const fecha = v.fecha_aprobacion ? new Date(v.fecha_aprobacion) : null;
    return fecha && fecha >= hace30;
  });

  historial.sort((a, b) => new Date(b.fecha_aprobacion) - new Date(a.fecha_aprobacion));

  return respOk(historial);
}
