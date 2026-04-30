// ═══════════════════════════════════════════════════════════════
// Vale Digital — Backend GAS v8.0  (Etapa 4 — Admin CRUD)
// Endpoints v1: login, guardarVale, getMisVales, eliminarVale
// Endpoints v2: getPendientesAprobacion, aprobarVale,
//               rechazarVale, getValesAprobados
// Endpoints v3: getValesPorGestionar, getValesPendientes, getEntregadosHoy,
//               iniciarPreparacion, entregarVale,
//               entregaParcialVale, cancelarVale, getTodosVales
// Endpoints v4: getUsuarios, crearUsuario, editarUsuario, toggleUsuario,
//               getObras, crearObra, editarObra, toggleObra
// ═══════════════════════════════════════════════════════════════

// ── CONFIGURACIÓN ───────────────────────────────────────────────
const SPREADSHEET_ID = '10XkTAarQdgucz8WIwNh5FhV2vyW9qdoRhRCKnQ7wb4k'; // ← ID del Sheet

const HOJA = {
  VALES    : 'VALES',
  USUARIOS : 'USUARIOS',
  OBRAS    : 'OBRAS'
};

// Roles con acceso al panel Admin
const ROLES_ADMIN = ['ADMIN'];

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
  email          : 1,
  nombre         : 2,
  rol            : 3,
  obra_codigo    : 4,
  activo         : 5,
  superior_email : 6
};

const COL_OBRAS = {
  codigo       : 1,
  descripcion  : 2,
  email_pañol  : 3,
  email_almacen: 4,
  activa       : 5
};

// Jerarquia de aprobacion por rol (fallback cuando superior_email esta vacio)
// SUPERVISOR aprueba vales de CAPATAZ
// JEFE_OBRA  aprueba vales de SUPERVISOR o CAPATAZ
const PUEDE_APROBAR = {
  'SUPERVISOR': ['CAPATAZ'],
  'JEFE_OBRA' : ['SUPERVISOR', 'CAPATAZ']
};

// Helper: dado un solicitante, retorna el email de su aprobador.
// Prioridad: superior_email explicito > logica por rol.
// Retorna null si el solicitante no tiene aprobador valido.
function obtenerEmailSuperior(solicitante) {
  if (solicitante.superior_email && solicitante.superior_email.trim() !== '') {
    return solicitante.superior_email.trim().toLowerCase();
  }
  return null; // usa logica por rol
}

// Helper: verifica si un aprobador puede aprobar al solicitante dado.
// Respeta superior_email si esta definido; sino usa PUEDE_APROBAR por rol.
function puedeAprobarA(aprobador, solicitante) {
  if (!aprobador || !solicitante) return false;
  if (aprobador.obra_codigo !== solicitante.obra_codigo) return false;
  // Si el solicitante tiene superior explicito
  const superiorExplicito = obtenerEmailSuperior(solicitante);
  if (superiorExplicito) {
    return aprobador.email.toLowerCase() === superiorExplicito;
  }
  // Fallback: logica por rol
  const rolesAprobables = PUEDE_APROBAR[aprobador.rol] || [];
  return rolesAprobables.includes(solicitante.rol);
}

// Destino que gestiona cada rol de almacén
const DESTINO_ROL = {
  'ALMACENERO': 'ALMACEN',
  'PAÑOLERO'  : 'PAÑOL'
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
      // ── Etapa 3 ──
      case 'getValesPorGestionar'     : return acGetValesPorGestionar(e.parameter);
      case 'getValesPendientes'       : return acGetValesPendientes(e.parameter);
      case 'iniciarPreparacion'       : return acIniciarPreparacion(e.parameter);
      case 'entregarVale'             : return acEntregarVale(e.parameter);
      case 'entregaParcialVale'       : return acEntregaParcialVale(e.parameter);
      case 'cancelarVale'             : return acCancelarVale(e.parameter);
      case 'getEntregadosHoy'         : return acGetEntregadosHoy(e.parameter);
      case 'getTodosVales'            : return acGetTodosVales(e.parameter);
      // ── Etapa 4 — Admin ──
      case 'getUsuarios'              : return acGetUsuarios(e.parameter);
      case 'crearUsuario'             : return acCrearUsuario(e.parameter);
      case 'editarUsuario'            : return acEditarUsuario(e.parameter);
      case 'toggleUsuario'            : return acToggleUsuario(e.parameter);
      case 'getObras'                 : return acGetObras(e.parameter);
      case 'crearObra'                : return acCrearObra(e.parameter);
      case 'editarObra'               : return acEditarObra(e.parameter);
      case 'toggleObra'               : return acToggleObra(e.parameter);
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

  const rolesAprobables = PUEDE_APROBAR[aprobador.rol];
  if (!rolesAprobables) {
    return respError('El rol ' + aprobador.rol + ' no tiene permisos para aprobar vales.');
  }

  const hUsuarios = getSheet(HOJA.USUARIOS);
  const usuarios  = sheetToObjects(hUsuarios, COL_USUARIOS);

  // Solicitantes que este aprobador puede aprobar (por superior explicito o por rol)
  const solicitantesAprobables = usuarios.filter(u =>
    u.activo && puedeAprobarA(aprobador, u)
  );
  const emailsSolicitantes = solicitantesAprobables.map(u => u.email.toLowerCase());

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

  const rolesAprobables = PUEDE_APROBAR[aprobador.rol];
  if (!rolesAprobables) {
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

  // Validar permiso usando superior_email o rol
  const emailSolicitante = filaData[COL_VALES.usuario_email - 1].toLowerCase();
  const solicitante      = buscarUsuario(emailSolicitante);
  if (!solicitante || !puedeAprobarA(aprobador, solicitante)) {
    return respError('No tenés permiso para aprobar este vale.');
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

  const rolesAprobables = PUEDE_APROBAR[aprobador.rol];
  if (!rolesAprobables) {
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

  // Validar permiso usando superior_email o rol
  const emailSolicitante = filaData[COL_VALES.usuario_email - 1].toLowerCase();
  const solicitante      = buscarUsuario(emailSolicitante);
  if (!solicitante || !puedeAprobarA(aprobador, solicitante)) {
    return respError('No tenés permiso para rechazar este vale.');
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

// ════════════════════════════════════════════════════════════════
// ETAPA 3 — ENDPOINTS ALMACÉN / PAÑOL
// ════════════════════════════════════════════════════════════════

// ── Helper: validar gestor ────────────────────────────────────────
// Retorna { ok, gestor, destino } o { ok: false, error }
function validarGestor(email) {
  const gestor = buscarUsuario(email);
  if (!gestor)        return { ok: false, error: 'Usuario no registrado.' };
  if (!gestor.activo) return { ok: false, error: 'Usuario inactivo.' };
  const destino = DESTINO_ROL[gestor.rol];
  if (!destino) return { ok: false, error: 'El rol ' + gestor.rol + ' no gestiona vales de almacén/pañol.' };
  return { ok: true, gestor, destino };
}

// ── Helper: enriquecer vale con nombre solicitante ────────────────
function enriquecerVales(vales, usuarios) {
  return vales.map(v => {
    const sol = usuarios.find(u => u.email.toLowerCase() === v.usuario_email.toLowerCase());
    return {
      ...v,
      solicitante_nombre: sol ? sol.nombre : v.usuario_email,
      solicitante_rol   : sol ? sol.rol    : ''
    };
  });
}

// ── ACCIÓN: getValesPorGestionar ─────────────────────────────────
// GET ?accion=getValesPorGestionar&email=xxx
// Retorna vales en estado APROBADO del destino del gestor, su obra.
// Ordenados: más antiguo primero (prioridad de cola).
function acGetValesPorGestionar(params) {
  const email = (params.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  const v = validarGestor(email);
  if (!v.ok) return respError(v.error);
  const { gestor, destino } = v;

  const hVales   = getSheet(HOJA.VALES);
  const hUsuarios = getSheet(HOJA.USUARIOS);
  const vales    = sheetToObjects(hVales, COL_VALES);
  const usuarios = sheetToObjects(hUsuarios, COL_USUARIOS);

  const cola = vales.filter(vale =>
    vale.estado === 'APROBADO' &&
    vale.destino === destino &&
    vale.obra_codigo === gestor.obra_codigo &&
    vale.eliminado !== true && vale.eliminado !== 'TRUE'
  );

  cola.sort((a, b) => new Date(a.fecha_aprobacion || a.fecha_hora) - new Date(b.fecha_aprobacion || b.fecha_hora));

  return respOk(enriquecerVales(cola, usuarios));
}

// ── ACCIÓN: getValesPendientes ────────────────────────────────────
// GET ?accion=getValesPendientes&email=xxx
// Retorna vales en estado PENDIENTE del destino del gestor, su obra.
// Ordenados por solicitante_nombre ASC para facilitar búsqueda presencial.
function acGetValesPendientes(params) {
  const email = (params.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  const v = validarGestor(email);
  if (!v.ok) return respError(v.error);
  const { gestor, destino } = v;

  const hVales    = getSheet(HOJA.VALES);
  const hUsuarios = getSheet(HOJA.USUARIOS);
  const vales     = sheetToObjects(hVales, COL_VALES);
  const usuarios  = sheetToObjects(hUsuarios, COL_USUARIOS);

  const enPrep = vales.filter(vale =>
    vale.estado === 'PENDIENTE' &&
    vale.destino === destino &&
    vale.obra_codigo === gestor.obra_codigo &&
    vale.eliminado !== true && vale.eliminado !== 'TRUE'
  );

  const enriquecidos = enriquecerVales(enPrep, usuarios);
  enriquecidos.sort((a, b) =>
    (a.solicitante_nombre || '').localeCompare(b.solicitante_nombre || '', 'es')
  );

  return respOk(enriquecidos);
}

// ── ACCIÓN: iniciarPreparacion ────────────────────────────────────
// GET ?accion=iniciarPreparacion&id_vale=xxx&gestionado_por=xxx
// APROBADO → PENDIENTE
function acIniciarPreparacion(params) {
  const { id_vale, gestionado_por } = params;
  if (!id_vale)        return respError('id_vale requerido.');
  if (!gestionado_por) return respError('gestionado_por requerido.');

  const v = validarGestor(gestionado_por);
  if (!v.ok) return respError(v.error);
  const { gestor, destino } = v;

  const hVales = getSheet(HOJA.VALES);
  const datos  = hVales.getDataRange().getValues();
  const fila   = buscarFilaVale(datos, id_vale);
  if (fila < 0) return respError('Vale no encontrado: ' + id_vale);

  const filaData = datos[fila - 1];
  const estado   = filaData[COL_VALES.estado - 1];
  if (estado !== 'APROBADO') return respError('El vale debe estar en estado APROBADO. Estado actual: ' + estado);

  const valeDestino = filaData[COL_VALES.destino - 1];
  if (valeDestino !== destino) return respError('Este vale es para ' + valeDestino + ', no para ' + destino + '.');

  const obraVale = filaData[COL_VALES.obra_codigo - 1];
  if (obraVale !== gestor.obra_codigo) return respError('El vale pertenece a otra obra.');

  const ahora = new Date().toISOString();
  hVales.getRange(fila, COL_VALES.estado,         1, 1).setValue('PENDIENTE');
  hVales.getRange(fila, COL_VALES.gestionado_por, 1, 1).setValue(gestionado_por);
  hVales.getRange(fila, COL_VALES.fecha_cierre,   1, 1).setValue(ahora); // fecha inicio preparación

  SpreadsheetApp.flush();
  return respOk({ id_vale, estado: 'PENDIENTE', gestionado_por, fecha: ahora });
}

// ── Helper compartido para cerrar un vale (ENTREGADO / ENTREGA_PARCIAL / CANCELADO) ──
function acCerrarVale(params, estadoDestino, notaObligatoria) {
  const { id_vale, gestionado_por, nota_cierre } = params;
  if (!id_vale)        return respError('id_vale requerido.');
  if (!gestionado_por) return respError('gestionado_por requerido.');

  const nota = (nota_cierre || '').trim();
  if (notaObligatoria && nota.length < 5) {
    return respError('nota_cierre obligatoria (mínimo 5 caracteres).');
  }

  const v = validarGestor(gestionado_por);
  if (!v.ok) return respError(v.error);
  const { gestor, destino } = v;

  const hVales = getSheet(HOJA.VALES);
  const datos  = hVales.getDataRange().getValues();
  const fila   = buscarFilaVale(datos, id_vale);
  if (fila < 0) return respError('Vale no encontrado: ' + id_vale);

  const filaData    = datos[fila - 1];
  const estadoActual = filaData[COL_VALES.estado - 1];

  // CANCELADO puede venir desde APROBADO o PENDIENTE
  // ENTREGADO y ENTREGA_PARCIAL solo desde PENDIENTE
  const estadosValidos = estadoDestino === 'CANCELADO'
    ? ['APROBADO', 'PENDIENTE']
    : ['PENDIENTE'];

  if (!estadosValidos.includes(estadoActual)) {
    return respError('Estado inválido para esta acción. Estado actual: ' + estadoActual);
  }

  const valeDestino = filaData[COL_VALES.destino - 1];
  if (valeDestino !== destino) return respError('Este vale es para ' + valeDestino + ', no para ' + destino + '.');

  const obraVale = filaData[COL_VALES.obra_codigo - 1];
  if (obraVale !== gestor.obra_codigo) return respError('El vale pertenece a otra obra.');

  const ahora = new Date().toISOString();
  hVales.getRange(fila, COL_VALES.estado,         1, 1).setValue(estadoDestino);
  hVales.getRange(fila, COL_VALES.gestionado_por, 1, 1).setValue(gestionado_por);
  hVales.getRange(fila, COL_VALES.fecha_cierre,   1, 1).setValue(ahora);
  if (nota) hVales.getRange(fila, COL_VALES.nota_cierre, 1, 1).setValue(nota);

  SpreadsheetApp.flush();
  return respOk({ id_vale, estado: estadoDestino, gestionado_por, fecha: ahora, nota_cierre: nota });
}

// ── ACCIÓN: entregarVale ──────────────────────────────────────────
// GET ?accion=entregarVale&id_vale=xxx&gestionado_por=xxx&nota_cierre=xxx (opcional)
// PENDIENTE → ENTREGADO
function acEntregarVale(params) {
  return acCerrarVale(params, 'ENTREGADO', false);
}

// ── ACCIÓN: entregaParcialVale ────────────────────────────────────
// GET ?accion=entregaParcialVale&id_vale=xxx&gestionado_por=xxx&nota_cierre=xxx (obligatorio)
// PENDIENTE → ENTREGA_PARCIAL
function acEntregaParcialVale(params) {
  return acCerrarVale(params, 'ENTREGA_PARCIAL', true);
}

// ── ACCIÓN: cancelarVale ──────────────────────────────────────────
// GET ?accion=cancelarVale&id_vale=xxx&gestionado_por=xxx&nota_cierre=xxx (obligatorio)
// APROBADO o PENDIENTE → CANCELADO
function acCancelarVale(params) {
  return acCerrarVale(params, 'CANCELADO', true);
}

// ── ACCIÓN: getTodosVales ─────────────────────────────────────────
// GET ?accion=getTodosVales&email=xxx&solicitante=xxx (opcional)
// Roles permitidos:
//   JEFE_OBRA    — todos los vales de su obra (últimos 30 días)
//   ALMACENERO   — idem JEFE_OBRA (todos los destinos, su obra)
//   PAÑOLERO     — idem JEFE_OBRA (todos los destinos, su obra)
//   SUPERVISOR   — solo vales que él aprobó (aprobado_por = su email)
//   ADMIN        — todos los vales de todas las obras (últimos 30 días)
function acGetTodosVales(params) {
  const email = (params.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  const usuario = buscarUsuario(email);
  if (!usuario)        return respError('Usuario no registrado.');
  if (!usuario.activo) return respError('Usuario inactivo.');

  const rolesPermitidos = ['JEFE_OBRA', 'ALMACENERO', 'PAÑOLERO', 'SUPERVISOR', 'ADMIN'];
  if (!rolesPermitidos.includes(usuario.rol)) return respError('Sin permiso para consultar todos los vales.');

  const hVales    = getSheet(HOJA.VALES);
  const hUsuarios = getSheet(HOJA.USUARIOS);
  const vales     = sheetToObjects(hVales, COL_VALES);
  const usuarios  = sheetToObjects(hUsuarios, COL_USUARIOS);

  const hace30 = new Date();
  hace30.setDate(hace30.getDate() - 30);

  const filtroSolicitante = (params.solicitante || '').trim().toLowerCase();
  const esSupervisor = usuario.rol === 'SUPERVISOR';
  const esAdmin      = usuario.rol === 'ADMIN';

  let resultado = vales.filter(v => {
    if (v.eliminado === true || v.eliminado === 'TRUE') return false;
    const fecha = new Date(v.fecha_hora);
    if (fecha < hace30) return false;
    // ADMIN: todas las obras
    if (esAdmin) return true;
    // Resto: solo su obra
    if (v.obra_codigo !== usuario.obra_codigo) return false;
    // SUPERVISOR: solo ve los que él aprobó
    if (esSupervisor) {
      return (v.aprobado_por || '').toLowerCase() === email;
    }
    return true;
  });

  const enriquecidos = enriquecerVales(resultado, usuarios);

  // Filtro por solicitante (nombre o email, búsqueda parcial case-insensitive)
  const filtrado = filtroSolicitante
    ? enriquecidos.filter(v =>
        (v.solicitante_nombre || '').toLowerCase().includes(filtroSolicitante) ||
        (v.usuario_email      || '').toLowerCase().includes(filtroSolicitante)
      )
    : enriquecidos;

  // Orden: más reciente primero
  filtrado.sort((a, b) => new Date(b.fecha_hora) - new Date(a.fecha_hora));

  return respOk(filtrado);
}

// ── ACCIÓN: getEntregadosHoy ─────────────────────────────────────────
// GET ?accion=getEntregadosHoy&email=xxx
// Retorna la cantidad de vales ENTREGADO o ENTREGA_PARCIAL de hoy del gestor.
function acGetEntregadosHoy(params) {
  const email = (params.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  const v = validarGestor(email);
  if (!v.ok) return respError(v.error);
  const { gestor, destino } = v;

  const hVales = getSheet(HOJA.VALES);
  const vales  = sheetToObjects(hVales, COL_VALES);

  const hoy = new Date();
  const inicioHoy = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());

  const entregados = vales.filter(vale => {
    if (vale.eliminado === true || vale.eliminado === 'TRUE') return false;
    if (!['ENTREGADO','ENTREGA_PARCIAL'].includes(vale.estado)) return false;
    if (vale.destino !== destino) return false;
    if (vale.obra_codigo !== gestor.obra_codigo) return false;
    const fechaCierre = vale.fecha_cierre ? new Date(vale.fecha_cierre) : null;
    return fechaCierre && fechaCierre >= inicioHoy;
  });

  return respOk({ cantidad: entregados.length });
}

// ════════════════════════════════════════════════════════════════
// ETAPA 4 — ADMIN CRUD USUARIOS Y OBRAS
// ════════════════════════════════════════════════════════════════

// ── Helper: validar admin ────────────────────────────────────────
// Retorna { ok: true, admin } o { ok: false, error }
function validarAdmin(email_admin) {
  const email = (email_admin || '').trim().toLowerCase();
  if (!email) return { ok: false, error: 'email_admin requerido.' };
  const admin = buscarUsuario(email);
  if (!admin)        return { ok: false, error: 'Usuario no registrado.' };
  if (!admin.activo) return { ok: false, error: 'Usuario inactivo.' };
  if (!ROLES_ADMIN.includes(admin.rol)) return { ok: false, error: 'Sin permiso: se requiere rol ADMIN.' };
  return { ok: true, admin };
}

// ── Helper: buscar fila en hoja por valor en columna ────────────
// Retorna número de fila (base 1, incluye header) o -1
function buscarFilaEnHoja(hoja, columna, valor) {
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][columna - 1]).toLowerCase() === String(valor).toLowerCase()) return i + 1;
  }
  return -1;
}

// ════════════════════════════════════════════════════════════════
// USUARIOS — CRUD
// ════════════════════════════════════════════════════════════════

// ── ACCIÓN: getUsuarios ──────────────────────────────────────────
// GET ?accion=getUsuarios&email_admin=xxx
// Retorna todos los usuarios (todas las obras), sin filtro de activo.
function acGetUsuarios(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const hUsuarios = getSheet(HOJA.USUARIOS);
  const usuarios  = sheetToObjects(hUsuarios, COL_USUARIOS);

  // Ordenar: por obra_codigo ASC, luego nombre ASC
  usuarios.sort((a, b) => {
    const obra = String(a.obra_codigo).localeCompare(String(b.obra_codigo), 'es');
    if (obra !== 0) return obra;
    return String(a.nombre).localeCompare(String(b.nombre), 'es');
  });

  return respOk(usuarios);
}

// ── ACCIÓN: crearUsuario ─────────────────────────────────────────
// GET ?accion=crearUsuario&email_admin=xxx&email=xxx&nombre=xxx&rol=xxx
//     &obra_codigo=xxx&activo=TRUE&superior_email=xxx
// Valida que el email no exista ya en USUARIOS.
function acCrearUsuario(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const email          = (params.email          || '').trim().toLowerCase();
  const nombre         = (params.nombre         || '').trim();
  const rol            = (params.rol            || '').trim().toUpperCase();
  const obra_codigo    = (params.obra_codigo    || '').trim().toUpperCase();
  const activo         = params.activo !== 'FALSE'; // default TRUE
  const superior_email = (params.superior_email || '').trim().toLowerCase();

  if (!email)       return respError('email requerido.');
  if (!nombre)      return respError('nombre requerido.');
  if (!rol)         return respError('rol requerido.');
  if (!obra_codigo) return respError('obra_codigo requerido.');

  const ROLES_VALIDOS = ['CAPATAZ', 'SUPERVISOR', 'JEFE_OBRA', 'ALMACENERO', 'PAÑOLERO', 'ADMIN'];
  if (!ROLES_VALIDOS.includes(rol)) return respError('Rol inválido: ' + rol);

  // Validar que la obra exista y esté activa
  const hObras = getSheet(HOJA.OBRAS);
  const obras  = sheetToObjects(hObras, COL_OBRAS);
  const obra   = obras.find(o => String(o.codigo).toUpperCase() === obra_codigo);
  if (!obra)        return respError('Obra no encontrada: ' + obra_codigo);
  if (!obra.activa) return respError('La obra ' + obra_codigo + ' está inactiva.');

  // Validar que el email no exista
  const hUsuarios = getSheet(HOJA.USUARIOS);
  const filaExistente = buscarFilaEnHoja(hUsuarios, COL_USUARIOS.email, email);
  if (filaExistente > 0) return respError('Ya existe un usuario con ese email: ' + email);

  // Validar superior_email si se provee
  if (superior_email) {
    const superior = buscarUsuario(superior_email);
    if (!superior) return respError('superior_email no encontrado: ' + superior_email);
  }

  hUsuarios.appendRow([email, nombre, rol, obra_codigo, activo, superior_email]);
  SpreadsheetApp.flush();

  return respOk({ email, nombre, rol, obra_codigo, activo, superior_email, accion: 'created' });
}

// ── ACCIÓN: editarUsuario ────────────────────────────────────────
// GET ?accion=editarUsuario&email_admin=xxx&email_usuario=xxx
//     &nombre=xxx&rol=xxx&obra_codigo=xxx&superior_email=xxx
// No permite cambiar el email (es la PK). No cambia activo (usar toggleUsuario).
function acEditarUsuario(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const email_usuario  = (params.email_usuario  || '').trim().toLowerCase();
  if (!email_usuario) return respError('email_usuario requerido.');

  const hUsuarios = getSheet(HOJA.USUARIOS);
  const fila = buscarFilaEnHoja(hUsuarios, COL_USUARIOS.email, email_usuario);
  if (fila < 0) return respError('Usuario no encontrado: ' + email_usuario);

  // Leer valores actuales
  const datos        = hUsuarios.getDataRange().getValues();
  const filaActual   = datos[fila - 1];
  const nombreActual        = filaActual[COL_USUARIOS.nombre         - 1];
  const rolActual           = filaActual[COL_USUARIOS.rol            - 1];
  const obraActual          = filaActual[COL_USUARIOS.obra_codigo    - 1];
  const superiorActual      = filaActual[COL_USUARIOS.superior_email - 1];

  const nombre         = (params.nombre         || '').trim()            || nombreActual;
  const rol            = (params.rol            || '').trim().toUpperCase() || rolActual;
  const obra_codigo    = (params.obra_codigo    || '').trim().toUpperCase() || obraActual;
  // superior_email puede borrarse enviando ''
  const superior_email = params.superior_email !== undefined
    ? (params.superior_email || '').trim().toLowerCase()
    : superiorActual;

  const ROLES_VALIDOS = ['CAPATAZ', 'SUPERVISOR', 'JEFE_OBRA', 'ALMACENERO', 'PAÑOLERO', 'ADMIN'];
  if (!ROLES_VALIDOS.includes(rol)) return respError('Rol inválido: ' + rol);

  // Validar obra
  const hObras = getSheet(HOJA.OBRAS);
  const obras  = sheetToObjects(hObras, COL_OBRAS);
  const obra   = obras.find(o => String(o.codigo).toUpperCase() === obra_codigo);
  if (!obra)        return respError('Obra no encontrada: ' + obra_codigo);
  if (!obra.activa) return respError('La obra ' + obra_codigo + ' está inactiva.');

  // Validar superior_email si se provee
  if (superior_email) {
    const superior = buscarUsuario(superior_email);
    if (!superior) return respError('superior_email no encontrado: ' + superior_email);
  }

  hUsuarios.getRange(fila, COL_USUARIOS.nombre,         1, 1).setValue(nombre);
  hUsuarios.getRange(fila, COL_USUARIOS.rol,            1, 1).setValue(rol);
  hUsuarios.getRange(fila, COL_USUARIOS.obra_codigo,    1, 1).setValue(obra_codigo);
  hUsuarios.getRange(fila, COL_USUARIOS.superior_email, 1, 1).setValue(superior_email);
  SpreadsheetApp.flush();

  return respOk({ email: email_usuario, nombre, rol, obra_codigo, superior_email, accion: 'updated' });
}

// ── ACCIÓN: toggleUsuario ────────────────────────────────────────
// GET ?accion=toggleUsuario&email_admin=xxx&email_usuario=xxx
// Invierte el campo activo. No permite desactivar al propio admin.
function acToggleUsuario(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const email_admin   = (params.email_admin   || '').trim().toLowerCase();
  const email_usuario = (params.email_usuario || '').trim().toLowerCase();
  if (!email_usuario) return respError('email_usuario requerido.');

  if (email_usuario === email_admin) return respError('No podés desactivarte a vos mismo.');

  const hUsuarios = getSheet(HOJA.USUARIOS);
  const fila = buscarFilaEnHoja(hUsuarios, COL_USUARIOS.email, email_usuario);
  if (fila < 0) return respError('Usuario no encontrado: ' + email_usuario);

  const datos       = hUsuarios.getDataRange().getValues();
  const activoActual = datos[fila - 1][COL_USUARIOS.activo - 1];
  const nuevoActivo  = !(activoActual === true || activoActual === 'TRUE');

  hUsuarios.getRange(fila, COL_USUARIOS.activo, 1, 1).setValue(nuevoActivo);
  SpreadsheetApp.flush();

  return respOk({ email: email_usuario, activo: nuevoActivo });
}

// ════════════════════════════════════════════════════════════════
// OBRAS — CRUD
// ════════════════════════════════════════════════════════════════

// ── ACCIÓN: getObras ─────────────────────────────────────────────
// GET ?accion=getObras&email_admin=xxx
// Retorna todas las obras (activas e inactivas).
function acGetObras(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const hObras = getSheet(HOJA.OBRAS);
  const obras  = sheetToObjects(hObras, COL_OBRAS);

  obras.sort((a, b) => String(a.codigo).localeCompare(String(b.codigo), 'es'));

  return respOk(obras);
}

// ── ACCIÓN: crearObra ────────────────────────────────────────────
// GET ?accion=crearObra&email_admin=xxx&codigo=xxx&descripcion=xxx
//     &email_pañol=xxx&email_almacen=xxx&activa=TRUE
function acCrearObra(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const codigo        = (params.codigo        || '').trim().toUpperCase();
  const descripcion   = (params.descripcion   || '').trim();
  const email_pañol   = (params.email_pañol   || '').trim().toLowerCase();
  const email_almacen = (params.email_almacen || '').trim().toLowerCase();
  const activa        = params.activa !== 'FALSE'; // default TRUE

  if (!codigo)      return respError('codigo requerido.');
  if (!descripcion) return respError('descripcion requerida.');

  const hObras = getSheet(HOJA.OBRAS);
  const filaExistente = buscarFilaEnHoja(hObras, COL_OBRAS.codigo, codigo);
  if (filaExistente > 0) return respError('Ya existe una obra con ese código: ' + codigo);

  hObras.appendRow([codigo, descripcion, email_pañol, email_almacen, activa]);
  SpreadsheetApp.flush();

  return respOk({ codigo, descripcion, email_pañol, email_almacen, activa, accion: 'created' });
}

// ── ACCIÓN: editarObra ───────────────────────────────────────────
// GET ?accion=editarObra&email_admin=xxx&codigo=xxx
//     &descripcion=xxx&email_pañol=xxx&email_almacen=xxx
// No permite cambiar el código (es la PK). No cambia activa (usar toggleObra).
function acEditarObra(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const codigo = (params.codigo || '').trim().toUpperCase();
  if (!codigo) return respError('codigo requerido.');

  const hObras = getSheet(HOJA.OBRAS);
  const fila   = buscarFilaEnHoja(hObras, COL_OBRAS.codigo, codigo);
  if (fila < 0) return respError('Obra no encontrada: ' + codigo);

  const datos         = hObras.getDataRange().getValues();
  const filaActual    = datos[fila - 1];
  const descActual    = filaActual[COL_OBRAS.descripcion   - 1];
  const pañolActual   = filaActual[COL_OBRAS.email_pañol   - 1];
  const almacenActual = filaActual[COL_OBRAS.email_almacen - 1];

  const descripcion   = (params.descripcion   || '').trim()              || descActual;
  const email_pañol   = params.email_pañol   !== undefined
    ? (params.email_pañol   || '').trim().toLowerCase() : pañolActual;
  const email_almacen = params.email_almacen !== undefined
    ? (params.email_almacen || '').trim().toLowerCase() : almacenActual;

  if (!descripcion) return respError('descripcion requerida.');

  hObras.getRange(fila, COL_OBRAS.descripcion,   1, 1).setValue(descripcion);
  hObras.getRange(fila, COL_OBRAS.email_pañol,   1, 1).setValue(email_pañol);
  hObras.getRange(fila, COL_OBRAS.email_almacen, 1, 1).setValue(email_almacen);
  SpreadsheetApp.flush();

  return respOk({ codigo, descripcion, email_pañol, email_almacen, accion: 'updated' });
}

// ── ACCIÓN: toggleObra ───────────────────────────────────────────
// GET ?accion=toggleObra&email_admin=xxx&codigo=xxx
// Invierte el campo activa.
function acToggleObra(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const codigo = (params.codigo || '').trim().toUpperCase();
  if (!codigo) return respError('codigo requerido.');

  const hObras = getSheet(HOJA.OBRAS);
  const fila   = buscarFilaEnHoja(hObras, COL_OBRAS.codigo, codigo);
  if (fila < 0) return respError('Obra no encontrada: ' + codigo);

  const datos      = hObras.getDataRange().getValues();
  const activaActual = datos[fila - 1][COL_OBRAS.activa - 1];
  const nuevaActiva  = !(activaActual === true || activaActual === 'TRUE');

  hObras.getRange(fila, COL_OBRAS.activa, 1, 1).setValue(nuevaActiva);
  SpreadsheetApp.flush();

  return respOk({ codigo, activa: nuevaActiva });
}
