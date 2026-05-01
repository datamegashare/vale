// ═══════════════════════════════════════════════════════════════
// Vale Digital — Backend GAS v10.0
// Endpoints v1: login, guardarVale, getMisVales, eliminarVale
// Endpoints v2: getPendientesAprobacion, aprobarVale,
//               rechazarVale, getValesAprobados
// Endpoints v3: getValesPorGestionar, getValesPendientes, getEntregadosHoy,
//               iniciarPreparacion, entregarVale,
//               entregaParcialVale, cancelarVale, getTodosVales
// Endpoints v4: getUsuarios, crearUsuario, editarUsuario, toggleUsuario,
//               getObras, crearObra, editarObra, toggleObra
// Endpoints v5: getNotificaciones, marcarLeidas
// Endpoints v6: getContexto (reemplaza múltiples calls de carga inicial)
//               getHistorialVale, registrarTransicion (interno)
// ═══════════════════════════════════════════════════════════════
// CAMBIOS v10 respecto a v9:
//   1. Nueva hoja HISTORIAL_ESTADOS — registra cada transición de estado
//      con timestamp, actor y nota. Creada automáticamente si no existe.
//   2. registrarTransicion() — helper interno llamado en cada cambio de estado:
//      guardarVale (creación, envío, auto-aprobación), aprobarVale,
//      rechazarVale, eliminarVale, iniciarPreparacion, acCerrarVale.
//   3. getContexto — nuevo endpoint que reemplaza las 5-7 calls paralelas
//      de carga inicial. Lee cada hoja UNA SOLA VEZ y devuelve todo el
//      contexto del usuario según su rol en un único JSON.
//   4. getHistorialVale — nuevo endpoint. Devuelve el historial de
//      transiciones de un vale específico, bajo demanda (al abrir detalle).
//   5. acGuardarVale ahora recibe estado_anterior para registrar correctamente
//      la transición BORRADOR→ENVIADO cuando el usuario envía un borrador.
// ═══════════════════════════════════════════════════════════════

// ── CONFIGURACIÓN ───────────────────────────────────────────────
const SPREADSHEET_ID = '10XkTAarQdgucz8WIwNh5FhV2vyW9qdoRhRCKnQ7wb4k';

const HOJA = {
  VALES             : 'VALES',
  USUARIOS          : 'USUARIOS',
  OBRAS             : 'OBRAS',
  NOTIFICACIONES    : 'NOTIFICACIONES',
  HISTORIAL_ESTADOS : 'HISTORIAL_ESTADOS'
};

// Columnas hoja NOTIFICACIONES
const COL_NOTIF = {
  email_destino : 1,
  mensaje       : 2,
  fecha         : 3,
  leido         : 4,
  id_vale       : 5
};

// Columnas hoja HISTORIAL_ESTADOS
const COL_HIST = {
  id_vale        : 1,
  estado_anterior: 2,
  estado_nuevo   : 3,
  usuario_email  : 4,
  usuario_nombre : 5,
  fecha          : 6,
  nota           : 7
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
  email_panol  : 3,
  email_almacen: 4,
  activa       : 5
};

// Jerarquia de aprobacion por rol
const PUEDE_APROBAR = {
  'SUPERVISOR': ['CAPATAZ'],
  'JEFE_OBRA' : ['SUPERVISOR', 'CAPATAZ']
};

// Destino que gestiona cada rol de almacén
const DESTINO_ROL = {
  'ALMACENERO': 'ALMACEN',
  'PAÑOLERO'  : 'PAÑOL'
};

// ── HELPERS GENERALES ────────────────────────────────────────────

function getSheet(nombre) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(nombre);
}

// Abre el spreadsheet UNA vez y retorna referencia. Usar en getContexto
// para evitar múltiples openById en la misma ejecución.
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
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

// Convierte filas de una hoja a array de objetos usando un mapa de columnas.
// Recibe el resultado de hoja.getDataRange().getValues() ya leído (rawData)
// para evitar lecturas repetidas.
function rawToObjects(rawData, colMap) {
  if (!rawData || rawData.length < 2) return [];
  const headers = Object.keys(colMap);
  return rawData.slice(1).map(fila => {
    const obj = {};
    headers.forEach(key => { obj[key] = fila[colMap[key] - 1]; });
    return obj;
  });
}

// Versión legacy que lee la hoja por su cuenta. Solo usar en endpoints
// que NO pasan por getContexto (mutaciones puntuales).
function sheetToObjects(hoja, colMap) {
  const datos = hoja.getDataRange().getValues();
  return rawToObjects(datos, colMap);
}

// Devuelve el objeto usuario o null. Recibe array ya cargado para evitar
// releer USUARIOS en cada llamada.
function buscarUsuarioEnArray(usuarios, email) {
  return usuarios.find(u => u.email.toLowerCase() === email.toLowerCase()) || null;
}

// Versión legacy que lee USUARIOS desde Sheets. Solo para endpoints que
// no reciben el array de usuarios ya cargado.
function buscarUsuario(email) {
  const hUsuarios = getSheet(HOJA.USUARIOS);
  const usuarios  = sheetToObjects(hUsuarios, COL_USUARIOS);
  return buscarUsuarioEnArray(usuarios, email);
}

// Encuentra la fila (base 1, incluye header) de un vale por id_vale.
// Retorna -1 si no existe.
function buscarFilaVale(datos, id_vale) {
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][COL_VALES.id_vale - 1] === id_vale) return i + 1;
  }
  return -1;
}

// Helper: dado un solicitante, retorna el email de su aprobador.
function obtenerEmailSuperior(solicitante) {
  if (solicitante.superior_email && solicitante.superior_email.trim() !== '') {
    return solicitante.superior_email.trim().toLowerCase();
  }
  return null;
}

// Helper: verifica si un aprobador puede aprobar al solicitante dado.
function puedeAprobarA(aprobador, solicitante) {
  if (!aprobador || !solicitante) return false;
  if (aprobador.obra_codigo !== solicitante.obra_codigo) return false;
  const superiorExplicito = obtenerEmailSuperior(solicitante);
  if (superiorExplicito) {
    return aprobador.email.toLowerCase() === superiorExplicito;
  }
  const rolesAprobables = PUEDE_APROBAR[aprobador.rol] || [];
  return rolesAprobables.includes(solicitante.rol);
}

// Helper: validar gestor de almacén/pañol.
function validarGestor(email) {
  const gestor = buscarUsuario(email);
  if (!gestor)        return { ok: false, error: 'Usuario no registrado.' };
  if (!gestor.activo) return { ok: false, error: 'Usuario inactivo.' };
  const destino = DESTINO_ROL[gestor.rol];
  if (!destino) return { ok: false, error: 'El rol ' + gestor.rol + ' no gestiona vales de almacén/pañol.' };
  return { ok: true, gestor, destino };
}

// Helper: enriquecer vales con nombre y rol del solicitante.
function enriquecerVales(vales, usuarios) {
  return vales.map(v => {
    const sol = buscarUsuarioEnArray(usuarios, v.usuario_email);
    return {
      ...v,
      solicitante_nombre: sol ? sol.nombre : v.usuario_email,
      solicitante_rol   : sol ? sol.rol    : ''
    };
  });
}

// Helper: buscar fila en hoja por valor en columna (base 1).
function buscarFilaEnHoja(hoja, columna, valor) {
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][columna - 1]).toLowerCase() === String(valor).toLowerCase()) return i + 1;
  }
  return -1;
}

// Helper: validar admin.
function validarAdmin(email_admin) {
  const email = (email_admin || '').trim().toLowerCase();
  if (!email) return { ok: false, error: 'email_admin requerido.' };
  const admin = buscarUsuario(email);
  if (!admin)        return { ok: false, error: 'Usuario no registrado.' };
  if (!admin.activo) return { ok: false, error: 'Usuario inactivo.' };
  if (!ROLES_ADMIN.includes(admin.rol)) return { ok: false, error: 'Sin permiso: se requiere rol ADMIN.' };
  return { ok: true, admin };
}

// ── ROUTER ──────────────────────────────────────────────────────

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
      // ── Etapa 5 — Notificaciones ──
      case 'getNotificaciones'        : return acGetNotificaciones(e.parameter);
      case 'marcarLeidas'             : return acMarcarLeidas(e.parameter);
      // ── v10 — Contexto e Historial ──
      case 'getContexto'              : return acGetContexto(e.parameter);
      case 'getHistorialVale'         : return acGetHistorialVale(e.parameter);
      default                         : return respError('Acción desconocida: ' + accion);
    }
  } catch (err) {
    return respError('Error interno [' + accion + ']: ' + err.message);
  }
}

function doPost(e) {
  return doGet(e);
}

// ════════════════════════════════════════════════════════════════
// v10 — HISTORIAL_ESTADOS
// ════════════════════════════════════════════════════════════════

// ── Helper: obtener o crear hoja HISTORIAL_ESTADOS ───────────────
function getSheetHistorial() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  let hoja   = ss.getSheetByName(HOJA.HISTORIAL_ESTADOS);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA.HISTORIAL_ESTADOS);
    hoja.appendRow([
      'id_vale', 'estado_anterior', 'estado_nuevo',
      'usuario_email', 'usuario_nombre', 'fecha', 'nota'
    ]);
  }
  return hoja;
}

// ── Helper interno: registrar una transición de estado ───────────
// Llamado desde cada endpoint que muta el estado de un vale.
// Nunca lanza excepción — una falla en el historial no debe
// bloquear la operación principal.
//
// Parámetros:
//   id_vale        — ID del vale afectado
//   estado_anterior — estado previo (string vacío en creación)
//   estado_nuevo    — estado resultante
//   usuario_email   — email de quien realizó la acción
//   usuario_nombre  — nombre (desnormalizado para no releer USUARIOS)
//   nota            — texto adicional (nota de rechazo/cierre) o ''
function registrarTransicion(id_vale, estado_anterior, estado_nuevo, usuario_email, usuario_nombre, nota) {
  try {
    const hoja  = getSheetHistorial();
    const fecha = new Date().toISOString();
    hoja.appendRow([
      id_vale,
      estado_anterior || '',
      estado_nuevo,
      usuario_email,
      usuario_nombre  || usuario_email,
      fecha,
      nota            || ''
    ]);
    // No llamar SpreadsheetApp.flush() aquí — el llamador ya lo hace
    // después de actualizar VALES, así agrupamos el flush.
  } catch (e) {
    console.error('registrarTransicion error [' + id_vale + ']:', e.message);
  }
}

// ── ACCIÓN: getHistorialVale ─────────────────────────────────────
// GET ?accion=getHistorialVale&id_vale=xxx
// Devuelve todas las transiciones de un vale, ordenadas cronológicamente.
// Usado por el frontend para renderizar la timeline en el modal detalle.
function acGetHistorialVale(params) {
  const id_vale = (params.id_vale || '').trim();
  if (!id_vale) return respError('id_vale requerido.');

  try {
    const hoja  = getSheetHistorial();
    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return respOk([]);

    const historial = [];
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (String(fila[COL_HIST.id_vale - 1]) !== id_vale) continue;
      historial.push({
        estado_anterior: fila[COL_HIST.estado_anterior - 1],
        estado_nuevo   : fila[COL_HIST.estado_nuevo    - 1],
        usuario_email  : fila[COL_HIST.usuario_email   - 1],
        usuario_nombre : fila[COL_HIST.usuario_nombre  - 1],
        fecha          : fila[COL_HIST.fecha           - 1],
        nota           : fila[COL_HIST.nota            - 1]
      });
    }

    // Orden cronológico ASC (más antiguo primero — natural para timeline)
    historial.sort((a, b) => new Date(a.fecha) - new Date(b.fecha));

    return respOk(historial);
  } catch (e) {
    return respError('Error leyendo historial: ' + e.message);
  }
}

// ════════════════════════════════════════════════════════════════
// v10 — GETCONTEXTO (endpoint principal de carga inicial)
// ════════════════════════════════════════════════════════════════

// ── ACCIÓN: getContexto ──────────────────────────────────────────
// GET ?accion=getContexto&email=xxx
//
// Lee VALES, USUARIOS, NOTIFICACIONES una sola vez cada una.
// Según el rol del usuario arma y devuelve todo el contexto necesario
// para inicializar la app en una única call GAS.
//
// Respuesta según rol:
//   Todos los roles:
//     perfil, misVales, notificaciones
//   SUPERVISOR / JEFE_OBRA:
//     + porAprobar
//   JEFE_OBRA / ALMACENERO / PAÑOLERO / SUPERVISOR / ADMIN:
//     + todosVales
//   ALMACENERO / PAÑOLERO:
//     + porGestionar, enPreparacion, entregadosHoy
function acGetContexto(params) {
  const email = (params.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  // ── 1. Abrir spreadsheet UNA vez ──────────────────────────────
  const ss = getSpreadsheet();

  // ── 2. Leer hojas en paralelo (una sola lectura por hoja) ──────
  const hVales    = ss.getSheetByName(HOJA.VALES);
  const hUsuarios = ss.getSheetByName(HOJA.USUARIOS);
  const hObras    = ss.getSheetByName(HOJA.OBRAS);

  const rawVales    = hVales    ? hVales.getDataRange().getValues()    : [];
  const rawUsuarios = hUsuarios ? hUsuarios.getDataRange().getValues() : [];
  const rawObras    = hObras    ? hObras.getDataRange().getValues()    : [];

  const vales    = rawToObjects(rawVales,    COL_VALES);
  const usuarios = rawToObjects(rawUsuarios, COL_USUARIOS);
  const obras    = rawToObjects(rawObras,    COL_OBRAS);

  // ── 3. Validar usuario ─────────────────────────────────────────
  const usuario = buscarUsuarioEnArray(usuarios, email);
  if (!usuario)        return respError('Usuario no registrado.');
  if (!usuario.activo) return respError('Usuario inactivo. Contactá al administrador.');

  const obra = obras.find(o => o.codigo === usuario.obra_codigo && o.activa);
  if (!obra) return respError('Obra asignada no encontrada o inactiva.');

  // ── 4. Armar perfil ────────────────────────────────────────────
  const perfil = {
    email      : usuario.email,
    nombre     : usuario.nombre,
    rol        : usuario.rol,
    obra_codigo: usuario.obra_codigo,
    obra_desc  : obra.descripcion
  };

  // ── 5. Fechas de corte ─────────────────────────────────────────
  const hace30 = new Date();
  hace30.setDate(hace30.getDate() - 30);
  const hoy = new Date();
  const inicioHoy = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());

  // ── 6. Mis Vales (todos los roles excepto ADMIN) ───────────────
  // ADMIN no crea vales propios, pero calculamos igual (resultado vacío).
  const misVales = vales
    .filter(v => {
      if (v.usuario_email.toLowerCase() !== email) return false;
      if (v.eliminado === true || v.eliminado === 'TRUE') return false;
      return new Date(v.fecha_hora) >= hace30;
    })
    .sort((a, b) => new Date(b.fecha_hora) - new Date(a.fecha_hora));

  // ── 7. Por aprobar (SUPERVISOR, JEFE_OBRA) ─────────────────────
  let porAprobar = [];
  if (PUEDE_APROBAR[usuario.rol]) {
    const solicitantesAprobables = usuarios.filter(u =>
      u.activo && puedeAprobarA(usuario, u)
    );
    const emailsSol = solicitantesAprobables.map(u => u.email.toLowerCase());

    const pendientes = vales.filter(v =>
      v.estado === 'ENVIADO' &&
      v.eliminado !== true && v.eliminado !== 'TRUE' &&
      emailsSol.includes(v.usuario_email.toLowerCase())
    );
    pendientes.sort((a, b) => new Date(a.fecha_hora) - new Date(b.fecha_hora));
    porAprobar = enriquecerVales(pendientes, usuarios);
  }

  // ── 8. Gestión (ALMACENERO, PAÑOLERO) ─────────────────────────
  let porGestionar  = [];
  let enPreparacion = [];
  let entregadosHoy = 0;

  const destino = DESTINO_ROL[usuario.rol];
  if (destino) {
    const obra_codigo = usuario.obra_codigo;

    const cola = vales.filter(v =>
      v.estado === 'APROBADO' &&
      v.destino === destino &&
      v.obra_codigo === obra_codigo &&
      v.eliminado !== true && v.eliminado !== 'TRUE'
    );
    cola.sort((a, b) =>
      new Date(a.fecha_aprobacion || a.fecha_hora) - new Date(b.fecha_aprobacion || b.fecha_hora)
    );
    porGestionar = enriquecerVales(cola, usuarios);

    const enPrep = vales.filter(v =>
      v.estado === 'PENDIENTE' &&
      v.destino === destino &&
      v.obra_codigo === obra_codigo &&
      v.eliminado !== true && v.eliminado !== 'TRUE'
    );
    const enriquecidos = enriquecerVales(enPrep, usuarios);
    enriquecidos.sort((a, b) =>
      (a.solicitante_nombre || '').localeCompare(b.solicitante_nombre || '', 'es')
    );
    enPreparacion = enriquecidos;

    entregadosHoy = vales.filter(v => {
      if (v.eliminado === true || v.eliminado === 'TRUE') return false;
      if (!['ENTREGADO', 'ENTREGA_PARCIAL'].includes(v.estado)) return false;
      if (v.destino !== destino) return false;
      if (v.obra_codigo !== obra_codigo) return false;
      const fechaCierre = v.fecha_cierre ? new Date(v.fecha_cierre) : null;
      return fechaCierre && fechaCierre >= inicioHoy;
    }).length;
  }

  // ── 9. Todos los vales (JEFE_OBRA, ALMACENERO, PAÑOLERO, SUPERVISOR, ADMIN) ──
  const ROLES_VER_TODOS = ['JEFE_OBRA', 'ALMACENERO', 'PAÑOLERO', 'SUPERVISOR', 'ADMIN'];
  let todosVales = [];
  if (ROLES_VER_TODOS.includes(usuario.rol)) {
    const esSupervisor = usuario.rol === 'SUPERVISOR';
    const esAdmin      = usuario.rol === 'ADMIN';

    let resultado = vales.filter(v => {
      if (v.eliminado === true || v.eliminado === 'TRUE') return false;
      if (new Date(v.fecha_hora) < hace30) return false;
      if (esAdmin) return true;
      if (v.obra_codigo !== usuario.obra_codigo) return false;
      if (esSupervisor) {
        return (v.aprobado_por || '').toLowerCase() === email;
      }
      return true;
    });

    resultado = enriquecerVales(resultado, usuarios);
    resultado.sort((a, b) => new Date(b.fecha_hora) - new Date(a.fecha_hora));
    todosVales = resultado;
  }

  // ── 10. Notificaciones no leídas ──────────────────────────────
  let notificaciones = [];
  try {
    const hNotif  = ss.getSheetByName(HOJA.NOTIFICACIONES);
    if (hNotif) {
      const rawNotif = hNotif.getDataRange().getValues();
      if (rawNotif.length >= 2) {
        for (let i = 1; i < rawNotif.length; i++) {
          const fila  = rawNotif[i];
          const dest  = String(fila[COL_NOTIF.email_destino - 1]).toLowerCase();
          const leido = fila[COL_NOTIF.leido - 1];
          if (dest !== email) continue;
          if (leido === true || leido === 'TRUE') continue;
          notificaciones.push({
            fila         : i + 1,
            email_destino: fila[COL_NOTIF.email_destino - 1],
            mensaje      : fila[COL_NOTIF.mensaje       - 1],
            fecha        : fila[COL_NOTIF.fecha         - 1],
            leido        : false,
            id_vale      : fila[COL_NOTIF.id_vale       - 1]
          });
        }
      }
    }
  } catch (e) {
    console.error('getContexto - error leyendo notificaciones:', e.message);
    // No propagamos — notificaciones no deben bloquear la carga
  }

  // ── 11. Devolver contexto completo ────────────────────────────
  return respOk({
    perfil,
    misVales,
    porAprobar,
    porGestionar,
    enPreparacion,
    entregadosHoy,
    todosVales,
    notificaciones
  });
}

// ════════════════════════════════════════════════════════════════
// ETAPA 1 — ENDPOINTS
// ════════════════════════════════════════════════════════════════

// ── ACCIÓN: login ────────────────────────────────────────────────
// GET ?accion=login&email=xxx
// Sigue existiendo para compatibilidad. En v10 el login llama primero
// a Google Identity Services y luego a getContexto (no a login).
// Se mantiene por si el frontend necesita validar sin cargar contexto.
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
// Mantenido por compatibilidad. El frontend v13 usa getContexto.
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
    return new Date(v.fecha_hora) >= hace30;
  });

  misVales.sort((a, b) => new Date(b.fecha_hora) - new Date(a.fecha_hora));
  return respOk(misVales);
}

// ── ACCIÓN: guardarVale ──────────────────────────────────────────
// GET ?accion=guardarVale&id_vale=xxx&...&estado_anterior=xxx
//
// v10: recibe estado_anterior para registrar correctamente la transición
// cuando un BORRADOR se envía (BORRADOR → ENVIADO).
// Si estado_anterior está vacío, se infiere desde la hoja.
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
    fecha_hora     : params.fecha_hora,
    estado_anterior: params.estado_anterior || ''
  };

  const { id_vale, usuario_email, usuario_nombre, obra_codigo,
          destino, titulo, contenido_html, estado, fecha_hora,
          estado_anterior } = body;

  if (!id_vale)        return respError('id_vale requerido.');
  if (!usuario_email)  return respError('usuario_email requerido.');
  if (!destino)        return respError('destino requerido (ALMACEN o PAÑOL).');
  if (!contenido_html) return respError('contenido_html requerido.');

  // ── Auto-aprobación JEFE_OBRA ─────────────────────────────────
  let estadoFinal      = estado || 'BORRADOR';
  let aprobado_por     = '';
  let fecha_aprobacion = '';

  if (estadoFinal === 'ENVIADO') {
    const solicitante = buscarUsuario(usuario_email);
    if (solicitante && solicitante.rol === 'JEFE_OBRA') {
      estadoFinal      = 'APROBADO';
      aprobado_por     = usuario_email;
      fecha_aprobacion = new Date().toISOString();
    }
  }

  const hVales = getSheet(HOJA.VALES);
  const datos  = hVales.getDataRange().getValues();
  let filaExistente = buscarFilaVale(datos, id_vale);

  const ahora    = new Date();
  const fechaHora = fecha_hora || ahora.toISOString();

  // Estado anterior real: lo que había antes de este save
  let estadoPrevio = estado_anterior || '';

  if (filaExistente > 0) {
    // UPDATE
    // Si no vino estado_anterior, leerlo de la hoja (fila existente)
    if (!estadoPrevio) {
      estadoPrevio = datos[filaExistente - 1][COL_VALES.estado - 1] || '';
    }
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
    estadoPrevio = ''; // nueva creación, no hay estado anterior
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
      aprobado_por,
      fecha_aprobacion,
      '',    // gestionado_por
      '',    // fecha_cierre
      '',    // nota_cierre
      false  // eliminado
    ];
    hVales.appendRow(nuevaFila);
  }

  SpreadsheetApp.flush();

  // ── Registrar transición(es) en HISTORIAL_ESTADOS ─────────────
  // Solo registramos si el estado cambió efectivamente.
  if (estadoFinal !== estadoPrevio) {
    registrarTransicion(id_vale, estadoPrevio, estadoFinal, usuario_email, usuario_nombre, '');
  }
  // Si hubo auto-aprobación, registrar también ENVIADO → APROBADO
  // (ya registramos '' → ENVIADO arriba; ahora la segunda transición)
  if (aprobado_por && estadoFinal === 'APROBADO' && estado === 'ENVIADO') {
    registrarTransicion(id_vale, 'ENVIADO', 'APROBADO', usuario_email, usuario_nombre, 'Auto-aprobado (JEFE_OBRA)');
  }

  SpreadsheetApp.flush();

  // ── Notificaciones ────────────────────────────────────────────
  if (estadoFinal === 'ENVIADO' || estadoFinal === 'APROBADO') {
    const valeParaNotif = leerVale(id_vale);
    if (valeParaNotif) {
      dispararNotificaciones(estadoFinal, valeParaNotif, null);
    }
  }

  return respOk({
    id_vale,
    estado       : estadoFinal,
    accion       : filaExistente > 0 ? 'updated' : 'inserted',
    auto_aprobado: estadoFinal === 'APROBADO' && aprobado_por !== ''
  });
}

// ── ACCIÓN: eliminarVale ─────────────────────────────────────────
// GET ?accion=eliminarVale&id_vale=xxx&usuario_email=xxx
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
    const ahora   = new Date().toISOString();
    hVales.getRange(numFila, COL_VALES.eliminado,    1, 1).setValue(true);
    hVales.getRange(numFila, COL_VALES.estado,       1, 1).setValue('ELIMINADO');
    hVales.getRange(numFila, COL_VALES.fecha_cierre, 1, 1).setValue(ahora);

    // Leer nombre del usuario para el historial
    const usuarioNombre = fila[COL_VALES.usuario_nombre - 1] || usuario_email;
    SpreadsheetApp.flush();

    registrarTransicion(id_vale, 'BORRADOR', 'ELIMINADO', usuario_email, usuarioNombre, '');
    SpreadsheetApp.flush();

    return respOk({ id_vale, eliminado: true });
  }

  return respError('Vale no encontrado: ' + id_vale);
}

// ════════════════════════════════════════════════════════════════
// ETAPA 2 — ENDPOINTS DE APROBACIÓN
// ════════════════════════════════════════════════════════════════

// ── ACCIÓN: getPendientesAprobacion ─────────────────────────────
// Mantenido por compatibilidad. El frontend v13 usa getContexto.
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

  pendientes.sort((a, b) => new Date(a.fecha_hora) - new Date(b.fecha_hora));
  return respOk(enriquecerVales(pendientes, usuarios));
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

  const filaData     = datos[fila - 1];
  const estadoActual = filaData[COL_VALES.estado - 1];
  if (estadoActual !== 'ENVIADO') {
    return respError('El vale no está en estado ENVIADO. Estado actual: ' + estadoActual);
  }

  const emailSolicitante = filaData[COL_VALES.usuario_email - 1].toLowerCase();
  const solicitante      = buscarUsuario(emailSolicitante);
  if (!solicitante || !puedeAprobarA(aprobador, solicitante)) {
    return respError('No tenés permiso para aprobar este vale.');
  }

  const obraVale = filaData[COL_VALES.obra_codigo - 1];
  if (obraVale !== aprobador.obra_codigo) {
    return respError('No podés aprobar vales de otra obra.');
  }

  const ahora = new Date().toISOString();
  hVales.getRange(fila, COL_VALES.estado,           1, 1).setValue('APROBADO');
  hVales.getRange(fila, COL_VALES.aprobado_por,     1, 1).setValue(aprobado_por);
  hVales.getRange(fila, COL_VALES.fecha_aprobacion, 1, 1).setValue(ahora);
  SpreadsheetApp.flush();

  registrarTransicion(id_vale, 'ENVIADO', 'APROBADO', aprobado_por, aprobador.nombre, '');
  SpreadsheetApp.flush();

  const valeAprobado = leerVale(id_vale);
  if (valeAprobado) dispararNotificaciones('APROBADO', valeAprobado, null);

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

  const filaData     = datos[fila - 1];
  const estadoActual = filaData[COL_VALES.estado - 1];
  if (estadoActual !== 'ENVIADO') {
    return respError('El vale no está en estado ENVIADO. Estado actual: ' + estadoActual);
  }

  const emailSolicitante = filaData[COL_VALES.usuario_email - 1].toLowerCase();
  const solicitante      = buscarUsuario(emailSolicitante);
  if (!solicitante || !puedeAprobarA(aprobador, solicitante)) {
    return respError('No tenés permiso para rechazar este vale.');
  }

  const obraVale = filaData[COL_VALES.obra_codigo - 1];
  if (obraVale !== aprobador.obra_codigo) {
    return respError('No podés rechazar vales de otra obra.');
  }

  const ahora = new Date().toISOString();
  hVales.getRange(fila, COL_VALES.estado,           1, 1).setValue('RECHAZADO');
  hVales.getRange(fila, COL_VALES.aprobado_por,     1, 1).setValue(rechazado_por);
  hVales.getRange(fila, COL_VALES.fecha_aprobacion, 1, 1).setValue(ahora);
  hVales.getRange(fila, COL_VALES.nota_cierre,      1, 1).setValue(nota);
  SpreadsheetApp.flush();

  registrarTransicion(id_vale, 'ENVIADO', 'RECHAZADO', rechazado_por, aprobador.nombre, nota);
  SpreadsheetApp.flush();

  const valeRechazado = leerVale(id_vale);
  if (valeRechazado) dispararNotificaciones('RECHAZADO', valeRechazado, nota);

  return respOk({ id_vale, estado: 'RECHAZADO', rechazado_por, fecha: ahora, nota_cierre: nota });
}

// ── ACCIÓN: getValesAprobados ────────────────────────────────────
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

// ── ACCIÓN: getValesPorGestionar ─────────────────────────────────
// Mantenido por compatibilidad. El frontend v13 usa getContexto.
function acGetValesPorGestionar(params) {
  const email = (params.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  const v = validarGestor(email);
  if (!v.ok) return respError(v.error);
  const { gestor, destino } = v;

  const hVales    = getSheet(HOJA.VALES);
  const hUsuarios = getSheet(HOJA.USUARIOS);
  const vales     = sheetToObjects(hVales, COL_VALES);
  const usuarios  = sheetToObjects(hUsuarios, COL_USUARIOS);

  const cola = vales.filter(vale =>
    vale.estado === 'APROBADO' &&
    vale.destino === destino &&
    vale.obra_codigo === gestor.obra_codigo &&
    vale.eliminado !== true && vale.eliminado !== 'TRUE'
  );

  cola.sort((a, b) =>
    new Date(a.fecha_aprobacion || a.fecha_hora) - new Date(b.fecha_aprobacion || b.fecha_hora)
  );
  return respOk(enriquecerVales(cola, usuarios));
}

// ── ACCIÓN: getValesPendientes ────────────────────────────────────
// Mantenido por compatibilidad.
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
  if (estado !== 'APROBADO') {
    return respError('El vale debe estar en estado APROBADO. Estado actual: ' + estado);
  }

  const valeDestino = filaData[COL_VALES.destino - 1];
  if (valeDestino !== destino) {
    return respError('Este vale es para ' + valeDestino + ', no para ' + destino + '.');
  }

  const obraVale = filaData[COL_VALES.obra_codigo - 1];
  if (obraVale !== gestor.obra_codigo) return respError('El vale pertenece a otra obra.');

  const ahora = new Date().toISOString();
  hVales.getRange(fila, COL_VALES.estado,         1, 1).setValue('PENDIENTE');
  hVales.getRange(fila, COL_VALES.gestionado_por, 1, 1).setValue(gestionado_por);
  hVales.getRange(fila, COL_VALES.fecha_cierre,   1, 1).setValue(ahora);
  SpreadsheetApp.flush();

  registrarTransicion(id_vale, 'APROBADO', 'PENDIENTE', gestionado_por, gestor.nombre, '');
  SpreadsheetApp.flush();

  return respOk({ id_vale, estado: 'PENDIENTE', gestionado_por, fecha: ahora });
}

// ── Helper compartido: cerrar vale ───────────────────────────────
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

  const filaData     = datos[fila - 1];
  const estadoActual = filaData[COL_VALES.estado - 1];

  const estadosValidos = estadoDestino === 'CANCELADO'
    ? ['APROBADO', 'PENDIENTE']
    : ['PENDIENTE'];

  if (!estadosValidos.includes(estadoActual)) {
    return respError('Estado inválido para esta acción. Estado actual: ' + estadoActual);
  }

  const valeDestino = filaData[COL_VALES.destino - 1];
  if (valeDestino !== destino) {
    return respError('Este vale es para ' + valeDestino + ', no para ' + destino + '.');
  }

  const obraVale = filaData[COL_VALES.obra_codigo - 1];
  if (obraVale !== gestor.obra_codigo) return respError('El vale pertenece a otra obra.');

  const ahora = new Date().toISOString();
  hVales.getRange(fila, COL_VALES.estado,         1, 1).setValue(estadoDestino);
  hVales.getRange(fila, COL_VALES.gestionado_por, 1, 1).setValue(gestionado_por);
  hVales.getRange(fila, COL_VALES.fecha_cierre,   1, 1).setValue(ahora);
  if (nota) hVales.getRange(fila, COL_VALES.nota_cierre, 1, 1).setValue(nota);
  SpreadsheetApp.flush();

  registrarTransicion(id_vale, estadoActual, estadoDestino, gestionado_por, gestor.nombre, nota);
  SpreadsheetApp.flush();

  const valeCerrado = leerVale(id_vale);
  if (valeCerrado) dispararNotificaciones(estadoDestino, valeCerrado, nota);

  return respOk({ id_vale, estado: estadoDestino, gestionado_por, fecha: ahora, nota_cierre: nota });
}

function acEntregarVale(params)       { return acCerrarVale(params, 'ENTREGADO',       false); }
function acEntregaParcialVale(params) { return acCerrarVale(params, 'ENTREGA_PARCIAL', true);  }
function acCancelarVale(params)       { return acCerrarVale(params, 'CANCELADO',       true);  }

// ── ACCIÓN: getTodosVales ─────────────────────────────────────────
// Mantenido por compatibilidad.
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
    if (new Date(v.fecha_hora) < hace30) return false;
    if (esAdmin) return true;
    if (v.obra_codigo !== usuario.obra_codigo) return false;
    if (esSupervisor) return (v.aprobado_por || '').toLowerCase() === email;
    return true;
  });

  const enriquecidos = enriquecerVales(resultado, usuarios);
  const filtrado = filtroSolicitante
    ? enriquecidos.filter(v =>
        (v.solicitante_nombre || '').toLowerCase().includes(filtroSolicitante) ||
        (v.usuario_email      || '').toLowerCase().includes(filtroSolicitante)
      )
    : enriquecidos;

  filtrado.sort((a, b) => new Date(b.fecha_hora) - new Date(a.fecha_hora));
  return respOk(filtrado);
}

// ── ACCIÓN: getEntregadosHoy ──────────────────────────────────────
// Mantenido por compatibilidad.
function acGetEntregadosHoy(params) {
  const email = (params.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  const v = validarGestor(email);
  if (!v.ok) return respError(v.error);
  const { gestor, destino } = v;

  const hVales = getSheet(HOJA.VALES);
  const vales  = sheetToObjects(hVales, COL_VALES);

  const hoy       = new Date();
  const inicioHoy = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());

  const entregados = vales.filter(vale => {
    if (vale.eliminado === true || vale.eliminado === 'TRUE') return false;
    if (!['ENTREGADO', 'ENTREGA_PARCIAL'].includes(vale.estado)) return false;
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

function acGetUsuarios(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const hUsuarios = getSheet(HOJA.USUARIOS);
  const usuarios  = sheetToObjects(hUsuarios, COL_USUARIOS);

  usuarios.sort((a, b) => {
    const obra = String(a.obra_codigo).localeCompare(String(b.obra_codigo), 'es');
    if (obra !== 0) return obra;
    return String(a.nombre).localeCompare(String(b.nombre), 'es');
  });

  return respOk(usuarios);
}

function acCrearUsuario(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const email          = (params.email          || '').trim().toLowerCase();
  const nombre         = (params.nombre         || '').trim();
  const rol            = (params.rol            || '').trim().toUpperCase();
  const obra_codigo    = (params.obra_codigo    || '').trim().toUpperCase();
  const activo         = params.activo !== 'FALSE';
  const superior_email = (params.superior_email || '').trim().toLowerCase();

  if (!email)       return respError('email requerido.');
  if (!nombre)      return respError('nombre requerido.');
  if (!rol)         return respError('rol requerido.');
  if (!obra_codigo) return respError('obra_codigo requerido.');

  const ROLES_VALIDOS = ['CAPATAZ', 'SUPERVISOR', 'JEFE_OBRA', 'ALMACENERO', 'PAÑOLERO', 'ADMIN'];
  if (!ROLES_VALIDOS.includes(rol)) return respError('Rol inválido: ' + rol);

  const hObras = getSheet(HOJA.OBRAS);
  const obras  = sheetToObjects(hObras, COL_OBRAS);
  const obra   = obras.find(o => String(o.codigo).toUpperCase() === obra_codigo);
  if (!obra)        return respError('Obra no encontrada: ' + obra_codigo);
  if (!obra.activa) return respError('La obra ' + obra_codigo + ' está inactiva.');

  const hUsuarios      = getSheet(HOJA.USUARIOS);
  const filaExistente  = buscarFilaEnHoja(hUsuarios, COL_USUARIOS.email, email);
  if (filaExistente > 0) return respError('Ya existe un usuario con ese email: ' + email);

  if (superior_email) {
    const superior = buscarUsuario(superior_email);
    if (!superior) return respError('superior_email no encontrado: ' + superior_email);
  }

  hUsuarios.appendRow([email, nombre, rol, obra_codigo, activo, superior_email]);
  SpreadsheetApp.flush();
  return respOk({ email, nombre, rol, obra_codigo, activo, superior_email, accion: 'created' });
}

function acEditarUsuario(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const email_usuario = (params.email_usuario || '').trim().toLowerCase();
  if (!email_usuario) return respError('email_usuario requerido.');

  const hUsuarios = getSheet(HOJA.USUARIOS);
  const fila      = buscarFilaEnHoja(hUsuarios, COL_USUARIOS.email, email_usuario);
  if (fila < 0) return respError('Usuario no encontrado: ' + email_usuario);

  const datos       = hUsuarios.getDataRange().getValues();
  const filaActual  = datos[fila - 1];

  const nombre         = (params.nombre      || '').trim()            || filaActual[COL_USUARIOS.nombre         - 1];
  const rol            = (params.rol         || '').trim().toUpperCase() || filaActual[COL_USUARIOS.rol            - 1];
  const obra_codigo    = (params.obra_codigo || '').trim().toUpperCase() || filaActual[COL_USUARIOS.obra_codigo    - 1];
  const superior_email = params.superior_email !== undefined
    ? (params.superior_email || '').trim().toLowerCase()
    : filaActual[COL_USUARIOS.superior_email - 1];

  const ROLES_VALIDOS = ['CAPATAZ', 'SUPERVISOR', 'JEFE_OBRA', 'ALMACENERO', 'PAÑOLERO', 'ADMIN'];
  if (!ROLES_VALIDOS.includes(rol)) return respError('Rol inválido: ' + rol);

  const hObras = getSheet(HOJA.OBRAS);
  const obras  = sheetToObjects(hObras, COL_OBRAS);
  const obra   = obras.find(o => String(o.codigo).toUpperCase() === obra_codigo);
  if (!obra)        return respError('Obra no encontrada: ' + obra_codigo);
  if (!obra.activa) return respError('La obra ' + obra_codigo + ' está inactiva.');

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

function acToggleUsuario(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const email_admin   = (params.email_admin   || '').trim().toLowerCase();
  const email_usuario = (params.email_usuario || '').trim().toLowerCase();
  if (!email_usuario) return respError('email_usuario requerido.');
  if (email_usuario === email_admin) return respError('No podés desactivarte a vos mismo.');

  const hUsuarios = getSheet(HOJA.USUARIOS);
  const fila      = buscarFilaEnHoja(hUsuarios, COL_USUARIOS.email, email_usuario);
  if (fila < 0) return respError('Usuario no encontrado: ' + email_usuario);

  const datos      = hUsuarios.getDataRange().getValues();
  const activoActual = datos[fila - 1][COL_USUARIOS.activo - 1];
  const nuevoActivo  = !(activoActual === true || activoActual === 'TRUE');

  hUsuarios.getRange(fila, COL_USUARIOS.activo, 1, 1).setValue(nuevoActivo);
  SpreadsheetApp.flush();
  return respOk({ email: email_usuario, activo: nuevoActivo });
}

function acGetObras(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const hObras = getSheet(HOJA.OBRAS);
  const obras  = sheetToObjects(hObras, COL_OBRAS);
  obras.sort((a, b) => String(a.codigo).localeCompare(String(b.codigo), 'es'));
  return respOk(obras);
}

function acCrearObra(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const codigo        = (params.codigo        || '').trim().toUpperCase();
  const descripcion   = (params.descripcion   || '').trim();
  const email_panol   = (params.email_panol   || params['email_pa\u00f1ol'] || '').trim().toLowerCase();
  const email_almacen = (params.email_almacen || '').trim().toLowerCase();
  const activa        = params.activa !== 'FALSE';

  if (!codigo)      return respError('codigo requerido.');
  if (!descripcion) return respError('descripcion requerida.');

  const hObras = getSheet(HOJA.OBRAS);
  const filaExistente = buscarFilaEnHoja(hObras, COL_OBRAS.codigo, codigo);
  if (filaExistente > 0) return respError('Ya existe una obra con ese código: ' + codigo);

  hObras.appendRow([codigo, descripcion, email_panol, email_almacen, activa]);
  SpreadsheetApp.flush();
  return respOk({ codigo, descripcion, email_panol, email_almacen, activa, accion: 'created' });
}

function acEditarObra(params) {
  const v = validarAdmin(params.email_admin);
  if (!v.ok) return respError(v.error);

  const codigo = (params.codigo || '').trim().toUpperCase();
  if (!codigo) return respError('codigo requerido.');

  const hObras = getSheet(HOJA.OBRAS);
  const fila   = buscarFilaEnHoja(hObras, COL_OBRAS.codigo, codigo);
  if (fila < 0) return respError('Obra no encontrada: ' + codigo);

  const datos       = hObras.getDataRange().getValues();
  const filaActual  = datos[fila - 1];

  const descripcion   = (params.descripcion   || '').trim() || filaActual[COL_OBRAS.descripcion   - 1];
  const email_panol   = params.email_panol   !== undefined
    ? (params.email_panol   || '').trim().toLowerCase()
    : filaActual[COL_OBRAS.email_panol - 1];
  const email_almacen = params.email_almacen !== undefined
    ? (params.email_almacen || '').trim().toLowerCase()
    : filaActual[COL_OBRAS.email_almacen - 1];

  if (!descripcion) return respError('descripcion requerida.');

  hObras.getRange(fila, COL_OBRAS.descripcion,   1, 1).setValue(descripcion);
  hObras.getRange(fila, COL_OBRAS.email_panol,   1, 1).setValue(email_panol);
  hObras.getRange(fila, COL_OBRAS.email_almacen, 1, 1).setValue(email_almacen);
  SpreadsheetApp.flush();
  return respOk({ codigo, descripcion, email_panol, email_almacen, accion: 'updated' });
}

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

// ════════════════════════════════════════════════════════════════
// ETAPA 5 — NOTIFICACIONES
// ════════════════════════════════════════════════════════════════

function getSheetNotificaciones() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  let hoja   = ss.getSheetByName(HOJA.NOTIFICACIONES);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA.NOTIFICACIONES);
    hoja.appendRow(['email_destino', 'mensaje', 'fecha', 'leido', 'id_vale']);
  }
  return hoja;
}

function obtenerEmailAprobador(solicitante, usuarios) {
  if (solicitante.superior_email && solicitante.superior_email.trim() !== '') {
    return solicitante.superior_email.trim().toLowerCase();
  }
  const rolAprobador = Object.keys(PUEDE_APROBAR).find(rol =>
    (PUEDE_APROBAR[rol] || []).includes(solicitante.rol)
  );
  if (!rolAprobador) return null;
  const aprobador = usuarios.find(u =>
    u.rol === rolAprobador &&
    u.obra_codigo === solicitante.obra_codigo &&
    u.activo
  );
  return aprobador ? aprobador.email.toLowerCase() : null;
}

function obtenerEmailsGestores(destino, obra_codigo, usuarios) {
  const rolGestor = destino === 'ALMACEN' ? 'ALMACENERO' : 'PAÑOLERO';
  return usuarios
    .filter(u => u.rol === rolGestor && u.obra_codigo === obra_codigo && u.activo)
    .map(u => u.email.toLowerCase());
}

function crearNotificacion(email_destino, mensaje, id_vale, datosEmail) {
  if (!email_destino) return;
  try {
    const hoja  = getSheetNotificaciones();
    const fecha = new Date().toISOString();
    hoja.appendRow([email_destino, mensaje, fecha, false, id_vale || '']);
    SpreadsheetApp.flush();
  } catch (e) {
    console.error('Error escribiendo notificación:', e.message);
  }
  if (datosEmail) {
    try {
      MailApp.sendEmail({
        to      : email_destino,
        subject : datosEmail.asunto,
        htmlBody: datosEmail.cuerpo
      });
    } catch (e) {
      console.error('Error enviando email a ' + email_destino + ':', e.message);
    }
  }
}

function construirCuerpoEmail(evento, vale, solicitante_nombre, nota) {
  const urlApp  = 'https://datamegashare.github.io/vale/';
  const estadoColor = {
    'ENVIADO'        : '#2874a6',
    'APROBADO'       : '#1e8449',
    'RECHAZADO'      : '#c0392b',
    'PENDIENTE'      : '#d68910',
    'ENTREGADO'      : '#1e8449',
    'ENTREGA_PARCIAL': '#d68910',
    'CANCELADO'      : '#717d7e'
  };
  const color = estadoColor[evento] || '#1A2B45';

  return `<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:32px 0;">
    <tr><td align="center">
      <table width="560" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.10);">
        <tr><td style="background:#1A2B45;padding:20px 32px;">
          <span style="color:#fff;font-size:20px;font-weight:700;">&#128196; Vale Digital</span>
        </td></tr>
        <tr><td style="padding:28px 32px 0;">
          <span style="display:inline-block;background:${color};color:#fff;font-size:13px;font-weight:700;padding:4px 14px;border-radius:20px;letter-spacing:.5px;text-transform:uppercase;">${evento.replace('_',' ')}</span>
        </td></tr>
        <tr><td style="padding:16px 32px 0;">
          <h2 style="margin:0;font-size:20px;color:#1A2B45;">${vale.titulo || 'Sin título'}</h2>
        </td></tr>
        <tr><td style="padding:20px 32px;">
          <table cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td style="padding:6px 0;color:#666;font-size:13px;width:140px;">Solicitante</td>
              <td style="padding:6px 0;color:#1A2B45;font-size:13px;font-weight:600;">${solicitante_nombre || vale.usuario_email}</td>
            </tr>
            <tr>
              <td style="padding:6px 0;color:#666;font-size:13px;">Obra</td>
              <td style="padding:6px 0;color:#1A2B45;font-size:13px;">${vale.obra_codigo}</td>
            </tr>
            <tr>
              <td style="padding:6px 0;color:#666;font-size:13px;">Destino</td>
              <td style="padding:6px 0;color:#1A2B45;font-size:13px;">${vale.destino}</td>
            </tr>
            <tr>
              <td style="padding:6px 0;color:#666;font-size:13px;">Fecha</td>
              <td style="padding:6px 0;color:#1A2B45;font-size:13px;">${new Date(vale.fecha_hora).toLocaleString('es-AR',{dateStyle:'short',timeStyle:'short'})}</td>
            </tr>
            ${nota ? `<tr>
              <td style="padding:6px 0;color:#666;font-size:13px;">Nota</td>
              <td style="padding:6px 0;color:#1A2B45;font-size:13px;">${nota}</td>
            </tr>` : ''}
          </table>
        </td></tr>
        <tr><td style="padding:0 32px 20px;">
          <div style="background:#f4f6f8;border-radius:6px;padding:14px 16px;font-size:13px;color:#333;line-height:1.6;">
            ${vale.contenido_html || ''}
          </div>
        </td></tr>
        <tr><td style="padding:0 32px 28px;" align="center">
          <a href="${urlApp}" style="display:inline-block;background:#2e6da4;color:#fff;font-size:14px;font-weight:700;padding:12px 32px;border-radius:6px;text-decoration:none;">Abrir Vale Digital</a>
        </td></tr>
        <tr><td style="background:#f4f6f8;padding:16px 32px;border-top:1px solid #e0e0e0;">
          <p style="margin:0;font-size:11px;color:#999;text-align:center;">Este email fue generado automáticamente por Vale Digital. No respondas este mensaje.</p>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>`;
}

function dispararNotificaciones(evento, vale, nota) {
  try {
    const hUsuarios = getSheet(HOJA.USUARIOS);
    const usuarios  = sheetToObjects(hUsuarios, COL_USUARIOS);

    const solicitante = usuarios.find(u =>
      u.email.toLowerCase() === vale.usuario_email.toLowerCase()
    );
    const solNombre = solicitante ? solicitante.nombre : vale.usuario_email;
    const asuntoBase = `[Vale Digital] ${evento.replace('_',' ')} — ${vale.titulo || 'Sin título'}`;

    switch (evento) {
      case 'ENVIADO': {
        const emailAprobador = solicitante
          ? obtenerEmailAprobador(solicitante, usuarios)
          : null;
        if (emailAprobador) {
          const msg = `Vale "${vale.titulo}" enviado por ${solNombre} — pendiente de aprobación.`;
          crearNotificacion(emailAprobador, msg, vale.id_vale, {
            asunto: asuntoBase,
            cuerpo: construirCuerpoEmail('ENVIADO', vale, solNombre, null)
          });
        }
        break;
      }
      case 'APROBADO': {
        crearNotificacion(vale.usuario_email,
          `Tu vale "${vale.titulo}" fue aprobado.`,
          vale.id_vale,
          { asunto: asuntoBase, cuerpo: construirCuerpoEmail('APROBADO', vale, solNombre, null) }
        );
        obtenerEmailsGestores(vale.destino, vale.obra_codigo, usuarios).forEach(eg => {
          crearNotificacion(eg,
            `Nuevo vale aprobado para gestionar: "${vale.titulo}" (${vale.destino}).`,
            vale.id_vale,
            { asunto: asuntoBase, cuerpo: construirCuerpoEmail('APROBADO', vale, solNombre, null) }
          );
        });
        break;
      }
      case 'RECHAZADO': {
        crearNotificacion(vale.usuario_email,
          `Tu vale "${vale.titulo}" fue rechazado.${nota ? ' Motivo: ' + nota : ''}`,
          vale.id_vale,
          { asunto: asuntoBase, cuerpo: construirCuerpoEmail('RECHAZADO', vale, solNombre, nota) }
        );
        break;
      }
      case 'ENTREGADO':
      case 'ENTREGA_PARCIAL':
      case 'CANCELADO': {
        crearNotificacion(vale.usuario_email,
          `Tu vale "${vale.titulo}" fue ${evento.replace('_',' ').toLowerCase()}.${nota ? ' Nota: ' + nota : ''}`,
          vale.id_vale,
          { asunto: asuntoBase, cuerpo: construirCuerpoEmail(evento, vale, solNombre, nota) }
        );
        break;
      }
    }
  } catch (e) {
    console.error('Error en dispararNotificaciones:', e.message);
  }
}

function leerVale(id_vale) {
  const hVales = getSheet(HOJA.VALES);
  const datos  = hVales.getDataRange().getValues();
  const fila   = buscarFilaVale(datos, id_vale);
  if (fila < 0) return null;
  const f   = datos[fila - 1];
  const obj = {};
  Object.keys(COL_VALES).forEach(key => { obj[key] = f[COL_VALES[key] - 1]; });
  return obj;
}

// ── ACCIÓN: getNotificaciones ─────────────────────────────────────
// Mantenido por compatibilidad. El frontend v13 usa getContexto.
function acGetNotificaciones(params) {
  const email = (params.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  try {
    const hoja   = getSheetNotificaciones();
    const datos  = hoja.getDataRange().getValues();
    if (datos.length < 2) return respOk([]);

    const noLeidas = [];
    for (let i = 1; i < datos.length; i++) {
      const fila  = datos[i];
      const dest  = String(fila[COL_NOTIF.email_destino - 1]).toLowerCase();
      const leido = fila[COL_NOTIF.leido - 1];
      if (dest !== email) continue;
      if (leido === true || leido === 'TRUE') continue;
      noLeidas.push({
        fila         : i + 1,
        email_destino: fila[COL_NOTIF.email_destino - 1],
        mensaje      : fila[COL_NOTIF.mensaje       - 1],
        fecha        : fila[COL_NOTIF.fecha         - 1],
        leido        : false,
        id_vale      : fila[COL_NOTIF.id_vale       - 1]
      });
    }
    return respOk(noLeidas);
  } catch (e) {
    return respError('Error leyendo notificaciones: ' + e.message);
  }
}

// ── ACCIÓN: marcarLeidas ──────────────────────────────────────────
function acMarcarLeidas(params) {
  const email = (params.email || '').trim().toLowerCase();
  if (!email) return respError('Email requerido.');

  try {
    const hoja  = getSheetNotificaciones();
    const datos = hoja.getDataRange().getValues();
    let count   = 0;

    for (let i = 1; i < datos.length; i++) {
      const dest  = String(datos[i][COL_NOTIF.email_destino - 1]).toLowerCase();
      const leido = datos[i][COL_NOTIF.leido - 1];
      if (dest !== email) continue;
      if (leido === true || leido === 'TRUE') continue;
      hoja.getRange(i + 1, COL_NOTIF.leido, 1, 1).setValue(true);
      count++;
    }

    if (count > 0) SpreadsheetApp.flush();
    return respOk({ marcadas: count });
  } catch (e) {
    return respError('Error marcando leídas: ' + e.message);
  }
}
