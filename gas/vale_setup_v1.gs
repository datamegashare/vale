// ═══════════════════════════════════════════════════════════════
// Vale Digital — Setup v1.0
// Corre UNA sola vez para inicializar el Sheet.
// Editá los valores de la sección CONFIGURACIÓN antes de ejecutar.
// ═══════════════════════════════════════════════════════════════

function setupValeDigital() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── CONFIGURACIÓN — editá estos valores ─────────────────────
  const ADMIN_EMAIL   = 'admin@tuempresa.com';
  const ADMIN_NOMBRE  = 'Administrador';
  const OBRA_CODIGO   = 'OB-001';
  const OBRA_DESC     = 'Obra Electromecánica 001';
  const EMAIL_PAÑOL   = 'pañol@tuempresa.com';
  const EMAIL_ALMACEN = 'almacen@tuempresa.com';
  // ────────────────────────────────────────────────────────────

  // ── Definición de hojas ─────────────────────────────────────
  const SHEETS = {
    VALES: [
      'id_vale', 'fecha_hora', 'usuario_email', 'usuario_nombre',
      'obra_codigo', 'destino', 'titulo', 'contenido_html',
      'estado', 'aprobado_por', 'fecha_aprobacion', 'gestionado_por',
      'fecha_cierre', 'nota_cierre', 'eliminado'
    ],
    USUARIOS: [
      'email', 'nombre', 'rol', 'obra_codigo', 'activo'
    ],
    OBRAS: [
      'codigo', 'descripcion', 'email_pañol', 'email_almacen', 'activa'
    ]
  };

  // ── Datos de prueba ─────────────────────────────────────────
  const SEED = {
    USUARIOS: [
      [ADMIN_EMAIL, ADMIN_NOMBRE, 'ADMIN', OBRA_CODIGO, true],
      ['capataz@tuempresa.com',   'Capataz Prueba',    'CAPATAZ',    OBRA_CODIGO, true],
      ['supervisor@tuempresa.com','Supervisor Prueba', 'SUPERVISOR', OBRA_CODIGO, true],
      ['jefe@tuempresa.com',      'Jefe de Obra',      'JEFE_OBRA',  OBRA_CODIGO, true],
      ['almacen@tuempresa.com',   'Almacenero Prueba', 'ALMACENERO', OBRA_CODIGO, true],
      ['pañol@tuempresa.com',     'Pañolero Prueba',   'PAÑOLERO',   OBRA_CODIGO, true],
    ],
    OBRAS: [
      [OBRA_CODIGO, OBRA_DESC, EMAIL_PAÑOL, EMAIL_ALMACEN, true]
    ]
  };

  // ── Crear / verificar hojas ──────────────────────────────────
  Object.entries(SHEETS).forEach(([nombre, columnas]) => {
    let hoja = ss.getSheetByName(nombre);

    if (!hoja) {
      hoja = ss.insertSheet(nombre);
      Logger.log('✅ Hoja creada: ' + nombre);
    } else {
      Logger.log('⚠️  Hoja ya existe, se omite creación: ' + nombre);
    }

    // Encabezados — solo si la fila 1 está vacía
    const primeraCelda = hoja.getRange(1, 1).getValue();
    if (!primeraCelda) {
      hoja.getRange(1, 1, 1, columnas.length)
          .setValues([columnas])
          .setFontWeight('bold')
          .setBackground('#1E3A5F')
          .setFontColor('#FFFFFF');
      hoja.setFrozenRows(1);
      Logger.log('  → Encabezados escritos en ' + nombre);
    } else {
      Logger.log('  → Encabezados ya presentes en ' + nombre + ', se omiten');
    }

    // Datos de prueba (solo si la hoja tiene seed y está vacía en fila 2)
    if (SEED[nombre]) {
      const filasDatos = hoja.getLastRow();
      if (filasDatos < 2) {
        hoja.getRange(2, 1, SEED[nombre].length, SEED[nombre][0].length)
            .setValues(SEED[nombre]);
        Logger.log('  → Datos de prueba cargados en ' + nombre);
      } else {
        Logger.log('  → ' + nombre + ' ya tiene datos, se omiten datos de prueba');
      }
    }

    // Ancho de columnas automático
    hoja.autoResizeColumns(1, columnas.length);
  });

  Logger.log('');
  Logger.log('══════════════════════════════════════');
  Logger.log('Setup de Vale Digital completado ✅');
  Logger.log('Revisá las hojas VALES, USUARIOS y OBRAS.');
  Logger.log('══════════════════════════════════════');
}
