/**
 * SCRIPT DE BACKUP DIARIO - SALES ORDER MASTER (VERSIÓN FINAL)
 *
 * Este script crea automáticamente una copia diaria del archivo "Sales Order 2025 Master"
 * (que ahora es el archivo principal de registro).
 * La copia se nombra según la fecha.
 *
 * CONFIGURACIÓN INICIAL:
 * 1. Pegar este script en el mismo proyecto de Google Apps Script.
 * 2. Configurar los IDs y carpetas.
 * 3. Ejecutar configuracionBackupDiario() una vez.
 */

// ===== CONFIGURACIÓN BACKUP DIARIO (SIMPLIFICADA SIN RESALTADO) =====
const BACKUP_CONFIG = {
  // ID del archivo FUENTE para el backup (AHORA ES EL MASTER PRINCIPAL)
  MASTER_FILE_ID: '1U8yZPOowD7Z9JhR4IpvT9_wRKdpc5eqdxjKQxeinrBU', // ID del archivo Sales Order 2025 Master

  // Nombres exactos de las hojas y configuración de encabezados del Master
  MASTER_SHEET_NAME: 'Original Data', // Nombre de la hoja principal en el Master
  MASTER_HEADER_ROWS: 1, // CONFIRMADO: Datos del Master empiezan en fila 2 (es decir, 1 fila de encabezado)
  MASTER_RESERVATION_DATE_COLUMN: 3, // Columna C: 'Date' - (se mantiene para referencia si es útil, pero no se usa para resaltar)

  // Configuración de nombres y ubicación del backup
  CARPETA_BACKUP_ID: '1yIuPm9t8UnwJSWKYJYAyD9_iT8ilrHZt', // null = misma carpeta del Master, o ID de carpeta específica para backups
  PREFIJO_NOMBRE_BACKUP: 'Sales Order Master', // Nombre que tendrá la copia: "Sales Order - 2025-06-10"
  FORMATO_FECHA: 'YYYY-MM-DD', // Resultado: "Sales Order - 2025-06-10"

  // Configuración de limpieza del backup (la copia del Master)
  // Se mantiene esta opción por si el Master tuviera otras hojas o marcas inesperadas.
  MANTENER_SOLO_MASTER_DATA_EN_COPIA: true, // Solo mantener hoja principal en la copia del Master.

  // Notificaciones
  NOTIFICATION_EMAIL: 'juanferojas19@gmail.com',
  NOTIFICAR_BACKUP_CREADO: true,
  NOTIFICAR_ERRORES: true,

  // Configuración de ejecución
  HORA_EJECUCION: 23, // 11 PM
  MINUTO_EJECUCION: 45, // 11:45 PM

  // Gestión de archivos antiguos
  ELIMINAR_BACKUPS_ANTIGUOS: true,
  DIAS_CONSERVAR_BACKUP: 30 // Conservar backups de los últimos 30 días
};

// ===== FUNCIÓN PRINCIPAL DEL BACKUP =====
/**
 * Crea el backup diario del Master.
 */
function crearBackupDiario() {
  try {
    console.log('📂 Iniciando backup diario del Master...');

    // Obtener fecha actual en formato requerido
    const fechaHoy = new Date();
    const fechaFormateada = formatearFecha(fechaHoy);
    const nombreBackup = `${fechaFormateada} - ${BACKUP_CONFIG.PREFIJO_NOMBRE_BACKUP}`; // CAMBIO AQUÍ

    console.log(`📅 Creando backup: "${nombreBackup}"`);

    // Abrir archivo fuente (AHORA ES EL MASTER PRINCIPAL)
    const archivoFuente = SpreadsheetApp.openById(BACKUP_CONFIG.MASTER_FILE_ID);
    console.log(`📁 Archivo fuente (Master): "${archivoFuente.getName()}"`);

    // Obtener la carpeta destino para el backup
    const carpetaDestino = obtenerCarpetaDestino(archivoFuente);
    const backupExistente = buscarBackupExistente(carpetaDestino, nombreBackup);

    if (backupExistente) {
      console.log('⚠️ Ya existe backup de hoy, eliminando el anterior...');
      Drive.Files.remove(backupExistente.getId());
    }

    // Crear copia del archivo Master
    console.log('🔄 Creando copia del archivo Master...');
    const archivoCopia = archivoFuente.copy(nombreBackup);

    // Mover a carpeta específica si está configurada
    if (BACKUP_CONFIG.CARPETA_BACKUP_ID) {
      const carpetaBackup = DriveApp.getFolderById(BACKUP_CONFIG.CARPETA_BACKUP_ID);
      DriveApp.getFileById(archivoCopia.getId()).moveTo(carpetaBackup);
      console.log(`📁 Archivo movido a carpeta de backup`);
    }

    // Limpiar archivo backup (la copia del Master: solo mantener hoja principal si aplica)
    console.log('🧹 Limpiando archivo backup (la copia del Master)...');
    limpiarArchivoBackup(archivoCopia);

    // NOTA: La función de resaltado ya NO se llama aquí, según la nueva instrucción del jefe.

    // Eliminar backups antiguos si está configurado
    if (BACKUP_CONFIG.ELIMINAR_BACKUPS_ANTIGUOS) {
      console.log('🗑️ Eliminando backups antiguos...');
      eliminarBackupsAntiguos(carpetaDestino);
    }

    // Información del backup creado
    const urlBackup = `https://docs.google.com/spreadsheets/d/${archivoCopia.getId()}`;
    const estadisticas = obtenerEstadisticasBackup(archivoCopia);

    console.log('✅ Backup del Master creado exitosamente');
    console.log(`📊 Estadísticas: ${estadisticas.filas} filas, ${estadisticas.hojas} hojas`);

    // Enviar notificación de backup
    if (BACKUP_CONFIG.NOTIFICAR_BACKUP_CREADO) {
      const mensaje = `
✅ BACKUP DIARIO DEL MASTER CREADO EXITOSAMENTE

📂 Archivo: ${nombreBackup}
📅 Fecha: ${fechaHoy.toLocaleDateString()}
⏰ Hora: ${new Date().toLocaleTimeString()}

📊 Contenido:
• ${estadisticas.filas} filas de datos (Master completo)
• ${estadisticas.hojas} hojas
• Archivo listo para descarga CSV (si aplica)

🔗 Enlace directo: ${urlBackup}

💾 Este archivo sirve como backup histórico.

🔄 Próximo backup programado: Mañana a las ${BACKUP_CONFIG.HORA_EJECUCION}:${BACKUP_CONFIG.MINUTO_EJECUCION.toString().padStart(2, '0')}
      `;

      enviarNotificacionBackup('✅ Backup Diario del Master Creado', mensaje);
    }

    return {
      exito: true,
      archivo: nombreBackup,
      url: urlBackup,
      estadisticas: estadisticas
    };

  } catch (error) {
    console.error('❌ Error creando backup:', error);

    if (BACKUP_CONFIG.NOTIFICAR_ERRORES) {
      enviarNotificacionBackup(
        '🚨 Error en Backup Diario',
        `Error al crear backup del Master: ${error.message}\n\nFecha: ${new Date()}`
      );
    }

    return {
      exito: false,
      error: error.message
    };
  }
}

// ===== FUNCIONES AUXILIARES =====
/**
 * Formatea la fecha según configuración
 */
function formatearFecha(fecha) {
  const año = fecha.getFullYear();
  const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
  const dia = fecha.getDate().toString().padStart(2, '0');

  switch (BACKUP_CONFIG.FORMATO_FECHA) {
    case 'YYYY-MM-DD':
      return `${año}-${mes}-${dia}`;
    case 'DD-MM-YYYY':
      return `${dia}-${mes}-${año}`;
    case 'MM-DD-YYYY':
      return `${mes}-${dia}-${año}`;
    default:
      return `${año}-${mes}-${dia}`;
  }
}

/**
 * Obtiene la carpeta destino para el backup
 */
function obtenerCarpetaDestino(archivoFuente) {
  if (BACKUP_CONFIG.CARPETA_BACKUP_ID) {
    return DriveApp.getFolderById(BACKUP_CONFIG.CARPETA_BACKUP_ID);
  } else {
    // Usar la misma carpeta del archivo fuente (Master)
    const archivoDrive = DriveApp.getFileById(archivoFuente.getId());
    const parents = archivoDrive.getParents();
    return parents.hasNext() ? parents.next() : DriveApp.getRootFolder(); // Retorna la primera carpeta padre o la raíz
  }
}

/**
 * Busca si ya existe un backup con el mismo nombre
 */
function buscarBackupExistente(carpeta, nombreBackup) {
  const archivos = carpeta.getFilesByName(nombreBackup);
  return archivos.hasNext() ? archivos.next() : null;
}

/**
 * Limpia el archivo backup (la copia del Master) eliminando hojas extras.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} archivoCopia El objeto del archivo de copia creado.
 */
function limpiarArchivoBackup(archivoCopia) {
  try {
    const spreadsheet = SpreadsheetApp.openById(archivoCopia.getId());

    // Si solo queremos mantener la hoja principal del Master
    if (BACKUP_CONFIG.MANTENER_SOLO_MASTER_DATA_EN_COPIA) {
      const hojas = spreadsheet.getSheets();
      hojas.forEach(hoja => {
        if (hoja.getName() !== BACKUP_CONFIG.MASTER_SHEET_NAME) {
          spreadsheet.deleteSheet(hoja);
          console.log(`🗑️ Eliminada hoja: ${hoja.getName()} de la copia de backup.`);
        }
      });
    }

    SpreadsheetApp.flush(); // Asegurar que los cambios se apliquen

  } catch (error) {
    console.log('⚠️ Error limpiando backup (la copia del Master, continuando):', error.message);
  }
}

/**
 * Elimina backups antiguos según configuración
 */
function eliminarBackupsAntiguos(carpeta) {
  try {
    const archivos = carpeta.getFiles();
    const fechaLimite = new Date();
    fechaLimite.setDate(fechaLimite.getDate() - BACKUP_CONFIG.DIAS_CONSERVAR_BACKUP);

    let eliminados = 0;
    while (archivos.hasNext()) {
      const archivo = archivos.next();
      const nombreArchivo = archivo.getName();

      // Verificar si es un backup de Sales Order y si es más antiguo que la fecha límite
      if (nombreArchivo.startsWith(BACKUP_CONFIG.PREFIJO_NOMBRE_BACKUP) &&
        archivo.getDateCreated() < fechaLimite) {
        archivo.setTrashed(true); // Mover a la papelera
        eliminados++;
        console.log(`🗑️ Eliminado backup antiguo: ${nombreArchivo}`);
      }
    }

    if (eliminados > 0) {
      console.log(`🧹 Eliminados ${eliminados} backups antiguos`);
    }

  } catch (error) {
    console.log('⚠️ Error eliminando backups antiguos:', error.message);
  }
}

/**
 * Obtiene estadísticas del backup creado
 */
function obtenerEstadisticasBackup(archivoCopia) {
  try {
    const spreadsheet = SpreadsheetApp.openById(archivoCopia.getId());
    const hojas = spreadsheet.getSheets();

    let totalFilas = 0;
    hojas.forEach(hoja => {
      totalFilas += hoja.getLastRow();
    });

    return {
      hojas: hojas.length,
      filas: totalFilas,
      tamaño: DriveApp.getFileById(archivoCopia.getId()).getSize()
    };

  } catch (error) {
    return {
      hojas: 'N/A',
      filas: 'N/A',
      tamaño: 'N/A'
    };
  }
}

/**
 * Envía notificaciones de backup
 */
function enviarNotificacionBackup(asunto, mensaje) {
  try {
    MailApp.sendEmail({
      to: BACKUP_CONFIG.NOTIFICATION_EMAIL,
      subject: `🏨 Club Resort - ${asunto}`,
      body: mensaje
    });
    console.log('📧 Notificación de backup enviada');
  } catch (error) {
    console.error('❌ Error enviando notificación de backup:', error);
  }
}

// ===== FUNCIONES DE CONFIGURACIÓN =====
/**
 * Configuración inicial del sistema de backup
 * EJECUTAR UNA SOLA VEZ
 */
function configuracionBackupDiario() {
  console.log('🛠️ Configurando sistema de backup diario...');

  try {
    // Verificar configuración inicial
    if (BACKUP_CONFIG.MASTER_FILE_ID === '1U8yZPOowD7Z9JhR4IpvT9_wRKdpc5eqdxjKQxeinrBU') {
      console.warn('⚠️ ADVERTENCIA: MASTER_FILE_ID aún tiene el ID de ejemplo. Asegúrate de que sea tu ID real del Master.');
    }
    
    // Eliminar triggers existentes de backup para evitar duplicados
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'crearBackupDiario') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // Crear trigger diario
    ScriptApp.newTrigger('crearBackupDiario')
      .timeBased()
      .everyDays(1)
      .atHour(BACKUP_CONFIG.HORA_EJECUCION)
      .nearMinute(BACKUP_CONFIG.MINUTO_EJECUCION)
      .create();

    console.log(`⏰ Backup programado diariamente a las ${BACKUP_CONFIG.HORA_EJECUCION}:${BACKUP_CONFIG.MINUTO_EJECUCION.toString().padStart(2, '0')}`);

    // Crear backup inicial para probar
    console.log('🧪 Creando backup inicial de prueba...');
    const resultado = crearBackupDiario();

    if (resultado.exito) {
      console.log('✅ Sistema de backup configurado exitosamente');
    } else {
      console.log('❌ Error en backup de prueba:', resultado.error);
    }

  } catch (error) {
    console.error('❌ Error en configuración:', error);
  }
}

/**
 * Crear backup manual (para testing o uso excepcional)
 */
function crearBackupManual() {
  console.log('🔧 Backup manual iniciado...');
  return crearBackupDiario();
}

/**
 * Probar configuración de backup
 */
function probarConfiguracionBackup() {
  console.log('🧪 Probando configuración de backup...');

  try {
    // Verificar acceso al archivo Master
    const masterFile = SpreadsheetApp.openById(BACKUP_CONFIG.MASTER_FILE_ID);
    console.log('✅ Archivo Master accesible:', masterFile.getName());

    // Verificar carpeta destino
    if (BACKUP_CONFIG.CARPETA_BACKUP_ID) {
      const carpeta = DriveApp.getFolderById(BACKUP_CONFIG.CARPETA_BACKUP_ID);
      console.log('✅ Carpeta backup accesible:', carpeta.getName());
    } else {
      console.log('✅ Usando misma carpeta del archivo Master como destino de backup');
    }

    // Mostrar configuración actual
    console.log('\n📋 CONFIGURACIÓN ACTUAL:');
    console.log(`ID Archivo Master (fuente): ${BACKUP_CONFIG.MASTER_FILE_ID}`);
    console.log(`Hoja principal Master: ${BACKUP_CONFIG.MASTER_SHEET_NAME}`);
    console.log(`Filas de encabezado Master: ${BACKUP_CONFIG.MASTER_HEADER_ROWS}`);
    console.log(`Columna Fecha de Reserva Master: ${BACKUP_CONFIG.MASTER_RESERVATION_DATE_COLUMN} (ya no se usa para resaltar)`);
    console.log(`Prefijo nombre backup: ${BACKUP_CONFIG.PREFIJO_NOMBRE_BACKUP}`);
    console.log(`Formato fecha: ${BACKUP_CONFIG.FORMATO_FECHA}`);
    console.log(`Hora ejecución: ${BACKUP_CONFIG.HORA_EJECUCION}:${BACKUP_CONFIG.MINUTO_EJECUCION.toString().padStart(2, '0')}`);
    console.log(`Email de notificación: ${BACKUP_CONFIG.NOTIFICATION_EMAIL}`);
    console.log(`Días a conservar backups: ${BACKUP_CONFIG.DIAS_CONSERVAR_BACKUP}`);

    console.log('\n✅ Configuración correcta - Lista para activar');

  } catch (error) {
    console.log('❌ Error en configuración:', error.message);
  }
}

/**
 * Desactivar sistema de backup
 */
function desactivarBackupDiario() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'crearBackupDiario') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  console.log('🛑 Sistema de backup desactivado');
}