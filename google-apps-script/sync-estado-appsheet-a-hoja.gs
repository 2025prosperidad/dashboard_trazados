/**
 * Sincroniza la columna Status de AppSheet (tabla Intake_form) hacia ESTADO_ACTUAL
 * en Google Sheets, vía API v2 de AppSheet.
 *
 * Por qué a veces ves "402 filas de AppSheet" pero "153 actualizadas":
 *
 * 1) El contador "actualizadas" solo suma filas donde el valor NUEVO es distinto
 *    al que ya había en la hoja. Si 249 trámites ya tenían el mismo ESTADO_ACTUAL,
 *    el log mostraría 153 aunque se hayan revisado 402.
 *
 * 2) "402" debe ser el largo real de response.Rows del JSON. Si logueas otro
 *    número (p. ej. un total de la app), puede no coincidir con lo que recorres.
 *
 * 3) Si buscas filas en la hoja solo hasta getLastRow() y el rango de IDs no
 *    cubre todos los expedientes, solo se actualizarán las que encuentres.
 *
 * Ajusta CONFIG y ejecuta syncEstadoDesdeAppSheet() o programa syncEstadoDesdeAppSheet
 * cada minuto en un disparador instalado.
 *
 * Dónde sacar APP_ID y APPLICATION_ACCESS_KEY (AppSheet):
 *   Editor de la app → menú "Data" → sección API / "Integrations" → habilitar API
 *   → copiar "App ID" y "Application Access Key" (no uses el texto TU_APP_ID de ejemplo).
 */

const CONFIG = {
  /** Reemplaza por el App ID real (UUID o cadena que muestra AppSheet). Si dejas TU_APP_ID, la API responde 400 "not found". */
  APP_ID: 'TU_APP_ID',
  /** Clave larga tipo V2-xxxxx... del mismo pantalla API de AppSheet */
  APPLICATION_ACCESS_KEY: 'TU_APPLICATION_ACCESS_KEY',
  /** Dominio según tu cuenta (global: www.appsheet.com; EU: eu.appsheet.com; etc.) */
  REGION_HOST: 'www.appsheet.com',
  /** Nombre exacto de la tabla en AppSheet (URL-encoded si tiene espacios) */
  TABLE_NAME: 'Intake_form',
  /** Columna en AppSheet que trae el estado (nombre exacto) */
  COLUMNA_STATUS_APPSHEET: 'Status',
  /** ID de la hoja de Google */
  SPREADSHEET_ID: 'TU_SPREADSHEET_ID',
  /** Nombre de la pestaña donde está Intake_form */
  NOMBRE_HOJA: 'Intake_form',
  /** Letra de la columna que es CLAVE para alinear con AppSheet (misma que en el dashboard, p. ej. "id") */
  COLUMNA_ID_EN_HOJA: 'A',
  /** Letra de la columna ESTADO_ACTUAL en la hoja */
  COLUMNA_ESTADO_EN_HOJA: 'B',
  /** Fila donde empiezan los datos (1 = cabecera en fila 1) */
  FILA_ENCABEZADO: 1,
};

/**
 * Evita llamar a AppSheet con los marcadores de plantilla (error 400 "App ... not found").
 */
function assertConfigFilled_() {
  const errs = [];
  if (!CONFIG.APP_ID || CONFIG.APP_ID === 'TU_APP_ID') {
    errs.push('CONFIG.APP_ID');
  }
  if (!CONFIG.APPLICATION_ACCESS_KEY || CONFIG.APPLICATION_ACCESS_KEY === 'TU_APPLICATION_ACCESS_KEY') {
    errs.push('CONFIG.APPLICATION_ACCESS_KEY');
  }
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID === 'TU_SPREADSHEET_ID') {
    errs.push('CONFIG.SPREADSHEET_ID');
  }
  if (errs.length) {
    throw new Error(
      'Configura en Código.gs los valores reales de: ' +
        errs.join(', ') +
        '. El error "App with id TU_APP_ID not found" significa que aún está el ejemplo: ' +
        'pega el App ID que muestra AppSheet (Data → API), guarda el proyecto y vuelve a ejecutar.'
    );
  }
}

/**
 * POST Find a AppSheet y devuelve el array Rows (o []).
 */
function fetchRowsFromAppSheet_() {
  const tableEnc = encodeURIComponent(CONFIG.TABLE_NAME);
  const url =
    'https://' +
    CONFIG.REGION_HOST +
    '/api/v2/apps/' +
    CONFIG.APP_ID +
    '/tables/' +
    tableEnc +
    '/Action';

  const payload = {
    Action: 'Find',
    Properties: {
      Locale: 'es-EC',
      Timezone: 'America/Guayaquil',
    },
    Rows: [],
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { ApplicationAccessKey: CONFIG.APPLICATION_ACCESS_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    var hint = '';
    if (code === 400 && text.indexOf('not found') !== -1 && String(CONFIG.APP_ID).indexOf('TU_') === 0) {
      hint =
        ' Parece que APP_ID sigue siendo de ejemplo; reemplázalo por el App ID real en CONFIG.';
    }
    throw new Error('AppSheet HTTP ' + code + ': ' + text.slice(0, 500) + hint);
  }

  const json = JSON.parse(text);
  const rows = json.Rows || json.rows || [];
  return rows;
}

/**
 * Construye Mapa id -> Status (string).
 * AppSheet puede usar distintos nombres de clave: ajusta las claves probadas.
 */
function mapStatusById_(appRows) {
  const map = {};
  const statusKey = CONFIG.COLUMNA_STATUS_APPSHEET;

  appRows.forEach(function (row) {
    const id =
      row.id != null && row.id !== ''
        ? String(row.id).trim()
        : row.ID != null && row.ID !== ''
          ? String(row.ID).trim()
          : row._RowNumber != null
            ? String(row._RowNumber).trim()
            : '';

    if (!id) return;

    const st = row[statusKey];
    map[id] = st == null ? '' : String(st).trim();
  });

  return map;
}

/**
 * Sincronización principal: lee AppSheet, recorre la hoja por ID y escribe ESTADO_ACTUAL.
 */
function syncEstadoDesdeAppSheet() {
  assertConfigFilled_();
  const appRows = fetchRowsFromAppSheet_();
  Logger.log('Filas en response.Rows (AppSheet): ' + appRows.length);

  const statusById = mapStatusById_(appRows);
  Logger.log('IDs únicos con Status en mapa: ' + Object.keys(statusById).length);

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.NOMBRE_HOJA);
  if (!sh) throw new Error('No existe la hoja: ' + CONFIG.NOMBRE_HOJA);

  const lastRow = sh.getLastRow();
  if (lastRow <= CONFIG.FILA_ENCABEZADO) {
    Logger.log('No hay datos debajo del encabezado.');
    return;
  }

  const startRow = CONFIG.FILA_ENCABEZADO + 1;
  const numRows = lastRow - CONFIG.FILA_ENCABEZADO;
  const idRange = sh.getRange(CONFIG.COLUMNA_ID_EN_HOJA + startRow + ':' + CONFIG.COLUMNA_ID_EN_HOJA + lastRow);
  const estadoRange = sh.getRange(CONFIG.COLUMNA_ESTADO_EN_HOJA + startRow + ':' + CONFIG.COLUMNA_ESTADO_EN_HOJA + lastRow);

  const ids = idRange.getValues();
  const estadosActuales = estadoRange.getValues();

  let matched = 0;
  let written = 0;
  let unchanged = 0;
  let notInAppSheet = 0;
  const out = [];

  for (let i = 0; i < numRows; i++) {
    const rawId = ids[i][0];
    const id = rawId == null || rawId === '' ? '' : String(rawId).trim();
    const prev = estadosActuales[i][0] == null ? '' : String(estadosActuales[i][0]).trim();

    if (!id) {
      out.push([prev]);
      continue;
    }

    if (!Object.prototype.hasOwnProperty.call(statusById, id)) {
      notInAppSheet++;
      out.push([prev]);
      continue;
    }

    matched++;
    const nuevo = statusById[id];
    if (nuevo !== prev) {
      written++;
      out.push([nuevo]);
    } else {
      unchanged++;
      out.push([prev]);
    }
  }

  estadoRange.setValues(out);

  Logger.log('Filas en hoja (rango ID): ' + numRows);
  Logger.log('Coincidencias ID AppSheet↔hoja: ' + matched);
  Logger.log('Celdas con valor distinto (escrituras): ' + written);
  Logger.log('Sin cambio (ya igual): ' + unchanged);
  Logger.log('ID en hoja sin match en AppSheet: ' + notInAppSheet);
}

/**
 * Ejecuta una vez syncEstadoDesdeAppSheet y deja el detalle en Registro de ejecución.
 * Si "escrituras" << "Filas en response.Rows", revisa:
 * - COLUMNA_ID_EN_HOJA debe ser la misma clave que usa el dashboard (campo `id`).
 * - COLUMNA_STATUS_APPSHEET debe coincidir exactamente con el nombre en AppSheet.
 * - Si muchas filas quedan "sin cambio", el script está bien: ya estaban sincronizadas.
 */
function ejecutarSyncConDiagnostico() {
  syncEstadoDesdeAppSheet();
}
