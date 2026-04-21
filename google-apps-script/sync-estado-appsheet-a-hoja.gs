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
 *
 * Si el log muestra "Filas en response.Rows (AppSheet): 0" pero la app sí tiene datos:
 *   - Revisa el nombre de TABLE_NAME (debe coincidir con Data → Tables en AppSheet).
 *   - Muy frecuente: Security filter en la tabla que con la identidad por defecto del API
 *     no devuelve filas. Pon RUN_AS_USER_EMAIL con un usuario que vea todos los trámites
 *     (p. ej. el dueño o un supervisor), según documentación RunAsUserEmail.
 */

const CONFIG = {
  APP_ID: '156f5e61-3921-4762-9710-87ffa1f49619',
  /**
   * Clave V2-... de AppSheet (Data → API). Preferible: déjala vacía y define la propiedad
   * de script APP_SHEET_APPLICATION_ACCESS_KEY (Proyecto → engrane → Propiedades del script)
   * para no guardar secretos en texto plano si compartes el código.
   */
  APPLICATION_ACCESS_KEY: '',
  /** Dominio según tu cuenta (global: www.appsheet.com; EU: eu.appsheet.com; etc.) */
  REGION_HOST: 'www.appsheet.com',
  /** Nombre exacto de la tabla en AppSheet (URL-encoded si tiene espacios) */
  TABLE_NAME: 'Intake_form',
  /** Columna en AppSheet que trae el estado (nombre exacto) */
  COLUMNA_STATUS_APPSHEET: 'Status',
  /** ID de la hoja de Google (entre /d/ y /edit en la URL) */
  SPREADSHEET_ID: '1LaATbQJpXc7iA-BHh5ZWx41bB_T0UwpyOH8eTDyXo_o',
  /** Nombre de la pestaña donde está Intake_form */
  NOMBRE_HOJA: 'Intake_form',
  /** Letra de la columna que es CLAVE para alinear con AppSheet (misma que en el dashboard, p. ej. "id") */
  COLUMNA_ID_EN_HOJA: 'A',
  /** Letra de la columna ESTADO_ACTUAL en la hoja */
  COLUMNA_ESTADO_EN_HOJA: 'B',
  /** Fila donde empiezan los datos (1 = cabecera en fila 1) */
  FILA_ENCABEZADO: 1,

  /**
   * Find ejecutado como este usuario (Security filters / USEREMAIL() en tablas).
   * Creador de la app / usuario con visión completa en filtros.
   */
  RUN_AS_USER_EMAIL: 'expert@infinity-solutions.community',

  /**
   * Opcional: Selector AppSheet fijo (ej. 'Filter(Intake_form, true)').
   * Si está vacío, el script prueba sin Selector y luego Filter(TABLE_NAME, true).
   */
  FIND_SELECTOR: '',
};

/**
 * Clave de API: CONFIG o propiedad de script APP_SHEET_APPLICATION_ACCESS_KEY.
 */
function getApplicationAccessKey_() {
  const fromProps = PropertiesService.getScriptProperties().getProperty('APP_SHEET_APPLICATION_ACCESS_KEY');
  const fromConfig = (CONFIG.APPLICATION_ACCESS_KEY || '').trim();
  return (fromProps && String(fromProps).trim()) || fromConfig;
}

/**
 * Evita ejecutar sin App ID, hoja o clave de acceso.
 */
function assertConfigFilled_() {
  const errs = [];
  if (!CONFIG.APP_ID || CONFIG.APP_ID === 'TU_APP_ID') {
    errs.push('CONFIG.APP_ID');
  }
  if (!getApplicationAccessKey_()) {
    errs.push('CONFIG.APPLICATION_ACCESS_KEY o propiedad APP_SHEET_APPLICATION_ACCESS_KEY');
  }
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID === 'TU_SPREADSHEET_ID') {
    errs.push('CONFIG.SPREADSHEET_ID');
  }
  if (errs.length) {
    throw new Error(
      'Configura en Código.gs: ' +
        errs.join(', ') +
        '. Para la clave: pégala en APPLICATION_ACCESS_KEY o crea la propiedad del script APP_SHEET_APPLICATION_ACCESS_KEY.'
    );
  }
}

/**
 * Extrae el array de filas del JSON de respuesta AppSheet (varias formas posibles).
 */
function extractRowsFromAppSheetJson_(json) {
  if (!json || typeof json !== 'object') return [];
  if (Array.isArray(json)) return json;
  if (Array.isArray(json.Rows)) return json.Rows;
  if (Array.isArray(json.rows)) return json.rows;
  if (json.data && Array.isArray(json.data.Rows)) return json.data.Rows;
  return [];
}

/**
 * Construye Properties del Find (Locale, opcional RunAsUserEmail, opcional Selector).
 */
function buildFindProperties_(selector) {
  const props = {
    Locale: 'es-EC',
    Timezone: 'America/Guayaquil',
  };
  const runAs = (CONFIG.RUN_AS_USER_EMAIL || '').trim();
  if (runAs.indexOf('@') !== -1) {
    props.RunAsUserEmail = runAs;
  }
  const sel = (selector || (CONFIG.FIND_SELECTOR || '').trim()).trim();
  if (sel) props.Selector = sel;
  return props;
}

/**
 * POST Find; devuelve { code, text, json } (json puede ser null si el cuerpo no es JSON).
 */
function appSheetFindRaw_(selector) {
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
    Properties: buildFindProperties_(selector),
    Rows: [],
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { ApplicationAccessKey: getApplicationAccessKey_() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();
  let json = null;
  try {
    json = JSON.parse(text);
  } catch (e) {
    json = null;
  }
  return { code: code, text: text, json: json };
}

/**
 * POST Find a AppSheet: varios intentos si Rows viene vacío (Selector / identidad).
 */
function fetchRowsFromAppSheet_() {
  const runAs = (CONFIG.RUN_AS_USER_EMAIL || '').trim();
  Logger.log('RunAsUserEmail en Find: ' + (runAs ? runAs : '(no definido — puede devolver 0 filas con security filters)'));

  const attempts = [];

  const fixedSel = (CONFIG.FIND_SELECTOR || '').trim();
  if (fixedSel) {
    attempts.push({ label: 'FIND_SELECTOR configurado', selector: fixedSel });
  } else {
    attempts.push({ label: 'Find sin Selector (documentación “Read all rows”)', selector: null });
    attempts.push({
      label: 'Find con Filter(TABLE_NAME, true)',
      selector: 'Filter(' + CONFIG.TABLE_NAME + ', true)',
    });
  }

  let lastText = '';
  let lastCode = 0;

  for (let a = 0; a < attempts.length; a++) {
    const att = attempts[a];
    Logger.log('AppSheet Find: intento — ' + att.label);

    const res = appSheetFindRaw_(att.selector);
    lastCode = res.code;
    lastText = res.text;

    if (res.code < 200 || res.code >= 300) {
      var hint = '';
      if (res.code === 400 && res.text.indexOf('not found') !== -1 && String(CONFIG.APP_ID).indexOf('TU_') === 0) {
        hint =
          ' Parece que APP_ID sigue siendo de ejemplo; reemplázalo por el App ID real en CONFIG.';
      }
      throw new Error('AppSheet HTTP ' + res.code + ': ' + res.text.slice(0, 500) + hint);
    }

    if (!res.json) {
      Logger.log('La respuesta no es JSON válido (primeros 400 caracteres): ' + res.text.substring(0, 400));
      throw new Error('AppSheet devolvió cuerpo no JSON. HTTP ' + res.code);
    }

    const rows = extractRowsFromAppSheetJson_(res.json);
    Logger.log('  → filas extraídas: ' + rows.length);

    if (rows.length > 0) {
      return rows;
    }

    Logger.log('  → claves en JSON: ' + JSON.stringify(Object.keys(res.json)));
  }

  Logger.log(
    'DIAGNÓSTICO: Todos los intentos devolvieron 0 filas. Muestra del cuerpo (max 1200 chars):\n' +
      lastText.substring(0, 1200)
  );
  Logger.log(
    'Si la app tiene datos: (1) Verifica TABLE_NAME = nombre exacto en Data → Tables. ' +
      '(2) Activa CONFIG.RUN_AS_USER_EMAIL con un usuario que vea la tabla (Security filter / USEREMAIL()). ' +
      '(3) Prueba en Postman el mismo Find y revisa Security filters de ' +
      CONFIG.TABLE_NAME +
      '.'
  );

  return [];
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
