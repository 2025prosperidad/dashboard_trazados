/* ==========================================
   API ORTEGON v2.0 - Conexion
   Dashboard Trazados Viales - Prefectura de Pichincha
   ========================================== */

const API_CONFIG = {
    endpoint: 'https://script.google.com/macros/s/AKfycbwGgHmLz3lTpYZ0eiinmOz8Fgmdnhq36s9mUlYDLjxbedCb1PRx_uBJ3WbLFhTec4rj/exec',
    apiKey: 'ortegon-2025-CAMBIA-ESTO',
    /** Evita pantalla de carga infinita si la red o GAS no responden */
    timeoutMs: 60000
};

/**
 * Usuarios de BD2 TRAZADOS VIALES APP PICHINCHA
 * Hoja: Usuarios (spreadsheet: 1aGRbWiozcDoCEEzPRKi2Jx11CzbroGB74plasRTlO44)
 * Nota: La API actual solo apunta a BD1. Cuando se actualice el GAS
 * para soportar multiples spreadsheets, reemplazar esto por una llamada API.
 */
const USUARIOS_BD2 = [
    { ID_Usuario: 1, Nombre: 'BENITEZ CARRILLO SONIA ELIZABETH', Rol: 'Ventanilla', Email: 'sbenitez@pichincha.gob.ec' },
    { ID_Usuario: 4, Nombre: 'ASIMBAYA SOCASI KATTY VANESSA', Rol: 'Tecnico', Email: 'asocaci@pichincha.gob.ec' },
    { ID_Usuario: 5, Nombre: 'GARCIA CANDO CRISTINA PAOLA', Rol: 'Tecnico', Email: 'cpgarcia@pichincha.gob.ec' },
    { ID_Usuario: 6, Nombre: 'GUALPA DIAZ CRISTIAN PATRICIO', Rol: 'Tecnico', Email: 'cgualpa@pichincha.gob.ec' },
    { ID_Usuario: 7, Nombre: 'VARGAS BARRERO EVELYN LORENA', Rol: 'Tecnico', Email: 'evargas@pichincha.gob.ec' },
    { ID_Usuario: 8, Nombre: 'SILVA CISNEROS SANTIAGO DAVID', Rol: 'Tecnico', Email: 'ssilva@pichincha.gob.ec' },
    { ID_Usuario: 9, Nombre: 'MORILLO OCHOA DANIELA NICOLE', Rol: 'Tecnico', Email: 'dmorillo@pichincha.gob.ec' },
    { ID_Usuario: 10, Nombre: 'NORONA MEDINA CARLA MONSERRATH', Rol: 'Tecnico', Email: 'cnorona@pichincha.gob.ec' },
    { ID_Usuario: 11, Nombre: 'CABEZAS MARCILLO IRENE ABIGAIL', Rol: 'Tecnico', Email: 'icabezas@pichincha.gob.ec' },
    { ID_Usuario: 12, Nombre: 'OSEJO CARDENAS JAIRO RAMIRO', Rol: 'Tecnico', Email: 'josejo@pichincha.gob.ec' },
    { ID_Usuario: 14, Nombre: 'MINO CHAVEZ JEFFERSON ROMARIO', Rol: 'Supervisor', Email: 'jrmino@pichincha.gob.ec' },
    { ID_Usuario: 15, Nombre: 'RIVERA ALVAREZ ALEX IVAN', Rol: 'Supervisor', Email: 'irivera@pichincha.gob.ec' },
    { ID_Usuario: 16, Nombre: 'EXPERT', Rol: 'Administrador', Email: 'expert@infinity-solutions.community' },
    { ID_Usuario: 17, Nombre: 'SANGOLUISA LLUMIQUINGA ANDREA', Rol: 'Supervisor', Email: 'asangoluisa@pichincha.gob.ec' },
    { ID_Usuario: 18, Nombre: 'LOAIZA TOSCANO DORIS SORAYA', Rol: 'Ventanilla', Email: 'dloaiza@pichincha.gob.ec' }
];

/**
 * Fetch data from API Ortegon v2.0
 */
function fetchWithTimeout(url, options = {}) {
    const ms = API_CONFIG.timeoutMs || 25000;
    const ctrl = new AbortController();
    const id = setTimeout(() => ctrl.abort(), ms);
    return fetch(url, { ...options, signal: ctrl.signal }).finally(() => clearTimeout(id));
}

async function apiRequest(action, params = {}) {
    const queryParams = new URLSearchParams({
        action: action,
        key: API_CONFIG.apiKey,
        ...params
    });

    const url = `${API_CONFIG.endpoint}?${queryParams.toString()}`;

    try {
        const response = await fetchWithTimeout(url, {
            method: 'GET',
            redirect: 'follow'
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const data = await response.json();

        if (data.success === false) {
            throw new Error(data.message || 'Error en la API');
        }

        return data;
    } catch (error) {
        const name = error && error.name === 'AbortError' ? 'Timeout' : (error && error.name);
        console.error(`API Error (${action}):`, name || error, error && error.message);
        throw error;
    }
}

async function getData(sheetName) {
    return apiRequest('export', { sheet: sheetName });
}

/**
 * Exporta filas de una hoja; si falla (hoja no en GAS o nombre distinto), devuelve [].
 * Hojas Equipos / Usuario_equipo (BD1): Id_Equipos = código fase (TV-01); ID_Usuario (gmail) + ID_Equipo en Usuario_equipo.
 */
async function getSheetRows(sheetName) {
    try {
        const res = await getData(sheetName);
        return Array.isArray(res?.data) ? res.data : [];
    } catch (e) {
        console.warn('[API] Hoja no cargada o no configurada:', sheetName, e.message || e);
        return [];
    }
}

/**
 * Si la primera variante del nombre devuelve vacío, prueba alias (espacios, mayúsculas).
 * Las hojas Der_Cat1 / Der_Cat2 deben estar permitidas en el script GAS (action export).
 */
async function getSheetRowsFirstNonEmpty(names) {
    let last = [];
    for (const name of names) {
        const rows = await getSheetRows(name);
        last = rows;
        if (Array.isArray(rows) && rows.length > 0) return rows;
    }
    return Array.isArray(last) ? last : [];
}

/**
 * Fetch all required dashboard data
 */
async function fetchAllDashboardData() {
    try {
        // Llamadas secuenciales para no saturar el GAS (evita AbortError por timeout)
        const intakeData  = await getData('Intake_form');
        const fasesData   = await getData('Fases_Tramite');
        const tiemposData = await getData('Tiempos_Fases');
        const tipoData    = await getData('Tipo_Tramite');
        const equiposRows = await getSheetRows('Equipos');
        const ueRowsAlt   = await getSheetRows('Usuario_equipo');
        const derCat1Rows = await getSheetRows('Der_Cat1');
        const derCat2Rows = await getSheetRows('Der_Cat2');

        const asArray = (res) => (Array.isArray(res?.data) ? res.data : []);
        let usuarioEquipo = ueRowsAlt;
        if (!usuarioEquipo.length) {
            usuarioEquipo = await getSheetRows('Usuario_Equipo');
        }

        let der1 = Array.isArray(derCat1Rows) ? derCat1Rows : [];
        let der2 = Array.isArray(derCat2Rows) ? derCat2Rows : [];
        if (!der1.length) {
            der1 = await getSheetRowsFirstNonEmpty(['Der Cat1', 'DER_CAT1', 'DerCat1']);
        }
        if (!der2.length) {
            der2 = await getSheetRowsFirstNonEmpty(['Der Cat2', 'DER_CAT2', 'DerCat2']);
        }

        // Segunda tanda: usuarios (no critico, carga despues)
        const usuariosRows = await getSheetRows('Usuarios');

        // Merge: API users take priority; fill gaps with USUARIOS_BD2 hardcoded list
        const apiUsers = Array.isArray(usuariosRows) ? usuariosRows : [];
        const apiEmails = new Set(apiUsers.map(u => {
            const em = u['Email (Ejemplo)'] || u.Email || u.email || '';
            return String(em).trim().toLowerCase();
        }).filter(Boolean));
        const bd2Complement = USUARIOS_BD2.filter(u => {
            const em = String(u.Email || '').trim().toLowerCase();
            return em && !apiEmails.has(em);
        });
        const usuarios = [...apiUsers, ...bd2Complement];

        return {
            intake: asArray(intakeData),
            fases: asArray(fasesData),
            tiempos: asArray(tiemposData),
            tipos: asArray(tipoData),
            equipos: Array.isArray(equiposRows) ? equiposRows : [],
            usuarioEquipo,
            usuarios,
            derCat1: der1,
            derCat2: der2
        };
    } catch (error) {
        console.error('Error fetching dashboard data:', error);
        return getFallbackData();
    }
}

/**
 * Fallback data
 */
function getFallbackData() {
    return {
        intake: [
            { id: 'aa40ce72', Fase_del_Tramite: 'RV-01', Name: 'UZCATEGUI CALISTO TEOFILO FABIAN', Responsible: 'evargas@pichincha.gob.ec', PARROQUIA: 'PUEMBO', DEX: 'GADPP-DVIA-CPVIA-2026-0128-DEX', 'Start date': '2026-02-10T08:00:00.000Z', Level_0: '3', Sumilla_Inicial: 'Media', ESTADO_ACTUAL: '', Estado: '' },
            { id: '45d218d9', Fase_del_Tramite: 'TV-01', Name: 'ALMEIDA VASQUEZ VICTOR OSWALDO', Responsible: 'asocaci@pichincha.gob.ec', PARROQUIA: 'CALACALI', DEX: 'GADPP-DVIA-CPVIA-2026-0129-DEX', 'Start date': '2026-02-11T08:00:00.000Z', Level_0: '1', Sumilla_Inicial: 'Media', ESTADO_ACTUAL: '', Estado: '' },
            { id: 'c053e099', Fase_del_Tramite: 'RV-01', Name: 'SILVA MORENO SANTIAGO', Responsible: 'ssilva@pichincha.gob.ec', PARROQUIA: 'AMAGUANA', DEX: 'GADPP-DVIA-CPVIA-2026-0130-DEX', 'Start date': '2026-02-13T08:00:00.000Z', Level_0: '3', Sumilla_Inicial: 'Media', ESTADO_ACTUAL: '', Estado: '' },
            { id: '120402cf', Fase_del_Tramite: 'DCP-01', Name: 'GARCIA LEON PEDRO', Responsible: 'asocaci@pichincha.gob.ec', PARROQUIA: 'CONOCOTO', DEX: 'GADPP-DVIA-CPVIA-2026-0131-DEX', 'Start date': '2026-02-14T08:00:00.000Z', Level_0: '1', Sumilla_Inicial: 'Alta', ESTADO_ACTUAL: '', Estado: '' },
            { id: 'e03ccc6a', Fase_del_Tramite: 'CV-01', Name: 'GARCIA MARTINEZ CARLOS', Responsible: 'cgarcia@pichincha.gob.ec', PARROQUIA: 'TUMBACO', DEX: 'GADPP-DVIA-CPVIA-2026-0132-DEX', 'Start date': '2026-02-14T08:00:00.000Z', Level_0: '2', Sumilla_Inicial: 'Media', ESTADO_ACTUAL: '', Estado: '' },
            { id: '8e172ffb', Fase_del_Tramite: 'RV-01', Name: 'MORILLO DIAZ DANIEL', Responsible: 'dmorillo@pichincha.gob.ec', PARROQUIA: 'CUMBAYA', DEX: 'GADPP-DVIA-CPVIA-2026-0133-DEX', 'Start date': '2026-02-14T08:00:00.000Z', Level_0: '3', Sumilla_Inicial: 'Baja', ESTADO_ACTUAL: '', Estado: '' },
            { id: 'c0feeb65', Fase_del_Tramite: 'CV-01', Name: 'GUALPA RODRIGUEZ CRISTINA', Responsible: 'cgualpa@pichincha.gob.ec', PARROQUIA: 'PUEMBO', DEX: 'GADPP-DVIA-CPVIA-2026-0134-DEX', 'Start date': '2026-02-14T08:00:00.000Z', Level_0: '2', Sumilla_Inicial: 'Media', ESTADO_ACTUAL: '', Estado: '' }
        ],
        fases: [
            { Codigo: 'TV-01', Tipo: 'TV', Nombre_Fase: 'Mapa de Competencia', Orden: 1, Avance: 16.67 },
            { Codigo: 'TV-02', Tipo: 'TV', Nombre_Fase: 'Informe de Pertinencia', Orden: 2, Avance: 33.33 },
            { Codigo: 'TV-03', Tipo: 'TV', Nombre_Fase: 'Informe de Topografia', Orden: 3, Avance: 50 },
            { Codigo: 'TV-04', Tipo: 'TV', Nombre_Fase: 'Informe de Diseno Geometrico', Orden: 4, Avance: 66.67 },
            { Codigo: 'TV-05', Tipo: 'TV', Nombre_Fase: 'Informe de Trazado Vial', Orden: 5, Avance: 83.33 },
            { Codigo: 'TV-06', Tipo: 'TV', Nombre_Fase: 'Resolucion y Oficio', Orden: 6, Avance: 100 },
            { Codigo: 'CV-01', Tipo: 'CV', Nombre_Fase: 'Informe de Certificacion Vial', Orden: 1, Avance: 100 },
            { Codigo: 'RV-01', Tipo: 'RV', Nombre_Fase: 'Informe de Replanteo Vial', Orden: 1, Avance: 100 },
            { Codigo: 'STP-01', Tipo: 'STP', Nombre_Fase: 'Informe de Seccion Transversal Proyectada', Orden: 1, Avance: 100 },
            { Codigo: 'CEV-01', Tipo: 'CEV', Nombre_Fase: 'Informe de Competencia', Orden: 1, Avance: 33.33 },
            { Codigo: 'CEV-02', Tipo: 'CEV', Nombre_Fase: 'Informe de Pertinencia', Orden: 2, Avance: 66.67 },
            { Codigo: 'CEV-03', Tipo: 'CEV', Nombre_Fase: 'Informe de Colocacion del Eje Vial', Orden: 3, Avance: 100 },
            { Codigo: 'CI-01', Tipo: 'CI', Nombre_Fase: 'Informe de Colocacion de Infraestructura', Orden: 1, Avance: 100 },
            { Codigo: 'DCP-01', Tipo: 'DCP', Nombre_Fase: 'Informe de Factibilidad para Declaratoria de Camino Publico', Orden: 1, Avance: 100 }
        ],
        tiempos: [
            { id_tramite: 'aa40ce72', fase: 'RV-01', fecha_hora: '2026-02-11T01:17:11.000Z', fecha_hora_fin: '', 'ASIGNADO A': '' },
            { id_tramite: '45d218d9', fase: 'TV-01', fecha_hora: '2026-02-11T02:09:41.000Z', fecha_hora_fin: '', 'ASIGNADO A': 'asocaci@pichincha.gob.ec' },
            { id_tramite: 'c053e099', fase: 'RV-01', fecha_hora: '2026-02-13T23:48:42.000Z', fecha_hora_fin: '', 'ASIGNADO A': 'ssilva@pichincha.gob.ec' }
        ],
        tipos: [
            { Codigo: 'TV', Nombre: 'Trazado Vial' },
            { Codigo: 'CV', Nombre: 'Certificacion Vial' },
            { Codigo: 'RV', Nombre: 'Replanteo Vial' },
            { Codigo: 'STP', Nombre: 'Seccion Transversal Proyectada' },
            { Codigo: 'CEV', Nombre: 'Colocacion de Eje Vial' },
            { Codigo: 'CI', Nombre: 'Colocacion de Infraestructura' },
            { Codigo: 'DCP', Nombre: 'Factibilidad Declaratoria Camino Publico' }
        ],
        /** Mismo modelo que BD1: Id_Equipos = código de fase; Usuario_equipo usa gmail + ID_Equipo */
        equipos: [
            { Id_Equipos: 'TV-01', Nombre: 'Trazado Vial - Competencia', contador: 4 },
            { Id_Equipos: 'RV-01', Nombre: 'Replanteo', contador: 1 }
        ],
        usuarioEquipo: [
            { 'ID_Usuario (gmail)': 'asocaci@pichincha.gob.ec', ID_Equipo: 'TV-01' },
            { 'ID_Usuario (gmail)': 'ssilva@pichincha.gob.ec', ID_Equipo: 'RV-01' }
        ],
        usuarios: USUARIOS_BD2,
        derCat1: [
            { ID: 1, Nombre: 'Gobierno Autónomo Descentralizado' },
            { ID: 2, Nombre: 'Entidad pública nacional' }
        ],
        derCat2: [
            { ID: 1, Nombre: 'Municipio de Quito' },
            { ID: 2, Nombre: 'Ministerio de Obras Públicas' }
        ]
    };
}
