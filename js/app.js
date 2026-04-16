/* ==========================================
   DASHBOARD TRAZADOS VIALES
   Prefectura de Pichincha
   ========================================== */

// Global state
let dashboardData = null;
let currentPage = 'ranking';
let filtroFase = 'Todos';
let filtroTipo = 'Todos';
let filtroUsuario = 'Todos';

// Nuevos Filtros globales
let filtroInicioDesde = '';
let filtroInicioHasta = '';
let filtroCumplDesde = '';
let filtroCumplHasta = '';
let filtroCanton = 'Todos';
let filtroParroquia = 'Todos';
let filtroNombre = '';

// Colores basados en el logo de Prefectura de Pichincha
const TYPE_COLORS = {
    'TV': '#1A3C6E',   // Azul navy
    'CV': '#007B4F',   // Verde
    'RV': '#EAB308',   // Dorado
    'STP': '#E88B00',  // Naranja
    'CEV': '#A52422',  // Rojo vino
    'CI': '#1A7878',   // Teal
    'DCP': '#6B3A2A'   // Cafe
};

const TYPE_NAMES = {
    'TV': 'Trazado Vial',
    'CV': 'Certificacion Vial',
    'RV': 'Replanteo Vial',
    'STP': 'Seccion Transversal Proyectada',
    'CEV': 'Colocacion de Eje Vial',
    'CI': 'Colocacion de Infraestructura',
    'DCP': 'Factibilidad Declaratoria Camino Publico'
};

// Estado posibles de un tramite (6 estados oficiales)
const ESTADO_CONFIG = {
    'en_proceso': { label: 'En Proceso', color: '#E88B00', bg: '#FFF3E0' },
    'finalizado': { label: 'Finalizado', color: '#007B4F', bg: '#E8F5E9' },
    'detenido': { label: 'Detenido', color: '#A52422', bg: '#FFEBEE' },
    'en_derivacion': { label: 'En Derivación', color: '#7B1FA2', bg: '#F3E5F5' },
    'solicitud_info': { label: 'Solicitud de Información', color: '#1565C0', bg: '#E3F2FD' },
    'archivado': { label: 'Archivado', color: '#6B3A2A', bg: '#EFEBE9' }
};

/* ==========================================
   CHART.JS — Estilo BI (Looker Studio / Power BI)
   ========================================== */
let rankingTipoHBarInst = null;
let rankingRespHBarInst = null;
let tipoEstadoDonutInst = null;

/** Barras tipo Looker Studio (25th Ward / reportes Google) */
const LS_BAR = {
    total: '#90CAF9',
    en_proceso: '#FFCA28',
    finalizado: '#66BB6A',
    detenido: '#EF5350',
    en_derivacion: '#CE93D8',
    solicitud_info: '#64B5F6',
    archivado: '#8D6E63'
};

function lsMetricCell(value, maxVal, color) {
    const v = Number(value) || 0;
    const max = Math.max(Number(maxVal) || 0, 1);
    const pct = Math.min(100, Math.round((v / max) * 100));
    return `
        <td class="ls-metric-cell">
            <div class="ls-metric-val">${v.toLocaleString()}</div>
            <div class="ls-bar-track"><div class="ls-bar-fill" style="width:${pct}%;background:${color}"></div></div>
        </td>`;
}

function newEstadoBucket() {
    return { total: 0, en_proceso: 0, finalizado: 0, detenido: 0, en_derivacion: 0, solicitud_info: 0, archivado: 0 };
}

function addEstadoToBucket(bucket, estado) {
    bucket.total++;
    if (bucket[estado] !== undefined) bucket[estado]++;
    else bucket.en_proceso++;
}

function destroyRankingCharts() {
    if (rankingTipoHBarInst) {
        try { rankingTipoHBarInst.destroy(); } catch (e) { /* ignore */ }
        rankingTipoHBarInst = null;
    }
    if (rankingRespHBarInst) {
        try { rankingRespHBarInst.destroy(); } catch (e) { /* ignore */ }
        rankingRespHBarInst = null;
    }
    if (window.myDonutResp) {
        try { window.myDonutResp.destroy(); } catch (e) { /* ignore */ }
        window.myDonutResp = null;
    }
}

function lsHBarOptions() {
    return {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: 'y',
        animation: { duration: 400 },
        plugins: {
            legend: { display: false },
            tooltip: {
                backgroundColor: 'rgba(255,255,255,0.97)',
                titleColor: '#202124',
                bodyColor: '#5F6368',
                borderColor: '#DADCE0',
                borderWidth: 1,
                padding: 10,
                cornerRadius: 4
            }
        },
        scales: {
            x: {
                beginAtZero: true,
                grid: { color: '#E0E0E0', drawBorder: false },
                ticks: { color: '#757575', font: { size: 11, family: "'Roboto', sans-serif" } },
                border: { display: false }
            },
            y: {
                grid: { display: false },
                ticks: { color: '#424242', font: { size: 11, family: "'Roboto', sans-serif" } },
                border: { display: false }
            }
        }
    };
}

function applyBiChartDefaults() {
    if (typeof Chart === 'undefined') return;
    try {
        Chart.defaults.font.family = "'Inter', 'Segoe UI', system-ui, -apple-system, sans-serif";
        Chart.defaults.font.size = 12;
        Chart.defaults.color = '#5F6368';
        const legend = Chart.defaults.plugins && Chart.defaults.plugins.legend;
        if (legend && legend.labels) {
            legend.labels.usePointStyle = true;
            legend.labels.pointStyle = 'rect';
            legend.labels.boxWidth = 8;
            legend.labels.padding = 14;
        }
        const tip = Chart.defaults.plugins && Chart.defaults.plugins.tooltip;
        if (tip) {
            tip.backgroundColor = 'rgba(255, 255, 255, 0.98)';
            tip.titleColor = '#202124';
            tip.bodyColor = '#5F6368';
            tip.borderColor = '#DADCE0';
            tip.borderWidth = 1;
            tip.padding = 10;
            tip.cornerRadius = 6;
            tip.titleFont = { weight: '600', size: 12 };
            tip.bodyFont = { size: 12 };
            tip.displayColors = true;
            tip.boxPadding = 4;
        }
        const scales = Chart.defaults.scales;
        if (scales && scales.linear) {
            scales.linear.grid = scales.linear.grid || {};
            scales.linear.grid.color = '#ECEFF1';
            scales.linear.ticks = scales.linear.ticks || {};
            scales.linear.ticks.color = '#5F6368';
            scales.linear.border = { display: false };
        }
        if (scales && scales.category) {
            scales.category.grid = scales.category.grid || {};
            scales.category.grid.color = '#ECEFF1';
            scales.category.ticks = scales.category.ticks || {};
            scales.category.ticks.color = '#5F6368';
            scales.category.border = { display: false };
        }
    } catch (e) {
        console.warn('No se aplicó el tema de gráficos (Chart.js):', e);
    }
}

const chartCenterTextPlugin = {
    id: 'chartCenterText',
    afterDraw(chart) {
        const plugins = chart.options.plugins;
        const opts = plugins && plugins.chartCenterText;
        if (!opts || !opts.display || chart.config.type !== 'doughnut') return;
        const { ctx, chartArea } = chart;
        if (!chartArea) return;
        const cx = (chartArea.left + chartArea.right) / 2;
        const cy = (chartArea.top + chartArea.bottom) / 2;
        ctx.save();
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillStyle = opts.color || '#1A3C6E';
        ctx.font = `600 ${opts.fontSize || 22}px 'Inter', 'Segoe UI', sans-serif`;
        ctx.fillText(String(opts.text != null ? opts.text : ''), cx, cy - 8);
        ctx.fillStyle = '#80868B';
        ctx.font = `400 11px 'Inter', 'Segoe UI', sans-serif`;
        ctx.fillText(opts.subtext || '', cx, cy + 10);
        ctx.restore();
    }
};

if (typeof Chart !== 'undefined') {
    try {
        Chart.register(chartCenterTextPlugin);
        applyBiChartDefaults();
    } catch (e) {
        console.warn('Chart.js: registro de plugin o defaults falló:', e);
    }
}

/* ==========================================
   GEOGRAPHICAL MAP
   ========================================== */
const CANTON_PARROQUIA_MAP = {
    "QUITO": ["BELISARIO QUEVEDO", "CARCELÉN", "CENTRO HISTÓRICO", "COCHAPAMBA", "COMITÉ DEL PUEBLO", "COTOCOLLAO", "CHILIBULO", "CHILLOGALLO", "CHIMBACALLE", "EL CONDADO", "GUAMANÍ", "IÑAQUITO", "ITCHIMBIA", "JIPIJAPA", "KENNEDY", "LA ARGELIA", "LA CONCEPCIÓN", "LA ECUATORIANA", "LA FERROVIARIA", "LA LIBERTAD", "LA MAGDALENA", "LA MENA", "MARISCAL SUCRE", "PONCEANO", "PUENGASÍ", "QUITUMBE", "RUMIPAMBA", "SAN BARTOLO", "SAN ISIDRO DEL INCA", "SAN JUAN", "SOLANDA", "TURUBAMBA", "ALANGASÍ", "AMAGUAÑA", "ATAHUALPA", "CALACALÍ", "CALDERÓN", "CONOCOTO", "CUMBAYÁ", "CHAVEZPAMBA", "CHECA", "EL QUINCHE", "GUALEA", "GUANGOPOLO", "GUAYLLABAMBA", "LA MERCED", "LLANO CHICO", "LLOA", "NANEGAL", "NANEGALITO", "NAYÓN", "NONO", "PACTO", "PERUCHO", "PIFO", "PÍNTAG", "POMASQUI", "PUÉLLARO", "PUEMBO", "SAN ANTONIO", "SAN JOSÉ DE MINAS", "TABABELA", "TUMBACO", "YARUQUÍ", "ZÁMBIZA"],
    "CAYAMBE": ["CAYAMBE", "JUAN MONTALVO", "ASCÁZUBI", "CANGAHUA", "OLMEDO", "OTÓN", "SANTA ROSA DE CUZUBAMBA", "SAN JOSÉ DE AYORA"],
    "MEJIA": ["MACHACHI", "ALOAG", "ALOASÍ", "CUTUGLAHUA", "EL CHAUPI", "MANUEL CORNEJO ASTORGA", "TAMBILLO", "UYUMBICHO"],
    "PEDRO MONCAYO": ["TABACUNDO", "LA ESPERANZA", "MALCHINGUÍ", "TOCACHI", "TUPIGACHI"],
    "RUMIÑAHUI": ["SANGOLQUÍ", "SAN PEDRO DE TABOADA", "SAN RAFAEL", "FAJARDO", "COTOGCHOA", "RUMIPAMBA"],
    "SAN MIGUEL DE LOS BANCOS": ["SAN MIGUEL DE LOS BANCOS", "MINDO"],
    "PEDRO VICENTE MALDONADO": ["PEDRO VICENTE MALDONADO"],
    "PUERTO QUITO": ["PUERTO QUITO"]
};

/* ==========================================
   HELPERS
   ========================================== */
function extractTipo(faseCode) {
    if (!faseCode) return '';
    return faseCode.replace(/-\d+.*$/, '');
}

function getEstadoTramite(item, tiempos) {
    // ESTADO_ACTUAL (sincronizado desde AppSheet via Apps Script) tiene prioridad;
    // si aún está vacío, se cae al campo Estado original como respaldo.
    const estadoRaw = (item.ESTADO_ACTUAL || item.Estado || '').toLowerCase().trim();
    if (estadoRaw.includes('archivado') || item.Archivado) return 'archivado';
    if (estadoRaw.includes('derivaci') || estadoRaw.includes('derivad')) return 'en_derivacion';
    if (estadoRaw.includes('detenid')) return 'detenido';
    if (estadoRaw.includes('solicitud') || estadoRaw.includes('informaci')) return 'solicitud_info';
    if (estadoRaw.includes('finaliz') || estadoRaw.includes('complet')) return 'finalizado';
    if (estadoRaw.includes('proceso') || estadoRaw.includes('progreso')) return 'en_proceso';

    // Check tiempos records
    const tramiteTiempos = tiempos.filter(t => t.id_tramite === item.id);
    if (tramiteTiempos.length === 0) return 'en_proceso';

    const allCompleted = tramiteTiempos.every(t => t.fecha_hora_fin && t.fecha_hora_fin !== '');
    if (allCompleted) return 'finalizado';
    return 'en_proceso';
}

/** Desglose para vistas ejecutivas */
function getEstadoBreakdownCounts(items, tiempos) {
    const b = newEstadoBucket();
    items.forEach(item => {
        addEstadoToBucket(b, getEstadoTramite(item, tiempos));
    });
    return b;
}

const EXEC_STATE_SEGMENTS = [
    { key: 'en_proceso', label: 'En Proceso', color: '#E88B00' },
    { key: 'finalizado', label: 'Finalizado', color: '#007B4F' },
    { key: 'detenido', label: 'Detenido', color: '#A52422' },
    { key: 'en_derivacion', label: 'En Derivación', color: '#7B1FA2' },
    { key: 'solicitud_info', label: 'Solicitud Info.', color: '#1565C0' },
    { key: 'archivado', label: 'Archivado', color: '#6B3A2A' }
];

/** Correo normalizado desde fila de hoja Usuarios (columnas pueden variar). */
function usuarioCorreoNorm(u) {
    const v = rowPick(u, ['Email (Ejemplo)', 'Email', 'email', 'Correo', 'CORREO', 'Gmail', 'USUARIO', 'Usuario', 'usuario']);
    return String(v || '').trim().toLowerCase();
}

function usuarioNombreDeFila(u) {
    const v = rowPick(u, ['Nombre', 'NOMBRE', 'Name', 'name']);
    return String(v || '').trim();
}

function getUserName(email) {
    if (!dashboardData || !email) return email || 'Sin asignar';
    const key = String(email).trim().toLowerCase();
    if (!key) return 'Sin asignar';
    const usuarios = dashboardData.usuarios || [];
    const user = usuarios.find(u => usuarioCorreoNorm(u) === key);
    if (user) {
        const nom = usuarioNombreDeFila(user);
        return nom || email;
    }
    return email;
}

function getUserNameShort(email) {
    const name = getUserName(email);
    if (name === email) return email;
    // Return first name + last name
    const parts = name.split(' ');
    if (parts.length >= 2) return parts[0] + ' ' + parts[1];
    return name;
}

/**
 * Etiqueta para ejes de gráficos cuando no hay nombre en catálogo: evita correo completo si el dato no coincide (ej. typo vs hoja Usuarios).
 */
function getResponsibleAxisLabel(rawEmail) {
    const resolved = getUserName(rawEmail);
    if (resolved && resolved !== rawEmail) return getUserNameShort(rawEmail);
    const local = String(rawEmail || '').split('@')[0].trim();
    return local || rawEmail || 'Sin asignar';
}

function parseDateOnly(dateString) {
    if (!dateString) return null;
    const date = new Date(dateString);
    if (isNaN(date.getTime())) return null;
    // Strip time for pure date comparisons
    date.setHours(0, 0, 0, 0);
    return date;
}

/** Fila de Fases_Tramite que coincide con el código de fase (única fuente de verdad para nombre y tipo). */
function getFaseRowFromCatalog(fasesList, faseCode) {
    if (!faseCode || !fasesList || !fasesList.length) return null;
    const searches = [String(faseCode).trim().toUpperCase(), String(faseCode).trim()];
    return fasesList.find(f => {
        const fCode = (f.ID_FASE_TRAMITE || f.CODIGO || f.Codigo || '').toString().trim();
        return searches.includes(fCode.toUpperCase());
    }) || null;
}

function getFaseName(code) {
    if (!dashboardData || !code) return code || 'Sin fase';
    if (!dashboardData.fases) return code;
    const fase = getFaseRowFromCatalog(dashboardData.fases, code);
    return fase ? (fase.NOMBRE_FASE || fase.Nombre_Fase || code) : code;
}

/** Tipo de trámite (TV, CEV, …) según catálogo Fases_Tramite; si falta, por prefijo del código. */
function tipoTramiteFromFaseRow(faseRow, faseCodeFallback) {
    if (faseRow) {
        const t = String(faseRow.Tipo || faseRow.TT || faseRow.TIPO || faseRow.Tipo_Tramite || '').trim();
        if (t) return t;
        const cod = (faseRow.Codigo || faseRow.CODIGO || faseRow.ID_FASE_TRAMITE || '').toString().trim();
        if (cod) return extractTipo(cod);
    }
    return extractTipo(faseCodeFallback || '');
}

/** Lee un campo de una fila de sheet con distintos nombres de columna posibles */
function rowPick(row, candidates) {
    if (!row || typeof row !== 'object') return '';
    const keys = Object.keys(row);
    for (const name of candidates) {
        const want = name.trim().toLowerCase().replace(/\s+/g, '_');
        const hit = keys.find(k => k.trim().toLowerCase().replace(/\s+/g, '_') === want);
        if (hit) {
            const v = row[hit];
            if (v !== undefined && v !== null && String(v).trim() !== '') return v;
        }
    }
    return '';
}

/** Índice id/código → etiqueta para tablas Der_Cat1 / Der_Cat2 */
function registerDerCatalogKeys(map, idRaw, label) {
    const lbl = String(label || '').trim();
    if (!lbl) return;
    if (idRaw === '' || idRaw === null || idRaw === undefined) return;
    const s = String(idRaw).trim();
    if (!s) return;
    map.set(s, lbl);
    map.set(s.toLowerCase(), lbl);
    const n = Number(s);
    if (!isNaN(n) && Number.isFinite(n)) map.set(String(n), lbl);
}

function buildDerCatalogMap(rows) {
    const map = new Map();
    (rows || []).forEach(row => {
        // Der_Cat1: ID_Nivel1 / Categoría Principal
        // Der_Cat2: ID_Nivel2 / Entidad Secundaria
        const id = rowPick(row, [
            'ID_Nivel1', 'ID_Nivel2',
            'ID', 'Id', 'id', 'CODIGO', 'Codigo', 'Codigo_Cat'
        ]);
        const name = rowPick(row, [
            'Categoría Principal', 'Categoria Principal',
            'Entidad Secundaria', 'Entidad Principal',
            'Nombre', 'NOMBRE', 'Descripcion', 'DESCRIPCION'
        ]);
        if (id === '' || id === null || id === undefined) return;
        const label = String(name || '').trim() || String(id);
        registerDerCatalogKeys(map, id, label);
    });
    return map;
}

function lookupDerCatalog(map, rawVal) {
    if (rawVal == null || rawVal === '') return '';
    const s = String(rawVal).trim();
    if (!s) return '';
    if (map.has(s)) return map.get(s);
    if (map.has(s.toLowerCase())) return map.get(s.toLowerCase());
    const n = Number(s);
    if (!isNaN(n) && Number.isFinite(n) && map.has(String(n))) return map.get(String(n));
    return '';
}

/** Etiqueta legible para agrupar derivados usando Der_Cat1 + Der_Cat2 del trámite y catálogos. */
function institucionDerivacionDesdeIntake(item, map1, map2) {
    const v1 = rowPick(item, ['Der_Cat1', 'DER_CAT1', 'der_cat1', 'Der_cat1']);
    const v2 = rowPick(item, ['Der_Cat2', 'DER_CAT2', 'der_cat2', 'Der_cat2']);
    const n1 = lookupDerCatalog(map1, v1);
    const n2 = lookupDerCatalog(map2, v2);
    const parts = [];
    if (n1) parts.push(n1);
    else if (v1 !== '' && v1 != null) parts.push(String(v1).trim());
    if (n2) parts.push(n2);
    else if (v2 !== '' && v2 != null) parts.push(String(v2).trim());
    if (parts.length) return parts.join(' · ');
    return 'Sin institución especificada';
}

function escapeHtml(str) {
    if (str == null) return '';
    const d = document.createElement('div');
    d.textContent = String(str);
    return d.innerHTML;
}

/**
 * Clave de equipo (hoja Equipos) -> Set de códigos de fase.
 * Modelo BD1: Id_Equipos = código de fase (ej. TV-01); Usuario_equipo.ID_Equipo apunta a ese mismo código.
 * Modelo legacy: ID_Equipo numérico + columna aparte con Codigo / Fase / ID_FASE_TRAMITE.
 */
function buildFasesPorEquipoKey(equipos) {
    const map = new Map();
    (equipos || []).forEach(eq => {
        const idEquipos = String(rowPick(eq, ['Id_Equipos', 'ID_EQUIPOS', 'id_equipos']) || '').trim();
        const idEquipoAlt = String(rowPick(eq, ['ID_Equipo', 'Id_Equipo', 'id_equipo', 'IdEquipo', 'IDEQUIPO']) || '').trim();
        const faseExplicit = String(rowPick(eq, [
            'Codigo_Fase', 'CODIGO_FASE', 'Fase_del_Tramite', 'Fase_Tramite', 'Fase',
            'ID_FASE_TRAMITE', 'Id_Fase_Tramite', 'Codigo', 'CODIGO', 'Código'
        ]) || '').trim();

        const keys = [...new Set([idEquipos, idEquipoAlt].filter(Boolean))];
        if (keys.length === 0) return;

        const fasesForRow = new Set();
        if (faseExplicit) fasesForRow.add(faseExplicit);
        keys.forEach(k => fasesForRow.add(k));

        keys.forEach((k) => {
            if (!map.has(k)) map.set(k, new Set());
            fasesForRow.forEach((f) => map.get(k).add(f));
        });

        keys.forEach((k) => {
            const n = Number(k);
            if (!Number.isNaN(n) && String(n) !== k) {
                const nk = String(n);
                if (!map.has(nk)) map.set(nk, new Set());
                fasesForRow.forEach((f) => map.get(nk).add(f));
            }
        });
    });
    return map;
}

/**
 * email (minúsculas) -> Set de fases, usando Usuario_equipo + Equipos y tabla usuarios por ID.
 */
function buildEmailToFasesAsignadas(dashboardData) {
    const { equipos, usuarioEquipo, usuarios } = dashboardData;
    const fasesPorClaveEquipo = buildFasesPorEquipoKey(equipos);
    const uidToEmail = new Map();
    (usuarios || []).forEach(u => {
        const id = u.ID_Usuario ?? u.Id_Usuario ?? u.id;
        const em = (u['Email (Ejemplo)'] || u.Email || u.email || '').toLowerCase().trim();
        if (id == null || !em) return;
        uidToEmail.set(String(id), em);
        uidToEmail.set(String(Number(id)), em);
    });

    const emailToFases = new Map();
    (usuarioEquipo || []).forEach(ue => {
        let email = String(rowPick(ue, [
            'Email', 'email', 'Correo', 'CORREO', 'Usuario_Email', 'USUARIO',
            'ID_Usuario (gmail)', 'ID_Usuario_(gmail)'
        ]) || '').toLowerCase().trim();
        if (!email) {
            const uid = rowPick(ue, ['ID_Usuario', 'Id_Usuario', 'id_usuario', 'IdUsuario']);
            if (uid !== '' && uid != null) {
                email = uidToEmail.get(String(uid)) || uidToEmail.get(String(Number(uid))) || '';
            }
        }
        const eidRaw = rowPick(ue, ['ID_Equipo', 'Id_Equipo', 'id_equipo', 'IdEquipo', 'Id_Equipos']);
        const eid = eidRaw === '' || eidRaw == null ? '' : String(eidRaw).trim();
        if (!email || !eid) return;
        let fases = fasesPorClaveEquipo.get(eid) || fasesPorClaveEquipo.get(String(Number(eid)));
        if (!fases || fases.size === 0) {
            fases = new Set([eid]);
        }
        if (!emailToFases.has(email)) emailToFases.set(email, new Set());
        fases.forEach((f) => emailToFases.get(email).add(f));
    });
    return emailToFases;
}

function sortFaseCodesByOrden(codes, fasesList) {
    const orden = new Map();
    (fasesList || []).forEach(f => {
        const c = (f.ID_FASE_TRAMITE || f.CODIGO || f.Codigo || '').toString().trim();
        if (!c) return;
        orden.set(c.toUpperCase(), parseFloat(f.ORDEN || f.Orden) || 0);
    });
    return [...codes].sort((a, b) => {
        const oa = orden.get(a.toUpperCase()) ?? 999;
        const ob = orden.get(b.toUpperCase()) ?? 999;
        if (oa !== ob) return oa - ob;
        return String(a).localeCompare(String(b), 'es');
    });
}

function getFilteredIntake() {
    if (!dashboardData) return [];
    let data = dashboardData.intake;

    // Filtros Antiguos
    if (filtroTipo !== 'Todos') {
        data = data.filter(i => extractTipo(i.Fase_del_Tramite) === filtroTipo);
    }
    if (filtroFase !== 'Todos') {
        data = data.filter(i => i.Fase_del_Tramite === filtroFase);
    }
    if (filtroUsuario !== 'Todos') {
        data = data.filter(i => (i.Responsible || '').toLowerCase() === filtroUsuario.toLowerCase());
    }

    // Filtros Nuevos: Ubicacion
    if (filtroCanton !== 'Todos') {
        data = data.filter(i => (i[' CANTÓN'] || i['CANTÓN'] || i['CANTON'] || '').trim() === filtroCanton);
    }
    if (filtroParroquia !== 'Todos') {
        data = data.filter(i => (i.PARROQUIA || '').trim() === filtroParroquia);
    }

    // Filtros Nuevos: Solicitante (Busqueda de Substring, Insensible a Mayusculas)
    if (filtroNombre.trim() !== '') {
        const search = filtroNombre.toLowerCase().trim();
        data = data.filter(i => (i.Name || '').toLowerCase().includes(search));
    }

    // Filtros Nuevos: Fechas de Inicio
    if (filtroInicioDesde || filtroInicioHasta) {
        const dDesde = parseDateOnly(filtroInicioDesde + "T00:00:00");
        const dHasta = parseDateOnly(filtroInicioHasta + "T00:00:00");

        data = data.filter(i => {
            const itemDate = parseDateOnly(i['Start date']);
            if (!itemDate) return false; // si no tiene fecha, lo ocultamos si hay filtro

            if (dDesde && itemDate < dDesde) return false;
            if (dHasta && itemDate > dHasta) return false;
            return true;
        });
    }

    // Filtros Nuevos: Fechas de Cumplimiento
    if (filtroCumplDesde || filtroCumplHasta) {
        const dDesde = parseDateOnly(filtroCumplDesde + "T00:00:00");
        const dHasta = parseDateOnly(filtroCumplHasta + "T00:00:00");

        data = data.filter(i => {
            const itemDate = parseDateOnly(i['Compliance date']);
            if (!itemDate) return false;

            if (dDesde && itemDate < dDesde) return false;
            if (dHasta && itemDate > dHasta) return false;
            return true;
        });
    }

    return data;
}

/* ==========================================
   INITIALIZATION
   ========================================== */
document.addEventListener('DOMContentLoaded', () => {
    loadData();
});

function normalizeDashboardData(d) {
    if (!d) return;
    ['intake', 'fases', 'tiempos', 'tipos', 'equipos', 'usuarioEquipo', 'derCat1', 'derCat2'].forEach((k) => {
        if (!Array.isArray(d[k])) d[k] = [];
    });
}

async function loadData() {
    showLoading(true);
    try {
        dashboardData = await fetchAllDashboardData();
        console.log('Dashboard data loaded:', dashboardData);
    } catch (error) {
        console.error('Fallo al obtener datos:', error);
        dashboardData = typeof getFallbackData === 'function' ? getFallbackData() : { intake: [], fases: [], tiempos: [], tipos: [], equipos: [], usuarioEquipo: [], usuarios: [], derCat1: [], derCat2: [] };
    }
    normalizeDashboardData(dashboardData);
    try {
        buildFilters();
        renderAllPages();
    } catch (err) {
        console.error('Error al pintar el dashboard:', err);
    } finally {
        showLoading(false);
    }
}

function showLoading(show) {
    const el = document.getElementById('loading');
    if (!el) return;
    if (show) el.classList.remove('hidden');
    else el.classList.add('hidden');
}

/* ==========================================
   FILTERS
   ========================================== */
function buildFilters() {
    if (!dashboardData) return;

    const fasesUnicas = [...new Set(dashboardData.intake.map(i => i.Fase_del_Tramite).filter(Boolean))].sort();
    const tiposUnicos = [...new Set(dashboardData.intake.map(i => extractTipo(i.Fase_del_Tramite)).filter(Boolean))].sort();
    const cantonesUnicos = [...new Set(dashboardData.intake.map(i => (i[' CANTÓN'] || i['CANTÓN'] || i['CANTON'] || '').trim()).filter(Boolean))].sort();
    const parroquiasUnicas = [...new Set(dashboardData.intake.map(i => (i.PARROQUIA || '').trim()).filter(Boolean))].sort();

    // Populate Tipo filter
    const tipoSelect = document.getElementById('filtro-tipo');
    if (tipoSelect) {
        tipoSelect.innerHTML = '<option value="Todos">Todos los Tipos</option>';
        tiposUnicos.forEach(t => {
            tipoSelect.innerHTML += `<option value="${t}">${TYPE_NAMES[t] || t}</option>`;
        });
        tipoSelect.value = filtroTipo;
    }

    // Populate Fase filter
    const faseSelect = document.getElementById('filtro-fase');
    const mobileFaseSelect = document.getElementById('filtro-fase-mobile');
    
    // Filtrar fases según el tipo seleccionado para que sea más práctico
    let fasesFiltradas = fasesUnicas;
    if (filtroTipo !== 'Todos') {
        fasesFiltradas = fasesUnicas.filter(f => extractTipo(f) === filtroTipo);
    }

    if (faseSelect) {
        faseSelect.innerHTML = '<option value="Todos">Todas las Fases</option>';
        fasesFiltradas.forEach(f => {
            const label = getFaseName(f);
            faseSelect.innerHTML += `<option value="${f}">${label}</option>`;
        });
        // Si la fase seleccionada ya no está en la lista filtrada, volver a 'Todos'
        if (filtroFase !== 'Todos' && !fasesFiltradas.includes(filtroFase)) {
            filtroFase = 'Todos';
            faseSelect.value = 'Todos';
        } else {
            faseSelect.value = filtroFase;
        }
    }

    if (mobileFaseSelect) {
        mobileFaseSelect.innerHTML = '<option value="Todos">Todas las Fases</option>';
        fasesFiltradas.forEach(f => {
            mobileFaseSelect.innerHTML += `<option value="${f}">${getFaseName(f)}</option>`;
        });
        mobileFaseSelect.value = filtroFase;
    }

    // Populate Usuario filter
    const userSelect = document.getElementById('filtro-usuario');
    const emailsEnUso = [...new Set(dashboardData.intake.map(i => i.Responsible).filter(Boolean))].sort();
    if (userSelect) {
        userSelect.innerHTML = '<option value="Todos">Todos los Usuarios</option>';
        emailsEnUso.forEach(email => {
            const name = getUserNameShort(email);
            userSelect.innerHTML += `<option value="${email}">${name}</option>`;
        });
        userSelect.value = filtroUsuario;
    }

    // Populate Cantón filter
    const cantonSelect = document.getElementById('filtro-canton');
    if (cantonSelect) {
        cantonSelect.innerHTML = '<option value="Todos">Todos</option>';
        cantonesUnicos.forEach(c => {
            cantonSelect.innerHTML += `<option value="${c}">${c}</option>`;
        });
        cantonSelect.value = filtroCanton;
    }

    // Filtrar parroquias según el cantón seleccionado
    let parroquiasFiltradas = parroquiasUnicas;
    if (filtroCanton !== 'Todos') {
        const permitidas = CANTON_PARROQUIA_MAP[filtroCanton] || [];
        parroquiasFiltradas = parroquiasUnicas.filter(p => permitidas.includes(p));
    }

    // Populate Parroquia filter
    const parroquiaSelect = document.getElementById('filtro-parroquia');
    if (parroquiaSelect) {
        parroquiaSelect.innerHTML = '<option value="Todos">Todas</option>';
        parroquiasFiltradas.forEach(p => {
            parroquiaSelect.innerHTML += `<option value="${p}">${p}</option>`;
        });
        
        // Si la parroquia seleccionada ya no esta en la lista filtrada, volver a 'Todos'
        if (filtroParroquia !== 'Todos' && !parroquiasFiltradas.includes(filtroParroquia)) {
            filtroParroquia = 'Todos';
            parroquiaSelect.value = 'Todos';
        } else {
            parroquiaSelect.value = filtroParroquia;
        }
    }

    // Sync mobile filter dropdowns
    syncMobileFilters(tiposUnicos, fasesUnicas, emailsEnUso, cantonesUnicos, parroquiasFiltradas);
}

function syncMobileFilters(tiposUnicos, fasesUnicas, emailsEnUso, cantonesUnicos, parroquiasFiltradas) {
    // Mobile Tipo
    const tipoMobile = document.getElementById('filtro-tipo-mobile');
    if (tipoMobile) {
        tipoMobile.innerHTML = '<option value="Todos">Todos los Tipos</option>';
        tiposUnicos.forEach(t => {
            tipoMobile.innerHTML += `<option value="${t}">${TYPE_NAMES[t] || t}</option>`;
        });
        tipoMobile.value = filtroTipo;
    }

    // Mobile Fase
    const faseMobile = document.getElementById('filtro-fase-mobile');
    if (faseMobile) {
        faseMobile.innerHTML = '<option value="Todos">Todas las Fases</option>';
        fasesUnicas.forEach(f => {
            faseMobile.innerHTML += `<option value="${f}">${getFaseName(f)}</option>`;
        });
        faseMobile.value = filtroFase;
    }

    // Mobile Usuario
    const userMobile = document.getElementById('filtro-usuario-mobile');
    if (userMobile) {
        userMobile.innerHTML = '<option value="Todos">Todos los Usuarios</option>';
        emailsEnUso.forEach(email => {
            const name = getUserNameShort(email);
            userMobile.innerHTML += `<option value="${email}">${name}</option>`;
        });
        userMobile.value = filtroUsuario;
    }

    // Mobile Cantón
    const cantonMobile = document.getElementById('filtro-canton-mobile');
    if (cantonMobile) {
        cantonMobile.innerHTML = '<option value="Todos">Todos</option>';
        cantonesUnicos.forEach(c => {
            cantonMobile.innerHTML += `<option value="${c}">${c}</option>`;
        });
        cantonMobile.value = filtroCanton;
    }

    // Mobile Parroquia
    const parroquiaMobile = document.getElementById('filtro-parroquia-mobile');
    if (parroquiaMobile) {
        parroquiaMobile.innerHTML = '<option value="Todos">Todas</option>';
        parroquiasFiltradas.forEach(p => {
            parroquiaMobile.innerHTML += `<option value="${p}">${p}</option>`;
        });
        parroquiaMobile.value = filtroParroquia;
    }
}

function onFiltroTipoChange(value) {
    filtroTipo = value;
    filtroFase = 'Todos';
    // Sync both desktop & mobile
    const faseSelect = document.getElementById('filtro-fase');
    const faseMobile = document.getElementById('filtro-fase-mobile');
    const tipoSelect = document.getElementById('filtro-tipo');
    const tipoMobile = document.getElementById('filtro-tipo-mobile');
    if (faseSelect) faseSelect.value = 'Todos';
    if (faseMobile) faseMobile.value = 'Todos';
    if (tipoSelect) tipoSelect.value = value;
    if (tipoMobile) tipoMobile.value = value;
    renderAllPages();
}

function onFiltroFaseChange(value) {
    filtroFase = value;
    // Sync both desktop & mobile
    const faseSelect = document.getElementById('filtro-fase');
    const faseMobile = document.getElementById('filtro-fase-mobile');
    if (faseSelect) faseSelect.value = value;
    if (faseMobile) faseMobile.value = value;
    renderAllPages();
}

function onFiltroUsuarioChange(value) {
    filtroUsuario = value;
    // Sync both desktop & mobile
    const userSelect = document.getElementById('filtro-usuario');
    const userMobile = document.getElementById('filtro-usuario-mobile');
    if (userSelect) userSelect.value = value;
    if (userMobile) userMobile.value = value;
    renderAllPages();
}

/* Nuevos Event Listeners */
function onFiltroInicioDesdeChange(value) {
    filtroInicioDesde = value;
    // Mobile sync is handled inline via `onchange="document.getElementById(...).value = this.value"`
    renderAllPages();
}

function onFiltroInicioHastaChange(value) {
    filtroInicioHasta = value;
    renderAllPages();
}

function onFiltroCumplDesdeChange(value) {
    filtroCumplDesde = value;
    renderAllPages();
}

function onFiltroCumplHastaChange(value) {
    filtroCumplHasta = value;
    renderAllPages();
}

function onFiltroCantonChange(value) {
    filtroCanton = value;
    const cantonSelect = document.getElementById('filtro-canton');
    const cantonMobile = document.getElementById('filtro-canton-mobile');
    if (cantonSelect) cantonSelect.value = value;
    if (cantonMobile) cantonMobile.value = value;
    renderAllPages();
}

function onFiltroParroquiaChange(value) {
    filtroParroquia = value;
    const parrSelect = document.getElementById('filtro-parroquia');
    const parrMobile = document.getElementById('filtro-parroquia-mobile');
    if (parrSelect) parrSelect.value = value;
    if (parrMobile) parrMobile.value = value;
    renderAllPages();
}

function onFiltroNombreChange(value) {
    filtroNombre = value;
    const nombreInput = document.getElementById('filtro-nombre');
    const nombreMobile = document.getElementById('filtro-nombre-mobile');
    if (nombreInput && nombreInput.value !== value) nombreInput.value = value;
    if (nombreMobile && nombreMobile.value !== value) nombreMobile.value = value;
    renderAllPages();
}

/* ==========================================
   MOBILE MENU
   ========================================== */
function toggleMobileMenu() {
    const overlay = document.getElementById('mobile-filter-overlay');
    const panel = document.getElementById('mobile-filter-panel');
    if (overlay && panel) {
        overlay.classList.toggle('hidden');
        panel.classList.toggle('hidden');
        // Prevent body scroll when menu is open
        if (!panel.classList.contains('hidden')) {
            document.body.style.overflow = 'hidden';
        } else {
            document.body.style.overflow = '';
        }
    }
}

/* ==========================================
   NAVIGATION
   ========================================== */
function navigateTo(page, detail = null) {
    if (page === 'ranking-responsables') page = 'ranking';
    currentPage = page;
    document.querySelectorAll('.page').forEach(p => p.classList.add('hidden'));
    const targetPage = document.getElementById(`page-${page}`);
    if (targetPage) targetPage.classList.remove('hidden');

    // Update desktop sidebar nav
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.toggle('active', item.dataset.page === page);
    });

    // Update mobile bottom nav
    document.querySelectorAll('.mobile-nav-item').forEach(item => {
        item.classList.toggle('active', item.dataset.page === page);
    });

    // Update browser title
    if (page === 'ranking') document.title = 'Ranking - Dashboard Trazados Viales';
    else if (page === 'productividad') document.title = 'Productividad - Dashboard Trazados Viales';
    else if (page === 'tipo') document.title = 'Gestión de Trámites - Dashboard Trazados Viales';
    else if (page === 'tendencias') document.title = 'Tendencias - Dashboard Trazados Viales';

    // Close mobile menu if open
    const overlay = document.getElementById('mobile-filter-overlay');
    const panel = document.getElementById('mobile-filter-panel');
    if (overlay && !overlay.classList.contains('hidden')) {
        overlay.classList.add('hidden');
        panel.classList.add('hidden');
        document.body.style.overflow = '';
    }

    if (page === 'detalle' && detail) {
        renderDetallePage(detail);
    }

    // Chart.js mide el canvas al crear el gráfico: si la pestaña estaba oculta, queda en 0×0.
    if (page === 'tipo') {
        requestAnimationFrame(() => {
            requestAnimationFrame(() => {
                try {
                    if (tipoEstadoDonutInst) tipoEstadoDonutInst.resize();
                } catch (e) { /* ignore */ }
            });
        });
    }

    // Scroll to top on page change
    const mainContent = document.querySelector('.main-content');
    if (mainContent) mainContent.scrollTop = 0;
}

/* ==========================================
   RENDER ALL PAGES
   ========================================== */
function renderAllPages() {
    if (!dashboardData) return;
    // Re-build dependent filters (like Phases) if they change based on others
    buildFilters();
    renderRankingPage();
    renderProductividadPage();
    renderTipoPage();
    renderTendenciasPage();
}

/* ==========================================
   PAGE 1: RANKING UNIFICADO (tipos + responsables, estilo Looker Studio)
   ========================================== */
function renderRankingPage() {
    if (!dashboardData) return;
    const intake = getFilteredIntake();
    const { tiempos } = dashboardData;

    destroyRankingCharts();

    const dateEl = document.getElementById('date-range');
    if (dateEl) {
        if (intake.length > 0) {
            const dates = intake.map(i => new Date(i['Start date'])).filter(d => !isNaN(d.getTime())).sort((a, b) => a - b);
            if (dates.length > 0) {
                dateEl.textContent = `${formatDate(dates[0])} - ${formatDate(dates[dates.length - 1])}`;
            } else dateEl.textContent = 'Sin fechas';
        } else dateEl.textContent = 'Sin datos';
    }

    const estadoCounts = {};
    intake.forEach(item => {
        const estado = getEstadoTramite(item, tiempos);
        estadoCounts[estado] = (estadoCounts[estado] || 0) + 1;
    });

    const estadoContainer = document.getElementById('estado-badges');
    if (estadoContainer) {
        estadoContainer.innerHTML = '';
        // TOTAL card first
        estadoContainer.innerHTML += `
            <div class="ls-scorecard">
                <span class="ls-scorecard-label">TOTAL</span>
                <span class="ls-scorecard-value" style="color:#1A3C6E">${intake.length.toLocaleString()}</span>
            </div>`;
        Object.entries(estadoCounts).forEach(([estado, count]) => {
            const cfg = ESTADO_CONFIG[estado] || ESTADO_CONFIG.en_proceso;
            estadoContainer.innerHTML += `
                <div class="ls-scorecard">
                    <span class="ls-scorecard-label">${cfg.label}</span>
                    <span class="ls-scorecard-value" style="color:${cfg.color}">${count.toLocaleString()}</span>
                </div>`;
        });
    }

    const tipoBuckets = {};
    Object.keys(TYPE_NAMES).forEach(code => { tipoBuckets[code] = newEstadoBucket(); });
    const respBuckets = {};

    intake.forEach(item => {
        const estado = getEstadoTramite(item, tiempos);
        const tipo = extractTipo(item.Fase_del_Tramite);
        if (tipo && tipoBuckets[tipo]) {
            addEstadoToBucket(tipoBuckets[tipo], estado);
        }
        const resp = item.Responsible || 'Sin asignar';
        if (!respBuckets[resp]) respBuckets[resp] = newEstadoBucket();
        addEstadoToBucket(respBuckets[resp], estado);
    });

    const tipoRows = Object.keys(TYPE_NAMES).map(code => ({
        ambito: 'Tipo',
        code,
        name: TYPE_NAMES[code] || code,
        stats: tipoBuckets[code]
    })).filter(r => r.stats.total > 0).sort((a, b) => b.stats.total - a.stats.total);

    const respRows = Object.entries(respBuckets).map(([email, stats]) => ({
        ambito: 'Responsable',
        code: email === 'Sin asignar' ? '—' : (email.split('@')[0] || email),
        name: getUserName(email),
        stats,
        email
    })).filter(r => r.stats.total > 0).sort((a, b) => b.stats.total - a.stats.total);

    const allRows = [...tipoRows, ...respRows];
    const maxT = Math.max(...allRows.map(r => r.stats.total), 1);
    const maxProc = Math.max(...allRows.map(r => r.stats.en_proceso), 1);
    const maxFin = Math.max(...allRows.map(r => r.stats.finalizado), 1);
    const maxDet = Math.max(...allRows.map(r => r.stats.detenido), 1);
    const maxDer = Math.max(...allRows.map(r => r.stats.en_derivacion), 1);
    const maxSol = Math.max(...allRows.map(r => r.stats.solicitud_info), 1);
    const maxArch = Math.max(...allRows.map(r => r.stats.archivado), 1);

    // ── TABLA 1: Tipos de trámite ──
    const tbodyTipo = document.getElementById('ranking-tipo-table-body');
    if (tbodyTipo) {
        let html = '';
        let idx = 0;
        tipoRows.forEach((row) => {
            idx++;
            html += `<tr class="ls-data-row">
                <td class="ls-col-idx">${idx}</td>
                <td><span class="ls-ambito-tag ls-ambito-tipo">${row.ambito}</span></td>
                <td><span class="ls-type-dot" style="background:${TYPE_COLORS[row.code] || '#888'}"></span>${escapeHtml(row.name)}</td>
                ${lsMetricCell(row.stats.total, maxT, LS_BAR.total)}
                ${lsMetricCell(row.stats.en_proceso, maxProc, LS_BAR.en_proceso)}
                ${lsMetricCell(row.stats.finalizado, maxFin, LS_BAR.finalizado)}
                ${lsMetricCell(row.stats.detenido, maxDet, LS_BAR.detenido)}
                ${lsMetricCell(row.stats.en_derivacion, maxDer, LS_BAR.en_derivacion)}
                ${lsMetricCell(row.stats.solicitud_info, maxSol, LS_BAR.solicitud_info)}
                ${lsMetricCell(row.stats.archivado, maxArch, LS_BAR.archivado)}
            </tr>`;
        });
        tbodyTipo.innerHTML = html || `<tr><td colspan="10" class="ls-empty">Sin datos</td></tr>`;
    }

    // ── TABLA 2: Responsables por tipo ──
    const tbodyResp = document.getElementById('ranking-resp-table-body');
    if (tbodyResp) {
        let html = '';
        let idx = 0;

        // Construir buckets por tipo → email
        const respByTipo = {};
        intake.forEach(item => {
            const tipo = extractTipo(item.Fase_del_Tramite);
            if (!tipo) return;
            const resp = item.Responsible || 'Sin asignar';
            const estado = getEstadoTramite(item, tiempos);
            if (!respByTipo[tipo]) respByTipo[tipo] = {};
            if (!respByTipo[tipo][resp]) respByTipo[tipo][resp] = newEstadoBucket();
            addEstadoToBucket(respByTipo[tipo][resp], estado);
        });

        // Ordenar tipos por total descendente
        const tiposOrdenadosPorTotal = Object.keys(TYPE_NAMES).filter(t => respByTipo[t]);
        tiposOrdenadosPorTotal.sort((a, b) => {
            const sumA = Object.values(respByTipo[a]).reduce((s, r) => s + r.total, 0);
            const sumB = Object.values(respByTipo[b]).reduce((s, r) => s + r.total, 0);
            return sumB - sumA;
        });

        tiposOrdenadosPorTotal.forEach(tipoKey => {
            const tipoLabel = TYPE_NAMES[tipoKey] || tipoKey;
            const tipoColor = TYPE_COLORS[tipoKey] || '#607D8B';
            html += `<tr class="ls-tipo-group-row"><td colspan="10"><span class="ls-type-dot" style="background:${tipoColor}"></span>${escapeHtml(tipoLabel)}</td></tr>`;

            const respsDelTipo = Object.entries(respByTipo[tipoKey])
                .sort((a, b) => b[1].total - a[1].total);

            respsDelTipo.forEach(([resp, stats]) => {
                const name = getUserName(resp);
                const code = resp === 'Sin asignar' ? '—' : (resp.split('@')[0] || resp);
                idx++;
                html += `<tr class="ls-data-row">
                <td class="ls-col-idx">${idx}</td>
                <td><span class="ls-ambito-tag ls-ambito-resp">Responsable</span></td>
                <td title="${escapeHtml(resp)}"><strong>${escapeHtml(name)}</strong><span class="ls-subcode">${escapeHtml(code)}</span></td>
                ${lsMetricCell(stats.total, maxT, LS_BAR.total)}
                ${lsMetricCell(stats.en_proceso, maxProc, LS_BAR.en_proceso)}
                ${lsMetricCell(stats.finalizado, maxFin, LS_BAR.finalizado)}
                ${lsMetricCell(stats.detenido, maxDet, LS_BAR.detenido)}
                ${lsMetricCell(stats.en_derivacion, maxDer, LS_BAR.en_derivacion)}
                ${lsMetricCell(stats.solicitud_info, maxSol, LS_BAR.solicitud_info)}
                ${lsMetricCell(stats.archivado, maxArch, LS_BAR.archivado)}
            </tr>`;
            });
        });

        // Responsables sin tipo
        const sinTipo = {};
        intake.filter(i => !extractTipo(i.Fase_del_Tramite)).forEach(item => {
            const resp = item.Responsible || 'Sin asignar';
            const estado = getEstadoTramite(item, tiempos);
            if (!sinTipo[resp]) sinTipo[resp] = newEstadoBucket();
            addEstadoToBucket(sinTipo[resp], estado);
        });
        const sinTipoEntries = Object.entries(sinTipo).sort((a, b) => b[1].total - a[1].total);
        if (sinTipoEntries.length > 0) {
            html += `<tr class="ls-subsection-row"><td colspan="10">Sin tipo de trámite identificado</td></tr>`;
            sinTipoEntries.forEach(([resp, stats]) => {
                const name = getUserName(resp);
                const code = resp === 'Sin asignar' ? '—' : (resp.split('@')[0] || resp);
                idx++;
                html += `<tr class="ls-data-row ls-row-sinvinculo">
                <td class="ls-col-idx">${idx}</td>
                <td><span class="ls-ambito-tag ls-ambito-resp">Responsable</span></td>
                <td title="${escapeHtml(resp)}"><strong>${escapeHtml(name)}</strong><span class="ls-subcode">${escapeHtml(code)}</span></td>
                ${lsMetricCell(stats.total, maxT, LS_BAR.total)}
                ${lsMetricCell(stats.en_proceso, maxProc, LS_BAR.en_proceso)}
                ${lsMetricCell(stats.finalizado, maxFin, LS_BAR.finalizado)}
                ${lsMetricCell(stats.detenido, maxDet, LS_BAR.detenido)}
                ${lsMetricCell(stats.en_derivacion, maxDer, LS_BAR.en_derivacion)}
                ${lsMetricCell(stats.solicitud_info, maxSol, LS_BAR.solicitud_info)}
                ${lsMetricCell(stats.archivado, maxArch, LS_BAR.archivado)}
            </tr>`;
            });
        }

        tbodyResp.innerHTML = html || `<tr><td colspan="10" class="ls-empty">Sin datos</td></tr>`;
    }

    const footTotal = document.getElementById('ls-foot-total');
    if (footTotal) footTotal.textContent = intake.length.toLocaleString();

    if (typeof Chart !== 'undefined') {
        const canvasTipo = document.getElementById('rankingTipoHBarChart');
        if (canvasTipo) {
            try {
                const labels = tipoRows.map(r => r.name);
                const data = tipoRows.map(r => r.stats.total);
                rankingTipoHBarInst = new Chart(canvasTipo.getContext('2d'), {
                    type: 'bar',
                    data: {
                        labels: labels.length ? labels : ['—'],
                        datasets: [{
                            label: 'Cantidad',
                            data: data.length ? data : [0],
                            backgroundColor: LS_BAR.total,
                            borderRadius: 2,
                            borderSkipped: false,
                            maxBarThickness: 18
                        }]
                    },
                    options: lsHBarOptions()
                });
            } catch (e) {
                console.error('Ranking chart tipos:', e);
            }
        }

        const canvasResp = document.getElementById('rankingRespHBarChart');
        if (canvasResp) {
            try {
                const top = respRows.slice(0, 14);
                const labels = top.map(r => getUserNameShort(r.email));
                const data = top.map(r => r.stats.total);
                rankingRespHBarInst = new Chart(canvasResp.getContext('2d'), {
                    type: 'bar',
                    data: {
                        labels: labels.length ? labels : ['—'],
                        datasets: [{
                            label: 'Cantidad',
                            data: data.length ? data : [0],
                            backgroundColor: LS_BAR.total,
                            borderRadius: 2,
                            borderSkipped: false,
                            maxBarThickness: 18
                        }]
                    },
                    options: lsHBarOptions()
                });
            } catch (e) {
                console.error('Ranking chart responsables:', e);
            }
        }
    }
}

/* ==========================================
   PAGE 2: PRODUCTIVIDAD POR USUARIO (con tiempos)
   ========================================== */
let prodTipoHBarInst = null;
let prodDuracionChartInst = null;
let prodFaseDuracionChartInst = null;

function renderProductividadPage() {
    if (prodTipoHBarInst) { try { prodTipoHBarInst.destroy(); } catch (e) {} prodTipoHBarInst = null; }
    if (prodDuracionChartInst) { try { prodDuracionChartInst.destroy(); } catch (e) {} prodDuracionChartInst = null; }
    if (prodFaseDuracionChartInst) { try { prodFaseDuracionChartInst.destroy(); } catch (e) {} prodFaseDuracionChartInst = null; }

    const intake = getFilteredIntake();
    const { tiempos } = dashboardData;
    const now = new Date();
    const intakeIds = new Set(intake.map(i => i.id));
    const filteredTiempos = tiempos.filter(t => intakeIds.has(t.id_tramite));

    // ── Días totales por trámite (suma de todas sus fases) ──
    const tramiteDias = {}; // id_tramite → total días acumulados
    filteredTiempos.forEach(t => {
        if (!t.fecha_hora) return;
        const inicio = new Date(t.fecha_hora);
        const fin = t.fecha_hora_fin ? new Date(t.fecha_hora_fin) : now;
        if (isNaN(inicio)) return;
        const dias = Math.max(0, Math.floor((fin - inicio) / 86400000));
        tramiteDias[t.id_tramite] = (tramiteDias[t.id_tramite] || 0) + dias;
    });

    // ── Días DEX por trámite (Fecha_Sol_Oficio → Start date) ──
    const tramiteDexDias = {};
    intake.forEach(item => {
        const start = item['Start date'] ? new Date(item['Start date']) : null;
        const dex   = item['Fecha_Sol_Oficio'] ? new Date(item['Fecha_Sol_Oficio']) : null;
        if (start && dex && !isNaN(start) && !isNaN(dex)) {
            tramiteDexDias[item.id] = Math.max(0, Math.floor((start - dex) / 86400000));
        } else {
            tramiteDexDias[item.id] = 0;
        }
    });

    // ── Agrupar por tipo → responsable con métricas de tiempo + días por estado ──
    const ESTADOS_PROD = ['en_proceso', 'finalizado', 'detenido', 'en_derivacion', 'solicitud_info', 'archivado'];
    const newProdBucket = () => ({
        total: 0, totalDias: 0, dexDias: 0,
        byEstado: Object.fromEntries(ESTADOS_PROD.map(e => [e, { d: 0, c: 0 }]))
    });
    const respByTipo = {};
    intake.forEach(item => {
        const tipo = extractTipo(item.Fase_del_Tramite);
        if (!tipo) return;
        const resp = item.Responsible || 'Sin asignar';
        const estado = getEstadoTramite(item, tiempos);
        if (!respByTipo[tipo]) respByTipo[tipo] = {};
        if (!respByTipo[tipo][resp]) respByTipo[tipo][resp] = newProdBucket();
        const dias = tramiteDias[item.id] || 0;
        const b = respByTipo[tipo][resp];
        b.total++;
        b.totalDias += dias;
        b.dexDias += tramiteDexDias[item.id] || 0;
        if (b.byEstado[estado]) { b.byEstado[estado].d += dias; b.byEstado[estado].c++; }
    });

    // ── KPIs ──
    const totalTramites = intake.length;
    let totalDiasGlobal = 0, countConTiempos = 0;
    Object.entries(tramiteDias).forEach(([id, d]) => {
        if (intakeIds.has(id)) { totalDiasGlobal += d; countConTiempos++; }
    });
    const promDiasTramite = countConTiempos > 0 ? (totalDiasGlobal / countConTiempos).toFixed(1) : '0';
    const allDias = Object.entries(tramiteDias).filter(([id]) => intakeIds.has(id)).map(([, d]) => d);
    const maxDiasGlobal = allDias.length ? Math.max(...allDias) : 0;
    const minDiasGlobal = allDias.filter(d => d > 0).length ? Math.min(...allDias.filter(d => d > 0)) : 0;

    document.getElementById('prod-kpis').innerHTML = `
        <div class="kpi-card">
            <span class="kpi-label">Total Trámites</span>
            <span class="kpi-value" style="color:#1A3C6E">${totalTramites}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">Prom. Días/Trámite</span>
            <span class="kpi-value" style="color:#A52422">${promDiasTramite}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">Mín. Días</span>
            <span class="kpi-value" style="color:#007B4F">${minDiasGlobal}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">Máx. Días</span>
            <span class="kpi-value" style="color:#7B1FA2">${maxDiasGlobal}</span>
        </div>
    `;

    // ── Máximo global de días por estado (para escalar barras) ──
    const maxByEstado = Object.fromEntries(ESTADOS_PROD.map(e => [e, 0]));
    let maxTotal = 0;
    Object.values(respByTipo).forEach(resps => Object.values(resps).forEach(b => {
        maxTotal = Math.max(maxTotal, b.total);
        ESTADOS_PROD.forEach(e => {
            const prom = b.byEstado[e].c > 0 ? Math.round(b.byEstado[e].d / b.byEstado[e].c) : 0;
            maxByEstado[e] = Math.max(maxByEstado[e], prom);
        });
    }));

    // ── Gráfico: Días promedio por tipo — barras apiladas (DEX + fases) ──
    if (typeof Chart !== 'undefined') {
        const fasesData = dashboardData.fases || [];

        // Normalizar cada fila de fases con los nombres reales de columna
        const fasesNorm = fasesData.map(f => ({
            code:   (f.ID_FASE_TRAMITE || f.CODIGO || f.Codigo || '').toString().trim(),
            name:   (f.NOMBRE_FASE || f.Nombre_Fase || ''),
            tipo:   (f.Tipo || f.TIPO || f.TT || f.Tipo_Tramite || '').toString().trim(),
            orden:  parseFloat(f.ORDEN || f.Orden) || 0
        })).filter(f => f.code && f.tipo);

        // Fases por tipo, ordenadas por Orden
        const tipoFasesMap = {}; // tipo → [{ code, name, orden }]
        fasesNorm.forEach(f => {
            if (!tipoFasesMap[f.tipo]) tipoFasesMap[f.tipo] = [];
            tipoFasesMap[f.tipo].push(f);
        });
        Object.keys(tipoFasesMap).forEach(t => {
            tipoFasesMap[t].sort((a, b) => a.orden - b.orden);
        });


        // Días por fase por trámite (usando t.fase que ya está normalizado)
        const tramiteFaseDias = {}; // id_tramite → { faseCode: dias }
        filteredTiempos.forEach(t => {
            if (!t.fecha_hora) return;
            const faseCode = (t.fase || '').toString().trim();
            if (!faseCode) return;
            const ini = new Date(t.fecha_hora);
            const fin = t.fecha_hora_fin ? new Date(t.fecha_hora_fin) : now;
            if (isNaN(ini)) return;
            const dias = Math.max(0, Math.floor((fin - ini) / 86400000));
            if (!tramiteFaseDias[t.id_tramite]) tramiteFaseDias[t.id_tramite] = {};
            tramiteFaseDias[t.id_tramite][faseCode] = (tramiteFaseDias[t.id_tramite][faseCode] || 0) + dias;
        });

        // Acumular por tipo: DEX y por código de fase
        const tipoStats = {}; // tipo → { dex:{sum,count}, fases:{ faseCode:{sum,count} } }
        intake.forEach(item => {
            const tipo = extractTipo(item.Fase_del_Tramite);
            if (!tipo) return;
            if (!tipoStats[tipo]) tipoStats[tipo] = { dex: { sum: 0, count: 0 }, fases: {} };
            tipoStats[tipo].dex.sum += tramiteDexDias[item.id] || 0;
            tipoStats[tipo].dex.count++;
            const fd = tramiteFaseDias[item.id] || {};
            Object.entries(fd).forEach(([code, dias]) => {
                if (!tipoStats[tipo].fases[code]) tipoStats[tipo].fases[code] = { sum: 0, count: 0 };
                tipoStats[tipo].fases[code].sum += dias;
                tipoStats[tipo].fases[code].count++;
            });
        });

        // Lista de tipos con datos, ordenados por total promedio desc
        const tipoList = Object.keys(tipoStats).filter(t => TYPE_NAMES[t]);
        tipoList.sort((a, b) => {
            const avgA = (tipoStats[a].dex.sum / (tipoStats[a].dex.count || 1))
                + Object.values(tipoStats[a].fases).reduce((s, f) => s + f.sum / (f.count || 1), 0);
            const avgB = (tipoStats[b].dex.sum / (tipoStats[b].dex.count || 1))
                + Object.values(tipoStats[b].fases).reduce((s, f) => s + f.sum / (f.count || 1), 0);
            return avgB - avgA;
        });

        // Número máximo de fases entre todos los tipos
        const maxFases = Math.max(...tipoList.map(t => (tipoFasesMap[t] || []).length), 0);

        // Tipo de referencia para nombres de leyenda: el que tenga más fases
        const refTipo = tipoList.reduce((best, t) =>
            (tipoFasesMap[t] || []).length > (tipoFasesMap[best] || []).length ? t : best,
            tipoList[0] || ''
        );
        const refFases = tipoFasesMap[refTipo] || [];

        // Colores por segmento
        const SEG_COLORS = ['#78909C','#1A3C6E','#007B4F','#E88B00','#A52422','#7B1FA2','#1565C0'];

        // Datasets: primero DEX, luego una por posición de fase
        const stackedDatasets = [];

        // Dataset DEX
        stackedDatasets.push({
            label: 'DEX',
            data: tipoList.map(t => {
                const s = tipoStats[t].dex;
                return s.count > 0 ? Math.round(s.sum / s.count) : 0;
            }),
            backgroundColor: SEG_COLORS[0],
            stack: 'total',
            maxBarThickness: 28
        });

        // Datasets por posición de fase — etiqueta = nombre real de la fase de referencia
        for (let i = 0; i < maxFases; i++) {
            const refFase = refFases[i];
            const label = (refFase && refFase.name) ? refFase.name : `Fase ${i + 1}`;
            stackedDatasets.push({
                label,
                data: tipoList.map(t => {
                    const fases = tipoFasesMap[t] || [];
                    if (i >= fases.length) return 0;
                    const code = fases[i].code;
                    const s = tipoStats[t].fases[code];
                    return s && s.count > 0 ? Math.round(s.sum / s.count) : 0;
                }),
                backgroundColor: SEG_COLORS[(i + 1) % SEG_COLORS.length],
                stack: 'total',
                maxBarThickness: 28
            });
        }

        const canvasProd = document.getElementById('prodTipoHBarChart');
        if (canvasProd) {
            // Destruir cualquier instancia anterior — incluye el caso donde prodTipoHBarInst quedó null
            try { Chart.getChart(canvasProd)?.destroy(); } catch (e) {}
            try {
                const opts = lsHBarOptions();
                opts.plugins.legend = {
                    display: true,
                    position: 'top',
                    labels: {
                        usePointStyle: true,
                        pointStyle: 'rect',
                        boxWidth: 10,
                        padding: 12,
                        font: { size: 11 }
                    }
                };
                opts.plugins.subtitle = {
                    display: true,
                    text: `${intake.length} trámite${intake.length !== 1 ? 's' : ''} en el rango seleccionado`,
                    color: '#5F6368',
                    font: { size: 11 },
                    padding: { bottom: 8 }
                };
                opts.plugins.tooltip.callbacks = {
                    title: items => TYPE_NAMES[tipoList[items[0].dataIndex]] || tipoList[items[0].dataIndex],
                    label: c => {
                        const val = c.parsed.x;
                        if (!val) return null;
                        const tipo = tipoList[c.dataIndex];
                        let segName = c.dataset.label;
                        if (c.datasetIndex > 0) {
                            const fasePos = c.datasetIndex - 1;
                            const fases = tipoFasesMap[tipo] || [];
                            if (fases[fasePos]) segName = fases[fasePos].name;
                        }
                        return ` ${segName}: ${val}d`;
                    }
                };
                opts.scales.x.stacked = true;
                opts.scales.x.title = { display: true, text: 'Días promedio por trámite', color: '#5F6368', font: { size: 11 } };
                opts.scales.y.stacked = true;
                prodTipoHBarInst = new Chart(canvasProd.getContext('2d'), {
                    type: 'bar',
                    data: {
                        labels: tipoList.map(t => TYPE_NAMES[t] || t),
                        datasets: stackedDatasets
                    },
                    options: opts
                });
            } catch (e) {
                console.error('Prod tipo chart:', e);
            }
        }

    }

    // ── Tabla agrupada por tipo (estilo Ranking) ──
    // Fases sin fecha_hora_fin se miden hasta hoy (fase aún activa)
    const fmtDias = d => d === 0 ? '< 1d' : `${d}d`;

    const tbody = document.getElementById('prod-table-body');
    let html = '';
    let idx = 0;

    const tiposOrdenados = Object.keys(TYPE_NAMES).filter(t => respByTipo[t]);
    tiposOrdenados.sort((a, b) =>
        Object.values(respByTipo[b]).reduce((s, r) => s + r.total, 0) -
        Object.values(respByTipo[a]).reduce((s, r) => s + r.total, 0)
    );

    // Helper: celda de días por estado con barra de color
    const prodEstadoCells = (byEstado) => ESTADOS_PROD.map(e => {
        const s = byEstado[e] || { d: 0, c: 0 };
        const prom = s.c > 0 ? Math.round(s.d / s.c) : 0;
        const pct = Math.min(100, Math.round((prom / Math.max(maxByEstado[e], 1)) * 100));
        const color = LS_BAR[e] || '#90A4AE';
        return s.c > 0
            ? `<td class="ls-metric-cell"><div class="ls-metric-val">${fmtDias(prom)}</div><div class="ls-bar-track"><div class="ls-bar-fill" style="width:${pct}%;background:${color}"></div></div></td>`
            : `<td class="ls-metric-cell"><div class="ls-metric-val" style="color:#ccc">—</div></td>`;
    }).join('');

    // Máximos globales para escalar las barras de T. Total y T. DEX
    let maxPromTotal = 0, maxPromDex = 0;
    Object.values(respByTipo).forEach(resps => Object.values(resps).forEach(b => {
        const pt = b.total > 0 ? Math.round(b.totalDias / b.total) : 0;
        const pd = b.total > 0 ? Math.round(b.dexDias / b.total) : 0;
        maxPromTotal = Math.max(maxPromTotal, pt);
        maxPromDex   = Math.max(maxPromDex, pd);
    }));

    // Helper: celdas de T. Total y T. DEX con barra
    const prodTiempoCells = (totalDias, dexDias, total) => {
        const promTotal = total > 0 ? Math.round(totalDias / total) : 0;
        const promDex   = total > 0 ? Math.round(dexDias   / total) : 0;
        const pctTotal  = Math.min(100, Math.round((promTotal / Math.max(maxPromTotal, 1)) * 100));
        const pctDex    = Math.min(100, Math.round((promDex   / Math.max(maxPromDex, 1)) * 100));
        const cellTotal = promTotal > 0
            ? `<td class="ls-metric-cell"><div class="ls-metric-val">${fmtDias(promTotal)}</div><div class="ls-bar-track"><div class="ls-bar-fill" style="width:${pctTotal}%;background:#1A3C6E"></div></div></td>`
            : `<td class="ls-metric-cell"><div class="ls-metric-val" style="color:#ccc">—</div></td>`;
        const cellDex = promDex > 0
            ? `<td class="ls-metric-cell"><div class="ls-metric-val">${fmtDias(promDex)}</div><div class="ls-bar-track"><div class="ls-bar-fill" style="width:${pctDex}%;background:#78909C"></div></div></td>`
            : `<td class="ls-metric-cell"><div class="ls-metric-val" style="color:#ccc">—</div></td>`;
        return cellTotal + cellDex;
    };

    // ── Sección 1: Totales por tipo ──
    html += `<tr class="ls-section-row"><td colspan="12">Tipos de trámite <span style="font-weight:400;font-size:11px;opacity:.7">(fases activas medidas hasta hoy · columnas = días prom. por estado)</span></td></tr>`;

    tiposOrdenados.forEach(tipoKey => {
        const tipoColor = TYPE_COLORS[tipoKey] || '#607D8B';
        const tipoByEstado = Object.fromEntries(ESTADOS_PROD.map(e => [e, { d: 0, c: 0 }]));
        let tipoTotal = 0, tipoTotalDias = 0, tipoDexDias = 0;
        Object.values(respByTipo[tipoKey]).forEach(b => {
            tipoTotal     += b.total;
            tipoTotalDias += b.totalDias;
            tipoDexDias   += b.dexDias;
            ESTADOS_PROD.forEach(e => {
                tipoByEstado[e].d += b.byEstado[e].d;
                tipoByEstado[e].c += b.byEstado[e].c;
            });
        });
        idx++;
        html += `<tr class="ls-data-row">
            <td class="ls-col-idx">${idx}</td>
            <td><span class="ls-ambito-tag" style="background:#E3F2FD;color:#1565C0;border-color:#BBDEFB">TIPO</span></td>
            <td><span class="ls-type-dot" style="background:${tipoColor}"></span><strong>${escapeHtml(TYPE_NAMES[tipoKey] || tipoKey)}</strong></td>
            ${lsMetricCell(tipoTotal, maxTotal, LS_BAR.total)}
            ${prodTiempoCells(tipoTotalDias, tipoDexDias, tipoTotal)}
            ${prodEstadoCells(tipoByEstado)}
        </tr>`;
    });

    // ── Sección 2: Responsables por tipo ──
    html += `<tr class="ls-section-row"><td colspan="12">Responsables por tipo de trámite</td></tr>`;

    tiposOrdenados.forEach(tipoKey => {
        const tipoColor = TYPE_COLORS[tipoKey] || '#607D8B';
        html += `<tr class="ls-tipo-group-row"><td colspan="12"><span class="ls-type-dot" style="background:${tipoColor}"></span>${escapeHtml(TYPE_NAMES[tipoKey] || tipoKey)}</td></tr>`;

        Object.entries(respByTipo[tipoKey])
            .sort((a, b) => b[1].total - a[1].total)
            .forEach(([resp, stats]) => {
                const name = getUserName(resp);
                const code = resp === 'Sin asignar' ? '—' : (resp.split('@')[0] || resp);
                idx++;
                html += `<tr class="ls-data-row">
                    <td class="ls-col-idx">${idx}</td>
                    <td><span class="ls-ambito-tag ls-ambito-resp">Responsable</span></td>
                    <td title="${escapeHtml(resp)}"><strong>${escapeHtml(name)}</strong><span class="ls-subcode">${escapeHtml(code)}</span></td>
                    ${lsMetricCell(stats.total, maxTotal, LS_BAR.total)}
                    ${prodTiempoCells(stats.totalDias, stats.dexDias, stats.total)}
                    ${prodEstadoCells(stats.byEstado)}
                </tr>`;
            });
    });

    // Sin tipo identificado
    const sinTipo = {};
    intake.filter(i => !extractTipo(i.Fase_del_Tramite)).forEach(item => {
        const resp = item.Responsible || 'Sin asignar';
        const estado = getEstadoTramite(item, tiempos);
        if (!sinTipo[resp]) sinTipo[resp] = newProdBucket();
        const dias = tramiteDias[item.id] || 0;
        sinTipo[resp].total++;
        sinTipo[resp].totalDias += dias;
        sinTipo[resp].dexDias += tramiteDexDias[item.id] || 0;
        if (sinTipo[resp].byEstado[estado]) { sinTipo[resp].byEstado[estado].d += dias; sinTipo[resp].byEstado[estado].c++; }
    });
    const sinTipoEntries = Object.entries(sinTipo).sort((a, b) => b[1].total - a[1].total);
    if (sinTipoEntries.length > 0) {
        html += `<tr class="ls-subsection-row"><td colspan="12">Sin tipo identificado</td></tr>`;
        sinTipoEntries.forEach(([resp, stats]) => {
            const name = getUserName(resp);
            const code = resp === 'Sin asignar' ? '—' : (resp.split('@')[0] || resp);
            idx++;
            html += `<tr class="ls-data-row ls-row-sinvinculo">
                <td class="ls-col-idx">${idx}</td>
                <td><span class="ls-ambito-tag ls-ambito-resp">Responsable</span></td>
                <td title="${escapeHtml(resp)}"><strong>${escapeHtml(name)}</strong><span class="ls-subcode">${escapeHtml(code)}</span></td>
                ${lsMetricCell(stats.total, maxTotal, LS_BAR.total)}
                ${prodTiempoCells(stats.totalDias, stats.dexDias, stats.total)}
                ${prodEstadoCells(stats.byEstado)}
            </tr>`;
        });
    }

    tbody.innerHTML = html || `<tr><td colspan="12" class="ls-empty">Sin datos de tiempos disponibles</td></tr>`;

    // ── Chart: Duración promedio por responsable (movido desde Tendencias) ──
    (function renderProdDuracionChart() {
        const respDur = {};
        filteredTiempos.forEach(t => {
            if (!t.fecha_hora) return;
            const tramite = intake.find(i => i.id === t.id_tramite);
            if (!tramite) return;
            const raw = (tramite.Responsible || '').trim();
            const key = raw.toLowerCase() || '__sin_asignar__';
            const inicio = new Date(t.fecha_hora);
            const fin = t.fecha_hora_fin ? new Date(t.fecha_hora_fin) : now;
            if (isNaN(inicio)) return;
            const dias = Math.max(0, Math.floor((fin - inicio) / 86400000));
            if (!respDur[key]) respDur[key] = { sum: 0, count: 0, raw: raw || 'Sin asignar' };
            respDur[key].sum += dias;
            respDur[key].count++;
        });

        const sorted = Object.entries(respDur)
            .map(([, { sum, count, raw }]) => ({ name: getResponsibleAxisLabel(raw), avg: Math.round(sum / count) }))
            .sort((a, b) => b.avg - a.avg).slice(0, 12);

        const ctx = document.getElementById('prodDuracionChart');
        if (!ctx || typeof Chart === 'undefined') return;
        if (prodDuracionChartInst) prodDuracionChartInst.destroy();

        const barColors = sorted.map(r => r.avg > 30 ? '#A52422' : r.avg > 14 ? '#E88B00' : '#007B4F');
        prodDuracionChartInst = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: sorted.map(r => r.name),
                datasets: [{ label: 'Días promedio', data: sorted.map(r => r.avg), backgroundColor: barColors, borderRadius: 4, borderSkipped: false, maxBarThickness: 22 }]
            },
            options: {
                indexAxis: 'y', responsive: true, maintainAspectRatio: false, animation: { duration: 400 },
                plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.parsed.x.toLocaleString()} días promedio` } } },
                scales: {
                    x: { beginAtZero: true, grid: { color: '#ECEFF1' }, ticks: { precision: 0 }, title: { display: true, text: 'Días', color: '#5F6368', font: { size: 11, weight: '500' } } },
                    y: { grid: { display: false }, ticks: { font: { size: 11 }, autoSkip: false, maxRotation: 0 } }
                }
            }
        });
    })();

    // ── Chart: Tiempo promedio por fase por técnico ──
    (function renderProdFaseDuracionChart() {
        // Agrupar tiempos por fase
        const faseDur = {};
        filteredTiempos.forEach(t => {
            if (!t.fecha_hora) return;
            const faseCode = (t.fase || '').trim();
            if (!faseCode) return;
            const inicio = new Date(t.fecha_hora);
            const fin = t.fecha_hora_fin ? new Date(t.fecha_hora_fin) : now;
            if (isNaN(inicio)) return;
            const dias = Math.max(0, Math.floor((fin - inicio) / 86400000));
            if (!faseDur[faseCode]) faseDur[faseCode] = { sum: 0, count: 0 };
            faseDur[faseCode].sum += dias;
            faseDur[faseCode].count++;
        });

        const sorted = Object.entries(faseDur)
            .map(([code, { sum, count }]) => ({ code, name: getFaseName(code), avg: Math.round(sum / count) }))
            .sort((a, b) => b.avg - a.avg).slice(0, 12);

        const ctx = document.getElementById('prodFaseDuracionChart');
        if (!ctx || typeof Chart === 'undefined') return;
        if (prodFaseDuracionChartInst) prodFaseDuracionChartInst.destroy();

        const tipo2color = sorted.map(r => TYPE_COLORS[extractTipo(r.code)] || '#607D8B');
        prodFaseDuracionChartInst = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: sorted.map(r => r.name.length > 30 ? r.name.slice(0, 28) + '…' : r.name),
                datasets: [{ label: 'Días promedio', data: sorted.map(r => r.avg), backgroundColor: tipo2color, borderRadius: 4, borderSkipped: false, maxBarThickness: 22 }]
            },
            options: {
                indexAxis: 'y', responsive: true, maintainAspectRatio: false, animation: { duration: 400 },
                plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${c.parsed.x.toLocaleString()} días promedio` } } },
                scales: {
                    x: { beginAtZero: true, grid: { color: '#ECEFF1' }, ticks: { precision: 0 }, title: { display: true, text: 'Días', color: '#5F6368', font: { size: 11, weight: '500' } } },
                    y: { grid: { display: false }, ticks: { font: { size: 11 }, autoSkip: false, maxRotation: 0 } }
                }
            }
        });
    })();

}

/* ==========================================
   PAGE 3: CARTERA — VISTA EJECUTIVA (+ listado colapsable)
   ========================================== */
function renderTipoPage() {
    if (tipoEstadoDonutInst) {
        try { tipoEstadoDonutInst.destroy(); } catch (e) { /* ignore */ }
        tipoEstadoDonutInst = null;
    }

    const intake = getFilteredIntake();
    const { fases, tiempos } = dashboardData;

    const typePhaseCount = {};
    fases.forEach(f => {
        const type = f.TT || f.TIPO || f.Tipo || '';
        typePhaseCount[type] = (typePhaseCount[type] || 0) + 1;
    });

    const typeTramiteCount = {};
    intake.forEach(item => {
        const tipo = extractTipo(item.Fase_del_Tramite);
        typeTramiteCount[tipo] = (typeTramiteCount[tipo] || 0) + 1;
    });

    const totalTypes = Object.keys(TYPE_NAMES).length;
    const totalPhases = fases.length;
    const tiposConCarga = Object.values(typeTramiteCount).filter(v => v > 0).length;
    const brGlobal = getEstadoBreakdownCounts(intake, tiempos);

    document.getElementById('tipos-badge').textContent = `${intake.length.toLocaleString()} trámites`;

    const countEl = document.getElementById('tipo-details-count');
    if (countEl) {
        countEl.textContent = intake.length ? `${intake.length.toLocaleString()} registros` : 'Sin registros';
    }

    const kpiContainer = document.getElementById('tipo-kpis');
    kpiContainer.innerHTML = `
        <div class="kpi-card" style="border-top-color:#1A3C6E">
            <span class="kpi-label">Total trámites</span>
            <span class="kpi-value" style="color:#1A3C6E">${intake.length.toLocaleString()}</span>
        </div>
        ${EXEC_STATE_SEGMENTS.map(s => `
        <div class="kpi-card" style="border-top-color:${s.color}">
            <span class="kpi-label">${s.label}</span>
            <span class="kpi-value" style="color:${s.color}">${(brGlobal[s.key] || 0).toLocaleString()}</span>
        </div>`).join('')}
        <div class="kpi-card" style="border-top-color:#EAB308">
            <span class="kpi-label">Tipos con carga</span>
            <span class="kpi-value" style="color:#B45309">${tiposConCarga} / ${totalTypes}</span>
        </div>
    `;

    // Donut + leyenda — estado de la cartera
    const donutDataVals = EXEC_STATE_SEGMENTS.map(s => brGlobal[s.key]);
    const legendEl = document.getElementById('tipo-donut-legend');
    const donutCanvas = document.getElementById('tipoEstadoDonut');
    if (legendEl) {
        if (intake.length === 0) {
            legendEl.innerHTML = '<p class="tipo-empty-hint">Sin trámites con los filtros actuales.</p>';
        } else {
            const sumLeg = donutDataVals.reduce((a, b) => a + b, 0);
            legendEl.innerHTML = EXEC_STATE_SEGMENTS.map((s, i) => {
                const n = donutDataVals[i];
                const pct = sumLeg ? Math.round((n / sumLeg) * 100) : 0;
                return `
                    <div class="tipo-donut-legend-item">
                        <span class="tipo-donut-swatch" style="background:${s.color}"></span>
                        <span class="tipo-donut-legend-text">${s.label}</span>
                        <span class="tipo-donut-legend-val">${n.toLocaleString()} <em>(${pct}%)</em></span>
                    </div>`;
            }).join('');
        }
    }

    if (typeof Chart !== 'undefined' && donutCanvas && intake.length > 0) {
        const labels = EXEC_STATE_SEGMENTS.map(s => s.label);
        const colors = EXEC_STATE_SEGMENTS.map(s => s.color);
        const filtered = labels.map((l, i) => ({ l, v: donutDataVals[i], c: colors[i] })).filter(x => x.v > 0);
        if (filtered.length > 0) {
            tipoEstadoDonutInst = new Chart(donutCanvas.getContext('2d'), {
                type: 'doughnut',
                data: {
                    labels: filtered.map(x => x.l),
                    datasets: [{
                        data: filtered.map(x => x.v),
                        backgroundColor: filtered.map(x => x.c),
                        borderWidth: 2,
                        borderColor: '#fff',
                        hoverOffset: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '58%',
                    animation: { duration: 450 },
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label(ctx) {
                                    const t = ctx.dataset.data[ctx.dataIndex];
                                    const sum = ctx.dataset.data.reduce((a, b) => a + b, 0);
                                    const pct = sum ? Math.round((t / sum) * 100) : 0;
                                    return ` ${t.toLocaleString()} (${pct}%)`;
                                }
                            }
                        }
                    }
                }
            });
            const pgTipo = document.getElementById('page-tipo');
            if (pgTipo && !pgTipo.classList.contains('hidden')) {
                requestAnimationFrame(() => {
                    try {
                        if (tipoEstadoDonutInst) tipoEstadoDonutInst.resize();
                    } catch (e) { /* ignore */ }
                });
            }
        }
    }

    // Tarjetas por tipo — desglose visual
    const execCards = document.getElementById('tipo-exec-cards');
    if (execCards) {
        let cardsHtml = '';
        Object.entries(TYPE_NAMES).forEach(([code, name]) => {
            const tramites = intake.filter(i => extractTipo(i.Fase_del_Tramite) === code);
            const n = tramites.length;
            const color = TYPE_COLORS[code] || '#888';
            const br = getEstadoBreakdownCounts(tramites, tiempos);
            const muted = n === 0 ? ' exec-type-card--empty' : '';

            let stackHtml = '';
            if (n > 0) {
                stackHtml = '<div class="exec-stack-track" role="img" aria-label="Distribución por estado">';
                EXEC_STATE_SEGMENTS.forEach(s => {
                    const c = br[s.key];
                    const pct = Math.max(0, Math.round((c / n) * 100));
                    if (pct > 0) {
                        stackHtml += `<span class="exec-stack-seg" style="width:${pct}%;background:${s.color}" title="${s.label}: ${c}"></span>`;
                    }
                });
                stackHtml += '</div>';
                stackHtml += '<ul class="exec-type-mini">';
                EXEC_STATE_SEGMENTS.forEach(s => {
                    if (br[s.key] > 0) {
                        stackHtml += `<li><span class="exec-mini-dot" style="background:${s.color}"></span>${s.label}: <strong>${br[s.key]}</strong></li>`;
                    }
                });
                stackHtml += '</ul>';
            } else {
                stackHtml = '<p class="exec-type-empty-msg">Sin expedientes en esta vista.</p>';
            }

            cardsHtml += `
                <div class="exec-type-card${muted}" style="--exec-accent:${color}">
                    <div class="exec-type-card-head">
                        <span class="ls-type-dot" style="background:${color}"></span>
                        <span class="exec-type-title">${escapeHtml(name)}</span>
                        <span class="exec-type-total">${n.toLocaleString()}</span>
                    </div>
                    ${stackHtml}
                </div>`;
        });
        execCards.innerHTML = cardsHtml;
    }

    // Top fases por carga, agrupadas por tipo de trámite
    const faseLoadEl = document.getElementById('tipo-fase-load-bars');
    if (faseLoadEl) {
        const faseCounts = {};
        intake.forEach(i => {
            const f = (i.Fase_del_Tramite || '').trim();
            if (!f) return;
            faseCounts[f] = (faseCounts[f] || 0) + 1;
        });

        if (Object.keys(faseCounts).length === 0) {
            faseLoadEl.innerHTML = '<p class="tipo-empty-hint">No hay datos de fase en la vista filtrada.</p>';
        } else {
            // Agrupar por tipo de trámite
            const byTipo = {};
            Object.entries(faseCounts).forEach(([fcode, cnt]) => {
                const tipo = extractTipo(fcode);
                if (!byTipo[tipo]) byTipo[tipo] = [];
                byTipo[tipo].push({ fcode, cnt });
            });

            const maxF = Math.max(...Object.values(faseCounts), 1);
            let fh = '';
            Object.entries(byTipo).sort((a, b) => {
                const sumA = a[1].reduce((s, x) => s + x.cnt, 0);
                const sumB = b[1].reduce((s, x) => s + x.cnt, 0);
                return sumB - sumA;
            }).forEach(([tipo, fases]) => {
                const tipoName = TYPE_NAMES[tipo] || tipo;
                const tipoColor = TYPE_COLORS[tipo] || '#607D8B';
                fh += `<div class="fase-load-tipo-header" style="border-left:3px solid ${tipoColor};padding-left:8px;margin:12px 0 6px;font-weight:600;font-size:12px;color:${tipoColor}">${escapeHtml(tipoName)}</div>`;
                fases.sort((a, b) => b.cnt - a.cnt).forEach(({ fcode, cnt }) => {
                    const label = getFaseName(fcode);
                    const w = Math.round((cnt / maxF) * 100);
                    fh += `
                        <div class="fase-load-row">
                            <span class="fase-load-label" title="${escapeHtml(fcode)}">${escapeHtml(label.length > 42 ? label.slice(0, 40) + '…' : label)}</span>
                            <div class="fase-load-bar-bg"><div class="fase-load-bar-fill" style="width:${w}%;background:${tipoColor}"></div></div>
                            <span class="fase-load-num">${cnt.toLocaleString()}</span>
                        </div>`;
                });
            });
            faseLoadEl.innerHTML = fh;
        }
    }

    // Barras comparativas por tipo
    const barsContainer = document.getElementById('tipo-bars');
    barsContainer.innerHTML = '';
    const maxTramites = Math.max(...Object.values(typeTramiteCount), 1);

    Object.entries(TYPE_NAMES).forEach(([code, typeLabel]) => {
        const count = typeTramiteCount[code] || 0;
        const widthPct = (count / maxTramites) * 100;
        const color = TYPE_COLORS[code];

        barsContainer.innerHTML += `
            <div class="hbar-item tipo-hbar-exec">
                <span class="hbar-label" style="color:${color}">${typeLabel}</span>
                <div class="hbar-bg">
                    <div class="hbar-fill" style="width:${widthPct}%;background:${color}"></div>
                </div>
                <span class="hbar-value" style="color:${color}">${count.toLocaleString()}</span>
            </div>
        `;
    });

    const summaryContainer = document.getElementById('tipo-summary');
    summaryContainer.innerHTML = `
        <div class="summary-item">
            <span class="summary-value" style="color:#1A3C6E">${intake.length.toLocaleString()}</span>
            <span class="summary-label">Total trámites</span>
        </div>
        <div class="summary-item">
            <span class="summary-value" style="color:#007B4F">${totalPhases}</span>
            <span class="summary-label">Fases catálogo</span>
        </div>
        <div class="summary-item">
            <span class="summary-value" style="color:#EAB308">${tiposConCarga}</span>
            <span class="summary-label">Tipos activos</span>
        </div>
    `;

    // Listado detallado (colapsable)
    const tbody = document.getElementById('tipo-table-body');
    tbody.innerHTML = '';

    Object.entries(TYPE_NAMES).forEach(([code, name]) => {
        const tramites = intake.filter(i => extractTipo(i.Fase_del_Tramite) === code);
        const phaseCount = typePhaseCount[code] || 0;

        if (tramites.length === 0) {
            tbody.innerHTML += `
                <tr>
                    <td><span class="ls-type-dot" style="background:${TYPE_COLORS[code]}"></span>${name}</td>
                    <td style="font-weight:600;color:#1A3C6E">—</td>
                    <td style="text-align:center">${phaseCount}</td>
                    <td><span class="phase-pill" style="background:#F5F5F5;color:#999">-</span></td>
                    <td>-</td>
                    <td><span class="status-pill" style="background:#F5F5F5;color:#999">Sin trámite</span></td>
                </tr>
            `;
        } else {
            tramites.forEach((item) => {
                const faseActual = item.Fase_del_Tramite || '-';
                const estado = getEstadoTramite(item, tiempos);
                const cfg = ESTADO_CONFIG[estado] || ESTADO_CONFIG['en_proceso'];
                const responsable = getUserNameShort(item.Responsible);
                const safeName = escapeHtml((item.Name || name).substring(0, 48));
                const safeTitle = escapeHtml(item.Name || '');

                tbody.innerHTML += `
                    <tr class="clickable-row" onclick="openDetail('${item.id}')">
                        <td><span class="ls-type-dot" style="background:${TYPE_COLORS[code]}"></span>${name}</td>
                        <td style="font-weight:600;color:#1A3C6E" title="${safeTitle}">${safeName}</td>
                        <td style="text-align:center">${phaseCount}</td>
                        <td><span class="phase-pill" style="background:#E3F2FD;color:#1565C0">${escapeHtml(getFaseName(faseActual))}</span></td>
                        <td title="${escapeHtml(item.Responsible || '')}" style="font-size:11px">${escapeHtml(responsable)}</td>
                        <td><span class="status-pill" style="background:${cfg.bg};color:${cfg.color}">${cfg.label}</span></td>
                    </tr>
                `;
            });
        }
    });
}

/* ==========================================
   PAGE 4: DETALLE DE TRAMITE
   ========================================== */
function openDetail(tramiteId) {
    navigateTo('detalle', tramiteId);
}

function renderDetallePage(tramiteId) {
    const { intake, fases, tiempos } = dashboardData;

    const item = intake.find(i => i.id === tramiteId);
    if (!item) return;

    const faseActual = item.Fase_del_Tramite || '';
    const typeCode = extractTipo(faseActual);
    const name = TYPE_NAMES[typeCode] || typeCode;
    const color = TYPE_COLORS[typeCode] || '#888';
    const typeFases = fases.filter(f => (f.TT || f.TIPO || f.Tipo) === typeCode)
        .sort((a, b) => (parseFloat(a.ORDEN || a.Orden) || 0) - (parseFloat(b.ORDEN || b.Orden) || 0));
    const tramiteTiempos = tiempos.filter(t => t.id_tramite === tramiteId);

    // Header
    document.getElementById('detail-code').textContent = name;
    document.getElementById('detail-code').style.background = color;
    document.getElementById('detail-name').textContent = item.Name || name;
    document.getElementById('detail-phase').textContent = `Fase Actual: ${getFaseName(faseActual)}`;

    const estado = getEstadoTramite(item, tiempos);
    const cfg = ESTADO_CONFIG[estado] || ESTADO_CONFIG['en_proceso'];
    const statusEl = document.getElementById('detail-status');
    statusEl.innerHTML = `<span class="status-dot" style="background:${cfg.color}"></span> ${cfg.label}`;
    statusEl.style.background = cfg.bg;
    statusEl.style.color = cfg.color;

    // Determine phase statuses
    const currentFaseOrden = typeFases.findIndex(f => (f.ID_FASE_TRAMITE || f.CODIGO || f.Codigo) === faseActual);
    let completedCount = 0;
    let inProgressCount = 0;
    let pendingCount = 0;

    const phaseStatuses = typeFases.map((fase, idx) => {
        const faseCode = fase.ID_FASE_TRAMITE || fase.CODIGO || fase.Codigo;
        const tiempoData = tramiteTiempos.find(t => t.fase === faseCode);

        let status = 'pending';
        let dateText = 'Pendiente';

        if (tiempoData) {
            if (tiempoData.fecha_hora_fin && tiempoData.fecha_hora_fin !== '') {
                status = 'completed';
                dateText = `Completada - ${formatDate(new Date(tiempoData.fecha_hora_fin))}`;
                completedCount++;
            } else {
                status = 'in-progress';
                dateText = `En progreso - Iniciada ${formatDate(new Date(tiempoData.fecha_hora))}`;
                inProgressCount++;
            }
        } else if (idx < currentFaseOrden) {
            status = 'completed';
            dateText = 'Completada';
            completedCount++;
        } else if (idx === currentFaseOrden) {
            status = 'in-progress';
            dateText = 'En progreso';
            inProgressCount++;
        } else {
            pendingCount++;
        }

        return { ...fase, status, dateText };
    });

    // Phases count
    document.getElementById('detail-phases-count').textContent =
        `${completedCount} de ${typeFases.length} completadas`;

    // Phases list
    const phasesContainer = document.getElementById('detail-phases-list');
    phasesContainer.innerHTML = '';

    phaseStatuses.forEach((fase) => {
        const faseCode = fase.ID_FASE_TRAMITE || fase.CODIGO || fase.Codigo;
        const faseName = fase.NOMBRE_FASE || fase.Nombre_Fase || getFaseName(faseCode);
        const isActive = fase.status === 'in-progress' ? 'active' : '';
        const isPending = fase.status === 'pending' ? 'pending-phase' : '';

        phasesContainer.innerHTML += `
            <div class="phase-row ${isActive} ${isPending}">
                <div class="phase-dot ${fase.status}">
                    ${fase.status === 'completed' ? '<i class="lucide lucide-check"></i>' : ''}
                </div>
                <div class="phase-info">
                    <div class="phase-name">${faseName}</div>
                    <div class="phase-desc ${fase.status}">${fase.dateText}</div>
                </div>
                <span class="phase-status ${fase.status}">
                    ${fase.status === 'completed' ? 'Completada' : fase.status === 'in-progress' ? 'En Progreso' : 'Pendiente'}
                </span>
            </div>
        `;
    });

    // Info fields
    const infoContainer = document.getElementById('detail-info-fields');
    const responsable = item.Responsible || '';
    const responsableName = getUserName(responsable);
    const fechaInicio = item['Start date'] || '';
    const parroquia = item.PARROQUIA || 'N/A';
    const dex = item.DEX || 'N/A';
    const prioridad = item.Sumilla_Inicial || item.Priority || 'N/A';
    const canton = item[' CANTÓN'] || item['CANTON'] || 'N/A';

    const diasTranscurridos = fechaInicio ?
        Math.max(0, Math.floor((new Date() - new Date(fechaInicio)) / (1000 * 60 * 60 * 24))) : 0;

    infoContainer.innerHTML = `
        <div class="info-field">
            <span class="info-field-label">Solicitante:</span>
            <span class="info-field-value">${item.Name || 'N/A'}</span>
        </div>
        <div class="info-field">
            <span class="info-field-label">Responsable:</span>
            <span class="info-field-value" title="${responsable}">${responsableName.length > 30 ? responsableName.substring(0, 30) + '...' : responsableName}</span>
        </div>
        <div class="info-field">
            <span class="info-field-label">DEX:</span>
            <span class="info-field-value" style="font-size:11px">${dex}</span>
        </div>
        <div class="info-field">
            <span class="info-field-label">Fecha Inicio:</span>
            <span class="info-field-value">${fechaInicio ? formatDate(new Date(fechaInicio)) : 'N/A'}</span>
        </div>
        <div class="info-field">
            <span class="info-field-label">Dias Transcurridos:</span>
            <span class="info-field-value" style="color:#E88B00">${diasTranscurridos} dias</span>
        </div>
        <div class="info-field">
            <span class="info-field-label">Canton:</span>
            <span class="info-field-value">${canton}</span>
        </div>
        <div class="info-field">
            <span class="info-field-label">Parroquia:</span>
            <span class="info-field-value">${parroquia}</span>
        </div>
        <div class="info-field">
            <span class="info-field-label">Prioridad:</span>
            <span class="priority-badge">${prioridad}</span>
        </div>
    `;

    // Progress bar
    const totalFasesCount = typeFases.length;
    const pctComplete = totalFasesCount > 0 ? Math.round((completedCount / totalFasesCount) * 100) : 0;

    document.getElementById('detail-progress-bar').style.width = `${pctComplete}%`;
    document.getElementById('detail-progress-pct').textContent = `${pctComplete}% Completado`;
    document.getElementById('detail-progress-frac').textContent = `${completedCount} / ${totalFasesCount} fases`;

    const statsContainer = document.getElementById('detail-progress-stats');
    statsContainer.innerHTML = `
        <div class="progress-stat listas">
            <span class="progress-stat-value" style="color:#007B4F">${completedCount}</span>
            <span class="progress-stat-label">Listas</span>
        </div>
        <div class="progress-stat activa">
            <span class="progress-stat-value" style="color:#E88B00">${inProgressCount}</span>
            <span class="progress-stat-label">Activa</span>
        </div>
        <div class="progress-stat pendientes">
            <span class="progress-stat-value" style="color:#999">${pendingCount}</span>
            <span class="progress-stat-label">Pendientes</span>
        </div>
    `;
}

/* ==========================================
   PAGE 5: TENDENCIAS Y ANALISIS TEMPORAL
   ========================================== */
// Chart instances for Tendencias
let tendLineChartInst = null;
let tendEstadosChartInst = null;

function renderTendenciasPage() {
    const intake = getFilteredIntake();
    const { tiempos } = dashboardData;
    if (!intake || !tiempos) return;

    // ── KPIs ──────────────────────────────────────────────────────────
    const now = new Date();
    const thisWeekStart = new Date(now);
    thisWeekStart.setDate(now.getDate() - now.getDay());
    thisWeekStart.setHours(0, 0, 0, 0);

    const thisWeekCount = intake.filter(i => {
        const d = parseDateOnly(i['Start date']);
        return d && d >= thisWeekStart;
    }).length;

    const completados = intake.filter(i => getEstadoTramite(i, tiempos) === 'finalizado').length;
    const enProgreso = intake.filter(i => getEstadoTramite(i, tiempos) === 'en_proceso').length;
    const intakeIdsSet = new Set(intake.map(i => i.id));
    const tiemposAbiertos = tiempos.filter(t =>
        intakeIdsSet.has(t.id_tramite) && t.fecha_hora && (!t.fecha_hora_fin || t.fecha_hora_fin === '')
    );
    const duraciones = tiemposAbiertos.map(t => {
        const start = new Date(t.fecha_hora);
        return isNaN(start) ? 0 : Math.floor((now - start) / 86400000);
    }).filter(d => d >= 0);
    const promDias = duraciones.length > 0 ? (duraciones.reduce((s, d) => s + d, 0) / duraciones.length).toFixed(1) : 0;

    const kpiContainer = document.getElementById('tend-kpis');
    if (kpiContainer) {
        kpiContainer.innerHTML = `
            <div class="kpi-card">
                <span class="kpi-label">Esta Semana</span>
                <span class="kpi-value" style="color:#1A3C6E">${thisWeekCount}</span>
            </div>
            <div class="kpi-card">
                <span class="kpi-label">En Proceso</span>
                <span class="kpi-value" style="color:#E88B00">${enProgreso}</span>
            </div>
            <div class="kpi-card">
                <span class="kpi-label">Finalizados</span>
                <span class="kpi-value" style="color:#007B4F">${completados}</span>
            </div>
            <div class="kpi-card">
                <span class="kpi-label">Prom. Días Abierto</span>
                <span class="kpi-value" style="color:#A52422">${promDias}</span>
            </div>
        `;
    }

    // Date range badge
    const dates = intake.map(i => parseDateOnly(i['Start date'])).filter(Boolean).sort((a, b) => a - b);
    const rangeEl = document.getElementById('tendencias-range');
    if (rangeEl && dates.length > 0) {
        rangeEl.textContent = `${formatDate(dates[0])} - ${formatDate(dates[dates.length - 1])}`;
    }

    // ── CHART 1: Line — Tramites nuevos por semana ────────────────────
    (function renderLineChart() {
        // Build weekly buckets
        if (dates.length === 0) return;
        const minDate = new Date(dates[0]);
        minDate.setDate(minDate.getDate() - minDate.getDay()); // align to Sunday
        minDate.setHours(0, 0, 0, 0);
        const maxDate = new Date();

        const weeks = [];
        const cursor = new Date(minDate);
        while (cursor <= maxDate) {
            weeks.push(new Date(cursor));
            cursor.setDate(cursor.getDate() + 7);
        }
        const weekLabels = weeks.map(w => {
            const months = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
            return `${w.getDate()} ${months[w.getMonth()]}`;
        });

        const types = Object.keys(TYPE_NAMES);
        const typeColors = Object.values(TYPE_COLORS);
        const datasets = types.map((code, i) => {
            const color = TYPE_COLORS[code] || typeColors[i % typeColors.length];
            const counts = weeks.map(wStart => {
                const wEnd = new Date(wStart);
                wEnd.setDate(wEnd.getDate() + 7);
                return intake.filter(item => {
                    if (extractTipo(item.Fase_del_Tramite) !== code) return false;
                    const d = parseDateOnly(item['Start date']);
                    return d && d >= wStart && d < wEnd;
                }).length;
            });
            return {
                label: TYPE_NAMES[code] || code,
                data: counts,
                borderColor: color,
                backgroundColor: color + '22',
                borderWidth: 2,
                pointRadius: 0,
                pointHoverRadius: 6,
                pointBackgroundColor: '#fff',
                pointBorderColor: color,
                pointBorderWidth: 2,
                fill: true,
                tension: 0.35
            };
        }).filter(ds => ds.data.some(v => v > 0));

        const ctx = document.getElementById('tendLineChart');
        if (!ctx) return;
        if (tendLineChartInst) tendLineChartInst.destroy();
        tendLineChartInst = new Chart(ctx, {
            type: 'line',
            data: { labels: weekLabels, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: { duration: 450 },
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            title: items => `Semana del ${items[0].label}`,
                            label: item => {
                                const v = item.parsed.y;
                                if (v == null) return '';
                                return ` ${item.dataset.label}: ${v.toLocaleString()}`;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        grid: { color: '#ECEFF1', drawTicks: true },
                        ticks: { maxRotation: 45, autoSkip: true, maxTicksLimit: 14 }
                    },
                    y: {
                        beginAtZero: true,
                        grid: { color: '#ECEFF1' },
                        ticks: { precision: 0 },
                        title: { display: true, text: 'Trámites', color: '#5F6368', font: { size: 11, weight: '500' } }
                    }
                }
            }
        });

        // Custom legend
        const legendEl = document.getElementById('tend-line-legend');
        if (legendEl) {
            legendEl.innerHTML = datasets.map(ds => `
                <div class="tend-legend-item">
                    <span class="tend-legend-dot" style="background:${ds.borderColor}"></span>
                    <span>${ds.label}</span>
                </div>
            `).join('');
        }
    })();

    // ── CHART 2: Stacked bar — Estados por Tipo (TODOS los tipos) ──────
    (function renderEstadosChart() {
        const types = Object.keys(TYPE_NAMES);

        const estadoKeys = EXEC_STATE_SEGMENTS.map(s => s.key);
        const estadoLabels = {};
        const estadoColors = {};
        EXEC_STATE_SEGMENTS.forEach(s => { estadoLabels[s.key] = s.label; estadoColors[s.key] = s.color; });

        const datasets = estadoKeys.map(key => ({
            label: estadoLabels[key],
            data: types.map(code =>
                intake.filter(i => extractTipo(i.Fase_del_Tramite) === code && getEstadoTramite(i, tiempos) === key).length
            ),
            backgroundColor: estadoColors[key],
            borderRadius: 2,
            borderSkipped: false
        })).filter(ds => ds.data.some(v => v > 0));

        const ctx = document.getElementById('tendEstadosChart');
        if (!ctx) return;
        if (tendEstadosChartInst) tendEstadosChartInst.destroy();

        tendEstadosChartInst = new Chart(ctx, {
            type: 'bar',
            data: { labels: types.map(c => TYPE_NAMES[c] || c), datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: { duration: 400 },
                plugins: {
                    legend: {
                        display: true,
                        position: 'bottom',
                        align: 'start',
                        labels: {
                            boxWidth: 8,
                            boxHeight: 8,
                            padding: 16,
                            usePointStyle: true,
                            pointStyle: 'rect',
                            color: '#5F6368',
                            font: { size: 11, weight: '400' }
                        }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        callbacks: {
                            footer: items => {
                                const sum = items.reduce((s, it) => s + (it.parsed.y || 0), 0);
                                return sum ? `Total: ${sum.toLocaleString()}` : '';
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        stacked: true,
                        grid: { display: false },
                        ticks: { maxRotation: 0 }
                    },
                    y: {
                        stacked: true,
                        beginAtZero: true,
                        grid: { color: '#ECEFF1' },
                        ticks: { precision: 0 },
                        title: { display: true, text: 'Cantidad', color: '#5F6368', font: { size: 11, weight: '500' } }
                    }
                }
            }
        });
    })();

    // ── CHART 3: Trámites Derivados por Institución ──────────────────
    (function renderDerivadosInst() {
        const container = document.getElementById('tend-derivados-container');
        if (!container) return;

        const derivados = intake.filter(i => getEstadoTramite(i, tiempos) === 'en_derivacion');

        if (derivados.length === 0) {
            container.innerHTML = '<p class="tipo-empty-hint" style="padding:16px">No hay trámites en derivación con los filtros actuales.</p>';
            return;
        }

        const map1 = buildDerCatalogMap(dashboardData.derCat1);
        const map2 = buildDerCatalogMap(dashboardData.derCat2);

        // Agrupar por Cat1 (categoría principal) → sub por institución completa (Cat1 — Cat2)
        const cat1Groups = {};
        derivados.forEach(item => {
            const cat1 = cat1NombreDesdeIntake(item, map1);
            const inst = institucionDerivacionDesdeIntake(item, map1, map2);
            if (!cat1Groups[cat1]) cat1Groups[cat1] = { total: 0, sub: {} };
            cat1Groups[cat1].total++;
            cat1Groups[cat1].sub[inst] = (cat1Groups[cat1].sub[inst] || 0) + 1;
        });

        const maxVal = Math.max(...Object.values(cat1Groups).map(g => g.total), 1);
        let html = `<div style="padding:4px 0 8px;font-size:12px;color:#5F6368;font-weight:500">${derivados.length} trámite(s) en derivación</div>`;

        Object.entries(cat1Groups).sort((a, b) => b[1].total - a[1].total).forEach(([cat1, g]) => {
            html += `<div class="fase-load-tipo-header" style="border-left:3px solid #7B1FA2;padding-left:8px;margin:12px 0 6px;font-weight:600;font-size:12px;color:#7B1FA2">${escapeHtml(cat1)} <span style="font-weight:400;color:#5F6368">(${g.total})</span></div>`;
            Object.entries(g.sub).sort((a, b) => b[1] - a[1]).forEach(([inst, cnt]) => {
                const w = Math.round((cnt / maxVal) * 100);
                html += `
                    <div class="fase-load-row">
                        <span class="fase-load-label" title="${escapeHtml(inst)}">${escapeHtml(inst.length > 50 ? inst.slice(0, 48) + '…' : inst)}</span>
                        <div class="fase-load-bar-bg"><div class="fase-load-bar-fill" style="width:${w}%;background:#7B1FA2"></div></div>
                        <span class="fase-load-num">${cnt.toLocaleString()}</span>
                    </div>`;
            });
        });
        container.innerHTML = html;
    })();
}

/* ==========================================
   DER_CAT HELPERS
   ========================================== */
function buildDerCatalogMap(rows) {
    const m = new Map();
    (rows || []).forEach(r => {
        // Der_Cat1: ID_Nivel1 / Categoría Principal
        // Der_Cat2: ID_Nivel2 / Entidad Secundaria
        const id = r.ID_Nivel1 ?? r.ID_Nivel2 ?? r.ID ?? r.Id ?? r.id ?? r.Codigo ?? r.CODIGO;
        const nombre = (
            r['Categoría Principal'] || r['Categoria Principal'] ||
            r['Entidad Secundaria'] || r['Entidad Principal'] ||
            r.Nombre || r.NOMBRE || r.Descripcion || r.DESCRIPCION || ''
        ).toString().trim();
        if (id != null && nombre) m.set(String(id).trim(), nombre);
    });
    return m;
}

function institucionDerivacionDesdeIntake(item, map1, map2) {
    const raw1 = String(item.Der_Cat1 ?? item['Der Cat1'] ?? item.DER_CAT1 ?? '').trim();
    const raw2 = String(item.Der_Cat2 ?? item['Der Cat2'] ?? item.DER_CAT2 ?? '').trim();
    const n1 = raw1 ? (map1.get(raw1) || raw1) : '';
    const n2 = raw2 ? (map2.get(raw2) || raw2) : '';
    if (n1 && n2) return `${n1} — ${n2}`;
    return n1 || n2 || 'Sin institución especificada';
}

function cat1NombreDesdeIntake(item, map1) {
    const raw1 = String(item.Der_Cat1 ?? item['Der Cat1'] ?? item.DER_CAT1 ?? '').trim();
    return raw1 ? (map1.get(raw1) || raw1) : 'Sin categoría';
}

/* ==========================================
   UTILITY FUNCTIONS
   ========================================== */
function formatDate(date) {
    if (!date || isNaN(date.getTime())) return 'N/A';
    const months = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
    return `${date.getDate()} ${months[date.getMonth()]}, ${date.getFullYear()}`;
}

function ordinal(n) {
    if (n === 1) return '1er';
    if (n === 2) return '2do';
    if (n === 3) return '3er';
    return `${n}to`;
}
