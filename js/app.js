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

// Estado posibles de un tramite
const ESTADO_CONFIG = {
    'en_progreso': { label: 'En Progreso', color: '#E88B00', bg: '#FFF3E0' },
    'completado': { label: 'Completado', color: '#007B4F', bg: '#E8F5E9' },
    'iniciado': { label: 'Iniciado', color: '#1A3C6E', bg: '#E3F2FD' },
    'sin_tramite': { label: 'Sin Tramite', color: '#999', bg: '#F5F5F5' },
    'derivado': { label: 'Derivado', color: '#A52422', bg: '#FFEBEE' },
    'archivado': { label: 'Archivado', color: '#6B3A2A', bg: '#EFEBE9' }
};

/* ==========================================
   HELPERS
   ========================================== */
function extractTipo(faseCode) {
    if (!faseCode) return '';
    return faseCode.replace(/-\d+.*$/, '');
}

function getEstadoTramite(item, tiempos) {
    // Check Estado field first
    const estadoRaw = (item.Estado || '').toLowerCase().trim();
    if (estadoRaw.includes('archivado') || item.Archivado) return 'archivado';
    if (estadoRaw.includes('derivad')) return 'derivado';
    if (estadoRaw.includes('complet')) return 'completado';

    // Check tiempos records
    const tramiteTiempos = tiempos.filter(t => t.id_tramite === item.id);
    if (tramiteTiempos.length === 0) return 'iniciado';

    const hasInProgress = tramiteTiempos.some(t => !t.fecha_hora_fin || t.fecha_hora_fin === '');
    const allCompleted = tramiteTiempos.every(t => t.fecha_hora_fin && t.fecha_hora_fin !== '');

    if (allCompleted) return 'completado';
    if (hasInProgress) return 'en_progreso';
    return 'iniciado';
}

function getUserName(email) {
    if (!dashboardData || !email) return email || 'Sin asignar';
    const user = dashboardData.usuarios.find(u => u.Email.toLowerCase() === email.toLowerCase());
    return user ? user.Nombre : email;
}

function getUserNameShort(email) {
    const name = getUserName(email);
    if (name === email) return email;
    // Return first name + last name
    const parts = name.split(' ');
    if (parts.length >= 2) return parts[0] + ' ' + parts[1];
    return name;
}

function parseDateOnly(dateString) {
    if (!dateString) return null;
    const date = new Date(dateString);
    if (isNaN(date.getTime())) return null;
    // Strip time for pure date comparisons
    date.setHours(0, 0, 0, 0);
    return date;
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
        data = data.filter(i => (i[' CANTÓN'] || '').trim() === filtroCanton);
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

async function loadData() {
    showLoading(true);
    try {
        dashboardData = await fetchAllDashboardData();
        console.log('Dashboard data loaded:', dashboardData);
        buildFilters();
        renderAllPages();
    } catch (error) {
        console.error('Failed to load data:', error);
        dashboardData = getFallbackData();
        buildFilters();
        renderAllPages();
    } finally {
        showLoading(false);
    }
}

function showLoading(show) {
    const el = document.getElementById('loading');
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
    const cantonesUnicos = [...new Set(dashboardData.intake.map(i => (i[' CANTÓN'] || '').trim()).filter(Boolean))].sort();
    const parroquiasUnicas = [...new Set(dashboardData.intake.map(i => (i.PARROQUIA || '').trim()).filter(Boolean))].sort();

    // Populate Tipo filter
    const tipoSelect = document.getElementById('filtro-tipo');
    if (tipoSelect) {
        tipoSelect.innerHTML = '<option value="Todos">Todos los Tipos</option>';
        tiposUnicos.forEach(t => {
            tipoSelect.innerHTML += `<option value="${t}">${t} - ${TYPE_NAMES[t] || t}</option>`;
        });
    }

    // Populate Fase filter
    const faseSelect = document.getElementById('filtro-fase');
    if (faseSelect) {
        faseSelect.innerHTML = '<option value="Todos">Todas las Fases</option>';
        fasesUnicas.forEach(f => {
            const faseInfo = dashboardData.fases.find(fs => fs.Codigo === f);
            const label = faseInfo ? `${f} - ${faseInfo.Nombre_Fase}` : f;
            faseSelect.innerHTML += `<option value="${f}">${label}</option>`;
        });
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
    }

    // Populate Cantón filter
    const cantonSelect = document.getElementById('filtro-canton');
    if (cantonSelect) {
        cantonSelect.innerHTML = '<option value="Todos">Todos</option>';
        cantonesUnicos.forEach(c => {
            cantonSelect.innerHTML += `<option value="${c}">${c}</option>`;
        });
    }

    // Populate Parroquia filter
    const parroquiaSelect = document.getElementById('filtro-parroquia');
    if (parroquiaSelect) {
        parroquiaSelect.innerHTML = '<option value="Todos">Todas</option>';
        parroquiasUnicas.forEach(p => {
            parroquiaSelect.innerHTML += `<option value="${p}">${p}</option>`;
        });
    }

    // Sync mobile filter dropdowns
    syncMobileFilters(tiposUnicos, fasesUnicas, emailsEnUso, cantonesUnicos, parroquiasUnicas);
}

function syncMobileFilters(tiposUnicos, fasesUnicas, emailsEnUso, cantonesUnicos, parroquiasUnicas) {
    // Mobile Tipo
    const tipoMobile = document.getElementById('filtro-tipo-mobile');
    if (tipoMobile) {
        tipoMobile.innerHTML = '<option value="Todos">Todos los Tipos</option>';
        tiposUnicos.forEach(t => {
            tipoMobile.innerHTML += `<option value="${t}">${t} - ${TYPE_NAMES[t] || t}</option>`;
        });
        tipoMobile.value = filtroTipo;
    }

    // Mobile Fase
    const faseMobile = document.getElementById('filtro-fase-mobile');
    if (faseMobile) {
        faseMobile.innerHTML = '<option value="Todos">Todas las Fases</option>';
        fasesUnicas.forEach(f => {
            const faseInfo = dashboardData.fases.find(fs => fs.Codigo === f);
            const label = faseInfo ? `${f} - ${faseInfo.Nombre_Fase}` : f;
            faseMobile.innerHTML += `<option value="${f}">${label}</option>`;
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
        parroquiasUnicas.forEach(p => {
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

    // Scroll to top on page change
    const mainContent = document.querySelector('.main-content');
    if (mainContent) mainContent.scrollTop = 0;
}

/* ==========================================
   RENDER ALL PAGES
   ========================================== */
function renderAllPages() {
    if (!dashboardData) return;
    renderRankingPage();
    renderProductividadPage();
    renderTipoPage();
}

/* ==========================================
   PAGE 1: RANKING POR TIPO + ESTADO
   ========================================== */
function renderRankingPage() {
    const intake = getFilteredIntake();
    const { tiempos } = dashboardData;

    // Date range
    const dateEl = document.getElementById('date-range');
    if (intake.length > 0) {
        const dates = intake.map(i => new Date(i['Start date'])).filter(d => !isNaN(d.getTime())).sort((a, b) => a - b);
        if (dates.length > 0) {
            dateEl.textContent = `${formatDate(dates[0])} - ${formatDate(dates[dates.length - 1])}`;
        } else dateEl.textContent = 'Sin fechas';
    } else dateEl.textContent = 'Sin datos';

    // Count by type
    const typeCounts = {};
    Object.keys(TYPE_NAMES).forEach(code => { typeCounts[code] = 0; });
    intake.forEach(item => {
        const tipo = extractTipo(item.Fase_del_Tramite);
        if (tipo) typeCounts[tipo] = (typeCounts[tipo] || 0) + 1;
    });

    const sorted = Object.entries(typeCounts).sort((a, b) => b[1] - a[1]);
    const total = sorted.reduce((sum, [, v]) => sum + v, 0);

    // Count by estado
    const estadoCounts = {};
    intake.forEach(item => {
        const estado = getEstadoTramite(item, tiempos);
        estadoCounts[estado] = (estadoCounts[estado] || 0) + 1;
    });

    // Ranking cards (top 3)
    const cardsContainer = document.getElementById('ranking-cards');
    cardsContainer.innerHTML = '';
    sorted.filter(([, c]) => c > 0).slice(0, 3).forEach(([code, count], idx) => {
        const name = TYPE_NAMES[code] || code;
        cardsContainer.innerHTML += `
            <div class="ranking-card" style="border-left: 4px solid ${TYPE_COLORS[code]}">
                <div class="ranking-position pos-${idx + 1}">${idx + 1}</div>
                <div class="ranking-info">
                    <div class="ranking-name">${ordinal(idx + 1)} mas frecuente</div>
                    <div class="ranking-value" style="color:${TYPE_COLORS[code]}">${count}</div>
                    <div class="ranking-label">${name}</div>
                </div>
            </div>
        `;
    });

    // Total
    document.getElementById('total-value').textContent = total;

    // Estado badges
    const estadoContainer = document.getElementById('estado-badges');
    if (estadoContainer) {
        estadoContainer.innerHTML = '';
        Object.entries(estadoCounts).forEach(([estado, count]) => {
            const cfg = ESTADO_CONFIG[estado] || ESTADO_CONFIG['iniciado'];
            estadoContainer.innerHTML += `
                <div class="estado-badge" style="background:${cfg.bg};color:${cfg.color};border:1px solid ${cfg.color}20">
                    <span class="estado-dot" style="background:${cfg.color}"></span>
                    ${cfg.label}: <strong>${count}</strong>
                </div>
            `;
        });
    }

    // Table
    const tbody = document.getElementById('ranking-table-body');
    tbody.innerHTML = '';
    sorted.forEach(([code, count]) => {
        const name = TYPE_NAMES[code] || code;
        tbody.innerHTML += `
            <tr>
                <td>
                    <span style="display:inline-flex;align-items:center;gap:8px;">
                        <span class="code-badge" style="background:${TYPE_COLORS[code] || '#888'}">${code}</span>
                        ${name}
                    </span>
                </td>
                <td style="text-align:center;font-weight:600">${count}</td>
            </tr>
        `;
    });
    document.getElementById('grand-total').textContent = total;

    // Donut chart
    renderDonutChart(sorted.filter(([, c]) => c > 0), total);
}

function renderDonutChart(sorted, total) {
    const canvas = document.getElementById('donutChart');
    const ctx = canvas.getContext('2d');
    const size = 200;
    const center = size / 2;
    const outerRadius = 90;
    const innerRadius = 55;

    ctx.clearRect(0, 0, size, size);

    if (total === 0) {
        ctx.beginPath();
        ctx.arc(center, center, outerRadius, 0, 2 * Math.PI);
        ctx.arc(center, center, innerRadius, 2 * Math.PI, 0, true);
        ctx.closePath();
        ctx.fillStyle = '#E0E0E0';
        ctx.fill();
    } else {
        let startAngle = -Math.PI / 2;
        sorted.forEach(([code, count]) => {
            if (count === 0) return;
            const sliceAngle = (count / total) * 2 * Math.PI;
            ctx.beginPath();
            ctx.arc(center, center, outerRadius, startAngle, startAngle + sliceAngle);
            ctx.arc(center, center, innerRadius, startAngle + sliceAngle, startAngle, true);
            ctx.closePath();
            ctx.fillStyle = TYPE_COLORS[code] || '#888';
            ctx.fill();
            startAngle += sliceAngle;
        });
    }

    // Center text
    ctx.fillStyle = '#1A3C6E';
    ctx.font = '700 28px Inter';
    ctx.textAlign = 'center';
    ctx.fillText(total, center, center + 4);
    ctx.font = '400 10px Inter';
    ctx.fillStyle = '#888';
    ctx.fillText('Total', center, center + 18);

    // Legend
    const legendContainer = document.getElementById('donut-legend');
    legendContainer.innerHTML = '';
    sorted.forEach(([code, count]) => {
        if (count === 0) return;
        const name = TYPE_NAMES[code] || code;
        legendContainer.innerHTML += `
            <div class="legend-item">
                <div class="legend-dot" style="background:${TYPE_COLORS[code] || '#888'}"></div>
                <span>${name} (${count})</span>
            </div>
        `;
    });
}

/* ==========================================
   PAGE 2: PRODUCTIVIDAD POR USUARIO
   ========================================== */
function renderProductividadPage() {
    const intake = getFilteredIntake();
    const { tiempos } = dashboardData;

    const intakeIds = new Set(intake.map(i => i.id));
    const filteredTiempos = tiempos.filter(t => intakeIds.has(t.id_tramite));

    // Group by responsible (user)
    const respMap = {};
    intake.forEach(item => {
        const resp = item.Responsible || 'Sin asignar';
        if (!respMap[resp]) respMap[resp] = { total: 0, completadas: 0, enProgreso: 0, iniciadas: 0 };
        respMap[resp].total++;

        const estado = getEstadoTramite(item, tiempos);
        if (estado === 'completado') respMap[resp].completadas++;
        else if (estado === 'en_progreso') respMap[resp].enProgreso++;
        else respMap[resp].iniciadas++;
    });

    const responsables = Object.entries(respMap).sort((a, b) => b[1].total - a[1].total);
    const totalTramites = intake.length;
    const totalComp = responsables.reduce((s, [, r]) => s + r.completadas, 0);
    const totalProg = responsables.reduce((s, [, r]) => s + r.enProgreso, 0);
    const totalIni = responsables.reduce((s, [, r]) => s + r.iniciadas, 0);

    // KPIs
    const kpiContainer = document.getElementById('prod-kpis');
    kpiContainer.innerHTML = `
        <div class="kpi-card">
            <span class="kpi-label">Total Tramites</span>
            <span class="kpi-value" style="color:#1A3C6E">${totalTramites}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">En Progreso</span>
            <span class="kpi-value" style="color:#E88B00">${totalProg}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">Completados</span>
            <span class="kpi-value" style="color:#007B4F">${totalComp}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">Responsables</span>
            <span class="kpi-value" style="color:#A52422">${responsables.length}</span>
        </div>
    `;

    // Responsables table
    const tbody = document.getElementById('prod-table-body');
    tbody.innerHTML = '';
    responsables.forEach(([resp, r]) => {
        const displayName = getUserNameShort(resp);
        const user = dashboardData.usuarios.find(u => u.Email.toLowerCase() === resp.toLowerCase());
        const rol = user ? user.Rol : '';
        const rolBadge = rol ? `<span class="rol-badge rol-${rol.toLowerCase()}">${rol}</span>` : '';

        tbody.innerHTML += `
            <tr>
                <td>
                    <div style="display:flex;flex-direction:column;gap:2px">
                        <span title="${resp}" style="font-weight:600">${displayName}</span>
                        ${rolBadge}
                    </div>
                </td>
                <td style="text-align:center">${r.total}</td>
                <td style="text-align:center;color:#E88B00;font-weight:600">${r.enProgreso}</td>
                <td style="text-align:center;color:#007B4F;font-weight:600">${r.completadas}</td>
            </tr>
        `;
    });

    document.getElementById('prod-total-fases').textContent = totalTramites;
    document.getElementById('prod-total-comp').textContent = totalComp;
    document.getElementById('prod-total-prog').textContent = totalProg;

    // Progress bars by type (avance basado en fase actual)
    const barsContainer = document.getElementById('prod-bars');
    barsContainer.innerHTML = '';

    const typeProgress = {};
    intake.forEach(item => {
        const tipo = extractTipo(item.Fase_del_Tramite);
        if (!tipo) return;
        if (!typeProgress[tipo]) typeProgress[tipo] = { total: 0, totalAdvance: 0 };
        typeProgress[tipo].total++;

        const faseInfo = dashboardData.fases.find(f => f.Codigo === item.Fase_del_Tramite);
        typeProgress[tipo].totalAdvance += faseInfo ? (parseFloat(faseInfo.Avance) || 0) : 0;
    });

    Object.entries(typeProgress).forEach(([code, data]) => {
        const avgAdvance = data.total > 0 ? Math.round(data.totalAdvance / data.total) : 0;
        const name = TYPE_NAMES[code] || code;
        const color = TYPE_COLORS[code] || '#888';
        barsContainer.innerHTML += `
            <div class="bar-item">
                <div class="bar-label-row">
                    <span class="bar-name">${name} (${data.total})</span>
                    <span class="bar-pct" style="color:${color}">${avgAdvance}%</span>
                </div>
                <div class="bar-bg">
                    <div class="bar-fill" style="width:${avgAdvance}%;background:${color}"></div>
                </div>
            </div>
        `;
    });
}

/* ==========================================
   PAGE 3: TIPO DE TRAMITE - FASES Y ESTADOS
   ========================================== */
function renderTipoPage() {
    const intake = getFilteredIntake();
    const { fases, tiempos } = dashboardData;

    const typePhaseCount = {};
    fases.forEach(f => { typePhaseCount[f.Tipo] = (typePhaseCount[f.Tipo] || 0) + 1; });

    const typeTramiteCount = {};
    intake.forEach(item => {
        const tipo = extractTipo(item.Fase_del_Tramite);
        typeTramiteCount[tipo] = (typeTramiteCount[tipo] || 0) + 1;
    });

    const totalTypes = Object.keys(TYPE_NAMES).length;
    const totalPhases = fases.length;

    document.getElementById('tipos-badge').textContent = `${intake.length} Tramites Activos`;

    // KPIs
    const kpiContainer = document.getElementById('tipo-kpis');
    kpiContainer.innerHTML = `
        <div class="kpi-card">
            <span class="kpi-label">Total Tipos</span>
            <span class="kpi-value" style="color:#1A3C6E">${totalTypes}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">Total Fases</span>
            <span class="kpi-value" style="color:#007B4F">${totalPhases}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">Tramites</span>
            <span class="kpi-value" style="color:#EAB308">${intake.length}</span>
        </div>
        <div class="kpi-card">
            <span class="kpi-label">Prom. Fases/Tipo</span>
            <span class="kpi-value" style="color:#E88B00">${totalTypes > 0 ? (totalPhases / totalTypes).toFixed(1) : 0}</span>
        </div>
    `;

    // Table - each tramite
    const tbody = document.getElementById('tipo-table-body');
    tbody.innerHTML = '';

    Object.entries(TYPE_NAMES).forEach(([code, name]) => {
        const tramites = intake.filter(i => extractTipo(i.Fase_del_Tramite) === code);
        const phaseCount = typePhaseCount[code] || 0;

        if (tramites.length === 0) {
            tbody.innerHTML += `
                <tr>
                    <td><span class="code-badge" style="background:${TYPE_COLORS[code]}">${code}</span></td>
                    <td style="font-weight:600;color:#1A3C6E">${name}</td>
                    <td style="text-align:center">${phaseCount}</td>
                    <td><span class="phase-pill" style="background:#F5F5F5;color:#999">-</span></td>
                    <td>-</td>
                    <td><span class="status-pill sin-tramite">Sin tramite</span></td>
                </tr>
            `;
        } else {
            tramites.forEach((item) => {
                const faseActual = item.Fase_del_Tramite || '-';
                const estado = getEstadoTramite(item, tiempos);
                const cfg = ESTADO_CONFIG[estado] || ESTADO_CONFIG['iniciado'];
                const responsable = getUserNameShort(item.Responsible);

                tbody.innerHTML += `
                    <tr class="clickable-row" onclick="openDetail('${item.id}')">
                        <td><span class="code-badge" style="background:${TYPE_COLORS[code]}">${code}</span></td>
                        <td style="font-weight:600;color:#1A3C6E" title="${item.Name || ''}">${(item.Name || name).substring(0, 30)}</td>
                        <td style="text-align:center">${phaseCount}</td>
                        <td><span class="phase-pill" style="background:#E3F2FD;color:#1565C0">${faseActual}</span></td>
                        <td title="${item.Responsible || ''}" style="font-size:11px">${responsable}</td>
                        <td><span class="status-pill" style="background:${cfg.bg};color:${cfg.color}">${cfg.label}</span></td>
                    </tr>
                `;
            });
        }
    });

    // Horizontal bar chart - tramites by type
    const barsContainer = document.getElementById('tipo-bars');
    barsContainer.innerHTML = '';
    const maxTramites = Math.max(...Object.values(typeTramiteCount), 1);

    Object.entries(TYPE_NAMES).forEach(([code]) => {
        const count = typeTramiteCount[code] || 0;
        const widthPct = (count / maxTramites) * 100;
        const color = TYPE_COLORS[code];

        barsContainer.innerHTML += `
            <div class="hbar-item">
                <span class="hbar-label" style="color:${color}">${code}</span>
                <div class="hbar-bg">
                    <div class="hbar-fill" style="width:${widthPct}%;background:${color}"></div>
                </div>
                <span class="hbar-value" style="color:${color}">${count}</span>
            </div>
        `;
    });

    // Summary
    const summaryContainer = document.getElementById('tipo-summary');
    summaryContainer.innerHTML = `
        <div class="summary-item">
            <span class="summary-value" style="color:#1A3C6E">${intake.length}</span>
            <span class="summary-label">Tramites</span>
        </div>
        <div class="summary-item">
            <span class="summary-value" style="color:#007B4F">${totalPhases}</span>
            <span class="summary-label">Fases</span>
        </div>
        <div class="summary-item">
            <span class="summary-value" style="color:#EAB308">${Object.values(typeTramiteCount).filter(v => v > 0).length}</span>
            <span class="summary-label">Tipos Activos</span>
        </div>
    `;
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
    const typeFases = fases.filter(f => f.Tipo === typeCode).sort((a, b) => (a.Orden || 0) - (b.Orden || 0));
    const tramiteTiempos = tiempos.filter(t => t.id_tramite === tramiteId);

    // Header
    document.getElementById('detail-code').textContent = typeCode;
    document.getElementById('detail-code').style.background = color;
    document.getElementById('detail-name').textContent = item.Name || name;
    document.getElementById('detail-phase').textContent = `Fase Actual: ${faseActual}`;

    const estado = getEstadoTramite(item, tiempos);
    const cfg = ESTADO_CONFIG[estado] || ESTADO_CONFIG['iniciado'];
    const statusEl = document.getElementById('detail-status');
    statusEl.innerHTML = `<span class="status-dot" style="background:${cfg.color}"></span> ${cfg.label}`;
    statusEl.style.background = cfg.bg;
    statusEl.style.color = cfg.color;

    // Determine phase statuses
    const currentFaseOrden = typeFases.findIndex(f => f.Codigo === faseActual);
    let completedCount = 0;
    let inProgressCount = 0;
    let pendingCount = 0;

    const phaseStatuses = typeFases.map((fase, idx) => {
        const faseCode = fase.Codigo;
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
        const faseCode = fase.Codigo;
        const faseName = fase.Nombre_Fase || faseCode;
        const isActive = fase.status === 'in-progress' ? 'active' : '';
        const isPending = fase.status === 'pending' ? 'pending-phase' : '';

        phasesContainer.innerHTML += `
            <div class="phase-row ${isActive} ${isPending}">
                <div class="phase-dot ${fase.status}">
                    ${fase.status === 'completed' ? '<i class="lucide lucide-check"></i>' : ''}
                </div>
                <div class="phase-info">
                    <div class="phase-name">${faseCode}: ${faseName}</div>
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
