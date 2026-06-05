// Configuración de LocalStorage
const STORAGE_KEY = 'adminpro_db';
const defaultData = {
    entradas: [],
    salidas: [],
    deudasCobrar: [],
    deudasPagar: [],
    trabajadores: [],
    proformas: []
};

// Utilidades de Datos y Migración
const DB = {
    load: () => {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (raw) {
            let data = JSON.parse(raw);
            // Migración desde versión anterior
            if (data.ventas) { data.entradas = data.ventas; delete data.ventas; }
            if (data.compras) { data.salidas = data.compras; delete data.compras; }
            return { ...defaultData, ...data };
        }
        return { ...defaultData };
    },
    save: (data) => {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
        if(currentView === 'dashboard') renderDashboard(); 
    },
    clear: () => {
        if(confirm("¿Estás seguro de borrar TODOS los datos locamente?")) {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(defaultData));
            location.reload();
        }
    }
};

let appData = DB.load();
let currentView = 'dashboard';
let myChart = null; // Instancia global de gráfico
let editingDeuda = { type: null, id: null };
const btnThemeToggle = document.getElementById('theme-toggle');
if(btnThemeToggle) {
    btnThemeToggle.addEventListener('click', () => {
        document.body.classList.toggle('light-mode');
        const isLight = document.body.classList.contains('light-mode');
        btnThemeToggle.innerHTML = isLight ? '<i class="ph ph-moon"></i>' : '<i class="ph ph-sun"></i>';
        if(currentView === 'dashboard') renderDashboard();
    });
}

const viewsContainer = document.getElementById('content-body');
const viewTitle = document.getElementById('view-title');
const navItems = document.querySelectorAll('.nav-item');
const btnExport = document.getElementById('btn-export-data');
const btnBackupDownload = document.getElementById('btn-backup-download');
const btnBackupUpload = document.getElementById('btn-backup-upload');
const fileBackupUpload = document.getElementById('file-backup-upload');
const printView = document.getElementById('print-view');

function setView(view) {
    currentView = view;
    navItems.forEach(nav => nav.classList.remove('active'));
    const targetNav = document.querySelector(`[data-target="${view}"]`);
    if(targetNav) targetNav.classList.add('active');
    
    viewTitle.textContent = view.charAt(0).toUpperCase() + view.slice(1);
    
    switch(view) {
        case 'dashboard': renderDashboard(); break;
        case 'entradas': renderTransactionView('entradas'); break;
        case 'salidas': renderTransactionView('salidas'); break;
        case 'deudas-cobrar': renderDeudasView('deudasCobrar'); break;
        case 'deudas-pagar': renderDeudasView('deudasPagar'); break;
        case 'trabajadores': renderTrabajadores(); break;
        case 'proformas': renderProformas(); break;
    }
}

navItems.forEach(item => {
    item.addEventListener('click', (e) => setView(e.currentTarget.dataset.target));
});
btnExport.addEventListener('click', () => {
    if(!window.XLSX) return alert("Cargando librería de Excel, intente en un segundo.");
    const wb = XLSX.utils.book_new();

    // Entradas
    const wsEntradas = XLSX.utils.json_to_sheet(appData.entradas || []);
    XLSX.utils.book_append_sheet(wb, wsEntradas, "Entradas");

    // Salidas
    const wsSalidas = XLSX.utils.json_to_sheet(appData.salidas || []);
    XLSX.utils.book_append_sheet(wb, wsSalidas, "Salidas");

    // Deudas
    const wsDeudasCobrar = XLSX.utils.json_to_sheet(appData.deudasCobrar || []);
    XLSX.utils.book_append_sheet(wb, wsDeudasCobrar, "Deudas por Cobrar");

    const wsDeudasPagar = XLSX.utils.json_to_sheet(appData.deudasPagar || []);
    XLSX.utils.book_append_sheet(wb, wsDeudasPagar, "Deudas por Pagar");

    // Trabajadores
    const wsTrabajadores = XLSX.utils.json_to_sheet(appData.trabajadores || []);
    XLSX.utils.book_append_sheet(wb, wsTrabajadores, "Trabajadores");

    // Proformas (Flattening)
    const proformasExport = (appData.proformas || []).map(p => {
        return {
            "Nro Cotización": p.nroCotizacion || "",
            Fecha: p.fecha || "",
            Cliente: p.cliente || "",
            RUC: p.ruc || "",
            Placa: p.placa || "",
            "Total A Pagar": p.totalAPagar || 0,
            "Cant. Items de Obra": (p.manoObra || []).length,
            "Cant. Repuestos": (p.repuestos || []).length
        }
    });
    const wsProformas = XLSX.utils.json_to_sheet(proformasExport);
    XLSX.utils.book_append_sheet(wb, wsProformas, "Proformas");
    // Detalle de Proformas
    const proformasDetalleExport = [];
    (appData.proformas || []).forEach(p => {
        (p.manoObra || []).forEach(mo => proformasDetalleExport.push({ "Nro Cotización": p.nroCotizacion || "", "Tipo": "Mano de Obra", "Descripción": mo.desc, "Cantidad": mo.cant, "Descuento": mo.descuento || 0, "Importe Total": mo.importe }));
        (p.repuestos || []).forEach(rep => proformasDetalleExport.push({ "Nro Cotización": p.nroCotizacion || "", "Tipo": "Repuesto", "Descripción": rep.desc, "Cantidad": rep.cant, "Descuento": rep.descuento || 0, "Importe Total": rep.importe }));
    });
    const wsProfDetalles = XLSX.utils.json_to_sheet(proformasDetalleExport);
    XLSX.utils.book_append_sheet(wb, wsProfDetalles, "Detalles Proformas");

    XLSX.writeFile(wb, "Reporte_AdminPro.xlsx");
});

// Funciones de Respaldo Nativo (JSON)
btnBackupDownload.addEventListener('click', () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(appData));
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", "Respaldo_AdminPro_" + new Date().toISOString().split('T')[0] + ".json");
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
});

btnBackupUpload.addEventListener('click', () => {
    fileBackupUpload.click();
});

fileBackupUpload.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = function(evt) {
        try {
            const uploadedData = JSON.parse(evt.target.result);
            if(uploadedData.entradas && uploadedData.proformas) {
                if(confirm("¡ATENCIÓN! Esto sobreescribirá todos los datos actuales con los del archivo. ¿Continuar?")) {
                    appData = { ...defaultData, ...uploadedData };
                    DB.save(appData);
                    alert("Copia de seguridad restaurada con éxito.");
                    location.reload();
                }
            } else {
                alert("El archivo no es un respaldo válido de AdminPro.");
            }
        } catch(err) {
            alert("Error al intentar leer el archivo de respaldo.");
        }
        fileBackupUpload.value = "";
    };
    reader.readAsText(file);
});

const formatMoney = (amount) => `S/ ${parseFloat(amount).toFixed(2)}`;
const generateId = () => Math.random().toString(36).substr(2, 9);

// --- VISTA: DASHBOARD ---
// Lógica de Filtros ("dia", "semana", "mes", "todo")
function isOnSameDay(d1, d2) {
    return d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();
}
function isOnSameWeek(d1, d2) {
    // Basic week check (ignoring year overflow for simplicity)
    const t1 = d1.getTime(), t2 = d2.getTime();
    return Math.abs(t1 - t2) <= 7 * 24 * 60 * 60 * 1000 && d1.getMonth() === d2.getMonth();
}
function isOnSameMonth(d1, d2) {
    return d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth();
}

let activeFilter = 'todo';

let currentDeudaFilters = {
    deudasCobrar: { persona: '', fecha: '' },
    deudasPagar: { persona: '', fecha: '' }
};

window.setDashboardFilter = function(val) {
    activeFilter = val;
    renderDashboard();
}

function renderDashboard() {
    const today = new Date();
    
    // Filtrar data
    let fEntradas = appData.entradas;
    let fSalidas = appData.salidas;
    let fCobrar = appData.deudasCobrar || [];
    let fPagar = appData.deudasPagar || [];

    if (activeFilter !== 'todo') {
        const filterFn = activeFilter === 'dia' ? isOnSameDay 
                       : activeFilter === 'semana' ? isOnSameWeek 
                       : isOnSameMonth;
                       
        fEntradas = appData.entradas.filter(e => filterFn(new Date(e.fecha.replace(/-/g, '\/')), today));
        fSalidas = appData.salidas.filter(s => filterFn(new Date(s.fecha.replace(/-/g, '\/')), today));
        fCobrar = (appData.deudasCobrar || []).filter(e => filterFn(new Date(e.fecha.replace(/-/g, '\/')), today));
        fPagar = (appData.deudasPagar || []).filter(s => filterFn(new Date(s.fecha.replace(/-/g, '\/')), today));
    }

    const totalCobrar = fCobrar.reduce((acc, curr) => acc + parseFloat(curr.monto), 0);
    const totalPagar = fPagar.reduce((acc, curr) => acc + parseFloat(curr.monto), 0);

    const totalEntradas = fEntradas.reduce((acc, curr) => acc + parseFloat(curr.monto), 0);
    const totalSalidas = fSalidas.reduce((acc, curr) => acc + parseFloat(curr.monto), 0);
    const balance = totalEntradas - totalSalidas;
    
    viewsContainer.innerHTML = `
        <div style="margin-bottom:1rem;">
            <label style="color:var(--text-muted); margin-right: 0.5rem;">Filtro:</label>
            <select class="form-input" style="width: auto; display: inline-block; padding:0.4rem 1rem;" onchange="setDashboardFilter(this.value)">
                <option value="todo" ${activeFilter === 'todo' ? 'selected' : ''}>Todos los tiempos</option>
                <option value="mes" ${activeFilter === 'mes' ? 'selected' : ''}>Este Mes</option>
                <option value="semana" ${activeFilter === 'semana' ? 'selected' : ''}>Esta Semana</option>
                <option value="dia" ${activeFilter === 'dia' ? 'selected' : ''}>Hoy</option>
            </select>
        </div>

        <div class="grid-cards fade-in" style="margin-bottom: 1rem; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));">
            <div class="glass-card stat-card" style="padding: 1rem;">
                <span class="stat-title" style="font-size:0.75rem">Entradas</span>
                <span class="stat-value income" style="font-size:1.5rem">${formatMoney(totalEntradas)}</span>
            </div>
            <div class="glass-card stat-card" style="padding: 1rem;">
                <span class="stat-title" style="font-size:0.75rem">Salidas</span>
                <span class="stat-value expense" style="font-size:1.5rem">${formatMoney(totalSalidas)}</span>
            </div>
            <div class="glass-card stat-card" style="padding: 1rem;">
                <span class="stat-title" style="font-size:0.75rem">Balance</span>
                <span class="stat-value balance" style="font-size:1.5rem">${formatMoney(balance)}</span>
            </div>
            <div class="glass-card stat-card" style="padding: 1rem; background: var(--warning-bg); border-color: var(--warning-border);">
                <span class="stat-title" style="color: var(--warning-text); font-size:0.75rem">Por Cobrar</span>
                <span class="stat-value" style="color: var(--warning-text); font-size:1.5rem">${formatMoney(totalCobrar)}</span>
            </div>
            <div class="glass-card stat-card" style="padding: 1rem; background: var(--danger-bg); border-color: var(--danger-border);">
                <span class="stat-title" style="color: var(--danger-text); font-size:0.75rem">Por Pagar</span>
                <span class="stat-value" style="color: var(--danger-text); font-size:1.5rem">${formatMoney(totalPagar)}</span>
            </div>
        </div>
        
        <div class="layout-split fade-in" style="grid-template-columns: 1fr 280px;">
            <div class="glass-card">
                <h3 class="gradient-text">Últimas Transacciones</h3>
                <table class="table-glass" style="margin-top:0.5rem">
                    <thead><tr><th>Tipo</th><th>Concepto</th><th>Monto</th></tr></thead>
                    <tbody>
                        ${[...fEntradas.map(v => ({...v, tipo: 'Entrada'})), ...fSalidas.map(c => ({...c, tipo: 'Salida'}))]
                            .sort((a,b) => new Date(b.fecha) - new Date(a.fecha))
                            .slice(0, 10)
                            .map(t => `
                            <tr>
                                <td><span class="badge ${t.tipo === 'Entrada' ? 'badge-income' : 'badge-expense'}">${t.tipo}</span></td>
                                <td>${t.concepto}</td>
                                <td>${formatMoney(t.monto)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
            <div class="glass-card" style="display:flex; flex-direction:column; align-items:center;">
                <h3 class="gradient-text" style="width:100%">Gráfico General</h3>
                <div style="width: 200px; height: 200px; margin-top:1rem;">
                    <canvas id="balanceChart"></canvas>
                </div>
            </div>
        </div>
    `;

    // Renderizar Chart.js si hay elementos
    const ctx = document.getElementById('balanceChart');
    if (ctx && window.Chart) {
        if(myChart) myChart.destroy();
        myChart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Entradas', 'Salidas'],
                datasets: [{
                    data: [totalEntradas, totalSalidas],
                    backgroundColor: ['#00d2ff', '#ff4b2b'],
                    borderWidth: 0,
                    hoverOffset: 4
                }]
            },
            options: {
                responsive: true,
                plugins: { legend: { position: 'bottom', labels: { color: document.body.classList.contains('light-mode') ? '#4b5563' : '#fff'} } }
            }
        });
    }
}

// --- VISTAS: ENTRADAS / SALIDAS ---
window.delTransaction = function(type, id) {
    if(confirm("¿Seguro de borrar este registro?")) {
        appData[type] = appData[type].filter(x => x.id !== id);
        DB.save(appData);
        if(currentView === type) renderTransactionView(type);
    }
}

function renderTransactionView(type) {
    const dataList = appData[type]; // 'entradas' o 'salidas'
    const nameCap = type.charAt(0).toUpperCase() + type.slice(1);
    
    viewsContainer.innerHTML = `
        <div class="layout-split fade-in">
            <div class="glass-card">
                <h3 class="gradient-text" style="margin-bottom: 1rem">Registrar ${nameCap}</h3>
                <form id="form-${type}">
                    <div class="form-group">
                        <label>Fecha</label>
                        <input type="date" id="t-fecha" class="form-input" required value="${new Date().toISOString().split('T')[0]}">
                    </div>
                    <div class="form-group">
                        <label>Concepto / Detalle</label>
                        <input type="text" id="t-concepto" class="form-input" required placeholder="Ej: Pago de cliente / Materiales...">
                    </div>
                    <div class="form-group">
                        <label>Monto</label>
                        <input type="number" step="0.01" id="t-monto" class="form-input" required placeholder="0.00">
                    </div>
                    <button type="submit" class="btn-primary">Guardar Registro</button>
                </form>
            </div>
            
            <div class="glass-card">
                <h3 class="gradient-text" style="margin-bottom: 1rem">Historial de ${nameCap}</h3>
                <table class="table-glass">
                    <thead><tr><th>Fecha</th><th>Concepto</th><th>Monto</th><th style="width:50px">Acción</th></tr></thead>
                    <tbody>
                        ${dataList.map(item => `
                            <tr>
                                <td>${item.fecha}</td>
                                <td>${item.concepto}</td>
                                <td class="${type === 'entradas' ? 'stat-value income' : 'stat-value expense'}" style="font-size:1rem">${formatMoney(item.monto)}</td>
                                <td>
                                    <button title="Borrar" onclick="delTransaction('${type}', '${item.id}')" style="background:transparent; border:none; cursor:pointer; color:var(--danger); font-size:1.2rem;"><i class="ph ph-trash"></i></button>
                                </td>
                            </tr>
                        `).join('') || '<tr><td colspan="4">Vacio</td></tr>'}
                    </tbody>
                </table>
            </div>
        </div>
    `;

    document.getElementById(`form-${type}`).addEventListener('submit', (e) => {
        e.preventDefault();
        const nRecord = {
            id: generateId(),
            fecha: document.getElementById('t-fecha').value,
            concepto: document.getElementById('t-concepto').value,
            monto: document.getElementById('t-monto').value
        };
        appData[type].push(nRecord);
        DB.save(appData);
        if(currentView === type) renderTransactionView(type);
    });
}

window.delDeuda = function(type, id) {
    if(confirm("¿Seguro de borrar esta deuda?")) {
        appData[type] = appData[type].filter(x => x.id !== id);
        DB.save(appData);
        const dashView = type === 'deudasCobrar' ? 'deudas-cobrar' : 'deudas-pagar';
        if(currentView === dashView) renderDeudasView(type);
    }
}

window.editDeuda = function(type, id) {
    const item = appData[type].find(x => x.id === id);
    if(item) {
        document.getElementById('d-fecha').value = item.fecha;
        document.getElementById('d-persona').value = item.persona;
        document.getElementById('d-concepto').value = item.concepto;
        document.getElementById('d-monto').value = item.monto;
        // Set editing state instead of removing immediately
        editingDeuda = { type, id };
        // Ensure UI reflects edit mode if needed (optional)
    }
}

window.filterDeudas = function(type) {
    const personaInput = document.getElementById('filter-persona');
    const fechaInput = document.getElementById('filter-fecha');
    if (!personaInput || !fechaInput) return;
    
    const personaVal = personaInput.value.toLowerCase();
    const fechaVal = fechaInput.value;
    
    currentDeudaFilters[type].persona = personaInput.value;
    currentDeudaFilters[type].fecha = fechaVal;
    
    const dataList = appData[type];
    const filteredList = dataList.filter(item => {
        const matchPersona = !personaVal || item.persona.toLowerCase().includes(personaVal);
        const matchFecha = !fechaVal || item.fecha === fechaVal;
        return matchPersona && matchFecha;
    });
    
    const isCobrar = type === 'deudasCobrar';
    const colorClass = isCobrar ? "color: #ffaa00" : "color: #ff3c3c";
    
    const tbody = document.getElementById('deudas-tbody');
    if (tbody) {
        tbody.innerHTML = filteredList.map(item => `
            <tr>
                <td>${item.fecha}</td>
                <td>${item.persona}</td>
                <td>${item.concepto}</td>
                <td style="font-weight:bold; ${colorClass}">${formatMoney(item.monto)}</td>
                <td>
                    <div style="display:flex; gap:0.5rem; justify-content: center;">
                        <button title="Editar" onclick="editDeuda('${type}', '${item.id}')" style="background:transparent; border:none; cursor:pointer; color:var(--accent-cyan); font-size:1.25rem;"><i class="ph ph-pencil-simple"></i></button>
                        <button title="Registrar Pago" onclick="openPaymentModal('${type}', '${item.id}')" style="background:transparent; border:none; cursor:pointer; color:var(--success); font-size:1.25rem;"><i class="ph ph-coins"></i></button>
                        <button title="Borrar" onclick="delDeuda('${type}', '${item.id}')" style="background:transparent; border:none; cursor:pointer; color:var(--danger); font-size:1.25rem;"><i class="ph ph-trash"></i></button>
                    </div>
                </td>
            </tr>
        `).join('') || '<tr><td colspan="5" style="text-align:center; padding: 1rem;">No se encontraron registros</td></tr>';
    }
}

window.clearDeudaFilters = function(type) {
    currentDeudaFilters[type] = { persona: '', fecha: '' };
    const pInput = document.getElementById('filter-persona');
    const fInput = document.getElementById('filter-fecha');
    if (pInput) pInput.value = '';
    if (fInput) fInput.value = '';
    filterDeudas(type);
}

window.printFilteredDeudas = function(type) {
    const filters = currentDeudaFilters[type];
    const dataList = appData[type];
    const isCobrar = type === 'deudasCobrar';
    
    const personaVal = filters.persona.toLowerCase();
    const fechaVal = filters.fecha;
    
    const filteredList = dataList.filter(item => {
        const matchPersona = !personaVal || item.persona.toLowerCase().includes(personaVal);
        const matchFecha = !fechaVal || item.fecha === fechaVal;
        return matchPersona && matchFecha;
    });
    
    const titulo = isCobrar ? "DEUDAS POR COBRAR" : "DEUDAS POR PAGAR";
    const labelPersona = isCobrar ? "DEUDOR / CLIENTE" : "ACREEDOR / PROVEEDOR";
    const totalDeuda = filteredList.reduce((acc, curr) => acc + parseFloat(curr.monto), 0);
    
    let filterInfo = [];
    if (filters.persona) filterInfo.push(`Búsqueda: "${filters.persona}"`);
    if (filters.fecha) filterInfo.push(`Fecha: ${filters.fecha}`);
    const filterInfoStr = filterInfo.length > 0 ? `Filtros aplicados: ${filterInfo.join(' | ')}` : 'Sin filtros (Todos los registros)';
    
    printView.innerHTML = `
        <div style="padding: 2cm; font-family: 'Inter', sans-serif; color: #000; background: #fff;">
            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 2px dashed #000; padding-bottom:10px; margin-bottom:20px;">
                <div style="text-align:left;">
                    <img src="logo.png" alt="Taller El Chino" style="max-width: 500px; max-height: 120px; object-fit: contain;">
                </div>
                <div style="text-align: right;">
                    <h3 style="margin: 0; font-size: 16px;">${titulo}</h3>
                    <p style="margin: 5px 0 0 0; font-size:11px; color:#777;">Fecha Impresión: ${new Date().toLocaleDateString()}</p>
                </div>
            </div>
            
            <div style="background:#f4f4f4; padding:10px; border-radius:5px; margin-bottom:20px; font-size:12px; border:1px solid #ddd;">
                <strong>Detalle del Filtro:</strong> ${filterInfoStr}<br>
                <strong>Total de Registros Encontrados:</strong> ${filteredList.length}
            </div>
            
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size:12px;">
                <thead>
                    <tr style="background-color: #f2f2f2; border-bottom:2px solid #000;">
                        <th style="padding: 8px; text-align: left; border: 1px solid #ddd; width:40px;">N°</th>
                        <th style="padding: 8px; text-align: left; border: 1px solid #ddd; width:100px;">FECHA</th>
                        <th style="padding: 8px; text-align: left; border: 1px solid #ddd;">${labelPersona}</th>
                        <th style="padding: 8px; text-align: left; border: 1px solid #ddd;">CONCEPTO / DETALLE</th>
                        <th style="padding: 8px; text-align: right; border: 1px solid #ddd; width:120px;">MONTO</th>
                    </tr>
                </thead>
                <tbody>
                    ${filteredList.map((item, idx) => `
                        <tr style="border-bottom: 1px solid #ddd;">
                            <td style="padding: 8px; border: 1px solid #ddd;">${idx + 1}</td>
                            <td style="padding: 8px; border: 1px solid #ddd;">${item.fecha}</td>
                            <td style="padding: 8px; border: 1px solid #ddd; font-weight:600;">${item.persona}</td>
                            <td style="padding: 8px; border: 1px solid #ddd;">${item.concepto}</td>
                            <td style="padding: 8px; border: 1px solid #ddd; text-align: right; font-weight:bold;">${formatMoney(item.monto)}</td>
                        </tr>
                    `).join('') || '<tr><td colspan="5" style="text-align:center; padding:15px; border: 1px solid #ddd;">No se encontraron deudas con los criterios especificados.</td></tr>'}
                </tbody>
            </table>
            
            <div style="display:flex; justify-content:flex-end;">
                <table style="width: 250px; font-size:14px; text-align:right; border-collapse:collapse;">
                    <tr>
                        <td style="padding: 5px; font-weight:bold; border-top:2px solid #000;">TOTAL DEUDA:</td>
                        <td style="padding: 5px; font-weight:bold; border-top:2px solid #000; width:120px;">${formatMoney(totalDeuda)}</td>
                    </tr>
                </table>
            </div>
        </div>
    `;
    
    setTimeout(() => {
        window.print();
    }, 200);
}

window.openPaymentModal = function(type, id) {
    const item = appData[type].find(x => x.id === id);
    if (!item) return;
    
    const isCobrar = type === 'deudasCobrar';
    const actionLabel = isCobrar ? "Cobrar (Ingreso)" : "Pagar (Salida)";
    const colorTheme = isCobrar ? "var(--warning-text)" : "var(--danger-text)";
    
    const modalOverlay = document.createElement('div');
    modalOverlay.id = 'payment-modal';
    modalOverlay.style.position = 'fixed';
    modalOverlay.style.top = '0';
    modalOverlay.style.left = '0';
    modalOverlay.style.width = '100vw';
    modalOverlay.style.height = '100vh';
    modalOverlay.style.backgroundColor = 'rgba(0,0,0,0.6)';
    modalOverlay.style.backdropFilter = 'blur(8px)';
    modalOverlay.style.display = 'flex';
    modalOverlay.style.alignItems = 'center';
    modalOverlay.style.justifyContent = 'center';
    modalOverlay.style.zIndex = '1000';
    modalOverlay.className = 'fade-in';
    
    const maxVal = parseFloat(item.monto).toFixed(2);
    
    modalOverlay.innerHTML = `
        <div class="glass" style="width: 100%; max-width: 400px; padding: 2rem; background: var(--bg-glass); border: 1px solid var(--border-glass); border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.5); color: var(--text-main);">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:1.5rem;">
                <h3 style="margin:0; font-family:'Outfit', sans-serif; color:${colorTheme}">Registrar Pago</h3>
                <button onclick="closePaymentModal()" style="background:transparent; border:none; color:var(--text-muted); cursor:pointer; font-size:1.5rem;"><i class="ph ph-x"></i></button>
            </div>
            
            <p style="font-size:0.9rem; color:var(--text-muted); margin-bottom: 1.5rem;">
                Deuda pendiente: <strong style="color:var(--text-main)">${formatMoney(item.monto)}</strong> de <span style="font-weight:600; color:var(--text-main)">${item.persona}</span> (${item.concepto}).
            </p>
            
            <form id="payment-modal-form">
                <div class="form-group" style="margin-bottom:1.2rem;">
                    <label style="display:block; margin-bottom:0.5rem; font-size:0.85rem; color:var(--text-muted)">Monto a pagar (S/)</label>
                    <input type="number" step="0.01" id="pay-monto" class="form-input" required min="0.01" max="${maxVal}" value="${maxVal}" style="font-size:1.1rem; font-weight:700;">
                </div>
                <div class="form-group" style="margin-bottom:1.2rem;">
                    <label style="display:block; margin-bottom:0.5rem; font-size:0.85rem; color:var(--text-muted)">Fecha de Pago</label>
                    <input type="date" id="pay-fecha" class="form-input" required value="${new Date().toISOString().split('T')[0]}">
                </div>
                <div class="form-group" style="margin-bottom:1.5rem;">
                    <label style="display:block; margin-bottom:0.5rem; font-size:0.85rem; color:var(--text-muted)">Concepto</label>
                    <input type="text" id="pay-concepto" class="form-input" required value="Pago de deuda: ${item.persona} - ${item.concepto}">
                </div>
                
                <button type="submit" class="btn-primary" style="background: linear-gradient(135deg, ${isCobrar ? 'var(--accent-cyan), var(--accent-blue)' : 'var(--accent-pink), var(--accent-orange)'}); color: ${isCobrar ? '#000' : '#fff'};">
                    Confirmar ${actionLabel}
                </button>
            </form>
        </div>
    `;
    
    document.body.appendChild(modalOverlay);
    
    document.getElementById('payment-modal-form').addEventListener('submit', (e) => {
        e.preventDefault();
        
        const payMonto = parseFloat(document.getElementById('pay-monto').value);
        const payFecha = document.getElementById('pay-fecha').value;
        const payConcepto = document.getElementById('pay-concepto').value;
        
        if (isNaN(payMonto) || payMonto <= 0) return alert("Ingrese un monto válido.");
        if (payMonto > parseFloat(item.monto)) return alert("El monto no puede superar la deuda pendiente.");
        
        const targetType = isCobrar ? 'entradas' : 'salidas';
        const nRecord = {
            id: generateId(),
            fecha: payFecha,
            concepto: payConcepto,
            monto: payMonto
        };
        appData[targetType].push(nRecord);
        
        const originalMonto = parseFloat(item.monto);
        const remainingMonto = originalMonto - payMonto;
        
        if (remainingMonto <= 0.005) {
            appData[type] = appData[type].filter(x => x.id !== id);
            alert("¡Deuda cancelada en su totalidad!");
        } else {
            item.monto = remainingMonto.toFixed(2);
            alert(`Pago registrado. Deuda restante: ${formatMoney(item.monto)}`);
        }
        
        DB.save(appData);
        closePaymentModal();
        renderDeudasView(type);
    });
}

window.closePaymentModal = function() {
    const modal = document.getElementById('payment-modal');
    if (modal) modal.remove();
}

function renderDeudasView(type) {
    // Reset editing state whenever the view is rendered to avoid stale edit references
    editingDeuda = { type: null, id: null };
    const isCobrar = type === 'deudasCobrar';
    const titulo = isCobrar ? "Deudas por Cobrar" : "Deudas por Pagar";
    const labelPersona = isCobrar ? "Cliente / Deudor" : "Acreedor / Proveedor";
    const colorClass = isCobrar ? "color: #ffaa00" : "color: #ff3c3c";
    
    const filters = currentDeudaFilters[type] || { persona: '', fecha: '' };
    
    viewsContainer.innerHTML = `
        <div class="layout-split fade-in">
            <div class="glass-card">
                <h3 class="gradient-text" style="margin-bottom: 1rem; ${colorClass}">Registrar ${titulo}</h3>
                <form id="form-${type}">
                    <div class="form-group">
                        <label>Fecha</label>
                        <input type="date" id="d-fecha" class="form-input" required value="${new Date().toISOString().split('T')[0]}">
                    </div>
                    <div class="form-group">
                        <label>${labelPersona}</label>
                        <input type="text" id="d-persona" class="form-input" required placeholder="Nombre...">
                    </div>
                    <div class="form-group">
                        <label>Concepto / Detalle</label>
                        <input type="text" id="d-concepto" class="form-input" required placeholder="Por qué motivo...">
                    </div>
                    <div class="form-group">
                        <label>Monto</label>
                        <input type="number" step="0.01" id="d-monto" class="form-input" required placeholder="0.00">
                    </div>
                    <button type="submit" class="btn-primary" style="${isCobrar ? 'background:rgba(255,170,0,0.2);border-color:#ffaa00;color:#ffaa00;' : 'background:rgba(255,60,60,0.2);border-color:#ff3c3c;color:#ff3c3c;'}">Guardar Deuda</button>
                </form>
            </div>
            
            <div class="glass-card">
                <h3 class="gradient-text" style="margin-bottom: 1rem; ${colorClass}">Historial de ${titulo}</h3>
                
                <div style="display: flex; gap: 0.8rem; margin-bottom: 1.2rem; flex-wrap: wrap; background: rgba(0,0,0,0.15); padding: 0.8rem; border-radius: 8px; border: 1px solid var(--border-glass)">
                    <div style="flex: 1; min-width: 140px;">
                        <label style="font-size: 0.8rem; color: var(--text-muted); display: block; margin-bottom: 0.2rem;">${isCobrar ? 'Deudor' : 'Acreedor'}</label>
                        <input type="text" id="filter-persona" class="form-input" placeholder="Buscar..." value="${filters.persona}" oninput="filterDeudas('${type}')" style="padding: 0.4rem 0.8rem; font-size: 0.9rem;">
                    </div>
                    <div style="width: 140px;">
                        <label style="font-size: 0.8rem; color: var(--text-muted); display: block; margin-bottom: 0.2rem;">Fecha</label>
                        <input type="date" id="filter-fecha" class="form-input" value="${filters.fecha}" onchange="filterDeudas('${type}')" style="padding: 0.4rem 0.8rem; font-size: 0.9rem;">
                    </div>
                    <div style="display: flex; align-items: flex-end; gap: 0.4rem;">
                        <button class="btn-secondary" onclick="clearDeudaFilters('${type}')" style="padding: 0.4rem 0.8rem; font-size: 0.9rem;" title="Limpiar Filtros">
                            <i class="ph ph-broom"></i>
                        </button>
                        <button class="btn-secondary" onclick="printFilteredDeudas('${type}')" style="padding: 0.4rem 0.8rem; font-size: 0.9rem; border-color: var(--accent-cyan); color: var(--accent-cyan);" title="Imprimir Reporte Filtrado">
                            <i class="ph ph-printer"></i> Imprimir
                        </button>
                    </div>
                </div>

                <table class="table-glass">
                    <thead><tr><th>Fecha</th><th>${isCobrar ? 'Deudor' : 'Acreedor'}</th><th>Concepto</th><th>Monto</th><th style="width:110px">Acciones</th></tr></thead>
                    <tbody id="deudas-tbody">
                        <!-- Se inyecta por JS -->
                    </tbody>
                </table>
            </div>
        </div>
    `;

    document.getElementById(`form-${type}`).addEventListener('submit', (e) => {
    e.preventDefault();
    const fecha = document.getElementById('d-fecha').value;
    const persona = document.getElementById('d-persona').value;
    const concepto = document.getElementById('d-concepto').value;
    const monto = document.getElementById('d-monto').value;

    if(editingDeuda.type === type && editingDeuda.id) {
        // Update existing record
        const idx = appData[type].findIndex(d => d.id === editingDeuda.id);
        if(idx >= 0) {
            appData[type][idx] = { ...appData[type][idx], fecha, persona, concepto, monto };
        } else {
            // Fallback: push as new if not found
            appData[type].push({ id: editingDeuda.id, fecha, persona, concepto, monto });
        }
        // Clear editing state
        editingDeuda = { type: null, id: null };
    } else {
        // New record
        const nRecord = {
            id: generateId(),
            fecha,
            persona,
            concepto,
            monto
        };
        appData[type].push(nRecord);
    }
    DB.save(appData);
    const dashView = type === 'deudasCobrar' ? 'deudas-cobrar' : 'deudas-pagar';
    if(currentView === dashView) renderDeudasView(type);
});

    filterDeudas(type);
}

// --- VISTA: TRABAJADORES ---
window.delWorker = function(id) {
    if(confirm("¿Eliminar trabajador?")) {
        appData.trabajadores = appData.trabajadores.filter(w => w.id !== id);
        DB.save(appData);
        renderTrabajadores();
    }
}
window.editWorker = function(id) {
    const ww = appData.trabajadores.find(w => w.id === id);
    if(ww) {
        document.getElementById('w-nombre').value = ww.nombre;
        document.getElementById('w-cargo').value = ww.cargo;
        // Remueve original para guardar como nueva edicion al enviar
        appData.trabajadores = appData.trabajadores.filter(w => w.id !== id);
    }
}

function renderTrabajadores() {
    viewsContainer.innerHTML = `
        <div class="layout-split fade-in">
            <div class="glass-card">
                <h3 class="gradient-text" style="margin-bottom: 1rem">Añadir / Editar Trabajador</h3>
                <form id="form-trabajador">
                    <div class="form-group">
                        <label>Nombre Completo</label>
                        <input type="text" id="w-nombre" class="form-input" required>
                    </div>
                    <div class="form-group">
                        <label>Cargo / Rol</label>
                        <input type="text" id="w-cargo" class="form-input" required>
                    </div>
                    <button type="submit" class="btn-primary">Guardar Trabajador</button>
                </form>
            </div>
            <div class="glass-card">
                <h3 class="gradient-text" style="margin-bottom: 1rem">Plantilla Activa</h3>
                <table class="table-glass">
                    <thead><tr><th>Nombre</th><th>Cargo</th><th style="width:80px">Acciones</th></tr></thead>
                    <tbody>
                        ${appData.trabajadores.map(w => `
                            <tr>
                                <td>${w.nombre}</td>
                                <td>${w.cargo}</td>
                                <td>
                                    <div style="display:flex; gap:0.5rem">
                                        <button title="Editar" onclick="editWorker('${w.id}')" style="background:transparent; border:none; cursor:pointer; color:var(--accent-cyan); font-size:1.2rem;"><i class="ph ph-pencil-simple"></i></button>
                                        <button title="Borrar" onclick="delWorker('${w.id}')" style="background:transparent; border:none; cursor:pointer; color:var(--danger); font-size:1.2rem;"><i class="ph ph-trash"></i></button>
                                    </div>
                                </td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;
    
    document.getElementById('form-trabajador')?.addEventListener('submit', (e) => {
        e.preventDefault();
        appData.trabajadores.push({
            id: generateId(),
            nombre: document.getElementById('w-nombre').value,
            cargo: document.getElementById('w-cargo').value
        });
        DB.save(appData);
        renderTrabajadores();
    });
}

// --- VISTA: PROFORMAS ---
const BANCOS_PROFORMA = [
    { banco: "BANCO DE CRÉDITO DEL PERÚ (BCP)-CTA Personal", moneda: "SOLES", cuenta: "47504669868095", cci: "00247510466986809528" },
    { banco: "BANCO DE CRÉDITO DEL PERÚ (BCP)-CTA RUC 10", moneda: "SOLES", cuenta: "47506575901082", cci: "00247510657590108222" },
    { banco: "BANCO DE CRÉDITO DEL PERÚ (BCP)-CTA RUC 20", moneda: "SOLES", cuenta: "4757233070091", cci: "00247500723307009124" },
];

let currentProforma = { manoObra: [], repuestos: [] };

function renderProformas() {
    viewsContainer.innerHTML = `
        <div class="layout-split fade-in" style="grid-template-columns: 1fr;">
            <div class="glass-card">
                <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap: wrap; gap: 1rem;">
                    <h3 class="gradient-text">Generador de Proforma "El Chino"</h3>
                    <button class="btn-secondary" id="btn-toggle-history" style="width:auto; display:flex; align-items:center; gap:0.5rem;">
                        <i class="ph ph-clock-counter-clockwise"></i> Ver Historial de Proformas
                    </button>
                </div>
                
                <div style="margin-top: 1rem; margin-bottom: 0.5rem;">
                    <div class="form-group" style="margin-bottom: 0; max-width: 250px;">
                        <label style="display: block; margin-bottom: 0.5rem; color: var(--text-muted); font-size: 0.9rem; font-weight: 600;">Número de Proforma</label>
                        <input type="text" id="prof-nro" class="form-input" value="${String(appData.proformas.length + 1).padStart(4, '0')}" style="font-weight: 700; font-size: 1.1rem; border-color: var(--accent-cyan);">
                    </div>
                </div>
                
                <div class="proforma-builder" style="margin-top: 1.5rem;">
                    <h4 style="color:var(--accent-cyan); margin-bottom: 0.5rem;">Datos del Cliente (Editables)</h4>
                    <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:1rem;">
                        <div class="form-group"><label>Señor(a) / Cliente</label><input type="text" id="prof-cliente" class="form-input" placeholder="CLIENTE NO DEFINIDO"></div>
                        <div class="form-group"><label>RUC / DNI</label><input type="text" id="prof-ruc" class="form-input" placeholder="0"></div>
                        <div class="form-group"><label>Placa del Vehículo</label><input type="text" id="prof-placa" class="form-input" placeholder="---"></div>
                        <div class="form-group"><label>Correo Electrónico</label><input type="email" id="prof-correo" class="form-input" placeholder="---"></div>
                        <div class="form-group"><label>Dirección</label><input type="text" id="prof-direccion" class="form-input" placeholder="---"></div>
                        <div class="form-group"><label>Celular</label><input type="text" id="prof-celular" class="form-input" placeholder="---"></div>
                        <div class="form-group"><label>Fecha de Emisión</label><input type="date" id="prof-fecha" class="form-input" value="${new Date().toISOString().split('T')[0]}"></div>
                        <div class="form-group"><label>Válido Hasta</label><input type="text" id="prof-validez" class="form-input" value="10 días"></div>
                        <div class="form-group"><label>Moneda</label><input type="text" id="prof-moneda" class="form-input" value="SOLES"></div>
                    </div>
                    
                    <hr style="border-color: var(--border-glass)">
                    <h4>Añadir Mano de Obra</h4>
                    <div class="proforma-item-row">
                        <input type="text" id="mo-desc" class="form-input" placeholder="Descripción (Ej: Cambio pastillas)">
                        <input type="number" id="mo-cant" class="form-input" placeholder="Cant." value="1">
                        <input type="number" id="mo-precio" class="form-input" placeholder="Precio Unit. S/" step="0.01">
                        <input type="number" id="mo-desc-pct" class="form-input" placeholder="Descto." value="0">
                        <button class="btn-primary" id="btn-add-mo" title="Añadir a Mano de Obra">+</button>
                    </div>

                    <h4>Añadir Repuestos</h4>
                    <div class="proforma-item-row">
                        <input type="text" id="rep-desc" class="form-input" placeholder="Descripción Producto">
                        <input type="number" id="rep-cant" class="form-input" placeholder="Cant." value="1">
                        <input type="number" id="rep-precio" class="form-input" placeholder="Precio Unit. S/" step="0.01">
                        <input type="number" id="rep-desc-pct" class="form-input" placeholder="Descto." value="0">
                        <button class="btn-primary" id="btn-add-rep" title="Añadir a Repuestos">+</button>
                    </div>
                    
                    <hr style="border-color: var(--border-glass); margin: 1.5rem 0;">
                    
                    <h4 style="color:var(--accent-cyan); margin-bottom: 0.5rem;">Previsualización de Ítems Actuales</h4>
                    <div id="preview-items-container" style="margin-bottom: 1.5rem;">
                        <!-- Tabla previsualización inyectada desde JS -->
                    </div>

                    <div class="form-group" style="margin-top: 1rem;">
                        <label>Observaciones (Para imprimir en el documento)</label>
                        <textarea id="prof-observaciones" class="form-input" placeholder="Escribe aquí consideraciones, diagnósticos o notas para el cliente..." rows="3"></textarea>
                    </div>

                    <button id="btn-save-print" class="btn-secondary" style="margin-top:2rem; width:100%; border-color:var(--accent-cyan); color:var(--accent-cyan);">
                        <i class="ph ph-device-floppy"></i> Guardar Nueva y Previsualizar
                    </button>
                </div>
            </div>
            
            <div class="glass-card hidden" id="proformas-history-panel" style="margin-top: 1.5rem;">
                <h3 class="gradient-text" style="margin-bottom: 1rem">Historial de Proformas Emitidas</h3>
                <table class="table-glass">
                    <thead><tr><th>Cod.</th><th>Emisión</th><th>Cliente</th><th>Total Importe</th><th style="width:80px">Acción</th></tr></thead>
                    <tbody>
                        ${[...appData.proformas].reverse().map(p => `
                            <tr>
                                <td>N° ${p.nroCotizacion || '---'}</td>
                                <td>${p.fecha}</td>
                                <td>${p.cliente}</td>
                                <td class="stat-value income" style="font-size:1rem">${formatMoney(p.totalAPagar)}</td>
                                <td>
                                    <div style="display:flex; gap:0.5rem">
                                        <button title="Imprimir" onclick="reprintProforma('${p.id}')" style="background:var(--accent-blue); color:#000; border:none; padding:0.3rem 0.5rem; border-radius:4px; font-weight:600; cursor:pointer;"><i class="ph ph-printer"></i></button>
                                        <button title="Editar" onclick="editProforma('${p.id}')" style="background:transparent; border:none; cursor:pointer; color:var(--success); font-size:1.2rem;"><i class="ph ph-pencil-simple"></i></button>
                                        <button title="Borrar" onclick="delProforma('${p.id}')" style="background:transparent; border:none; cursor:pointer; color:var(--danger); font-size:1.2rem;"><i class="ph ph-trash"></i></button>
                                    </div>
                                </td>
                            </tr>
                        `).join('') || '<tr><td colspan="5" style="text-align:center">No hay registros almacenados</td></tr>'}
                    </tbody>
                </table>
            </div>
        </div>
    `;

    document.getElementById('btn-add-mo').addEventListener('click', () => addItemToProforma('manoObra'));
    document.getElementById('btn-add-rep').addEventListener('click', () => addItemToProforma('repuestos'));
    document.getElementById('btn-toggle-history').addEventListener('click', () => document.getElementById('proformas-history-panel').classList.toggle('hidden'));
    window.renderItemsPreviewTable();

    // Crear y guardar
    document.getElementById('btn-save-print').addEventListener('click', () => {
        const pData = buildProformaObject();
        pData.nroCotizacion = document.getElementById('prof-nro').value || String(appData.proformas.length + 1).padStart(4, '0');
        pData.id = generateId();
        appData.proformas.push(pData);
        DB.save(appData);
        launchPrintView(pData);
        
        // Reset local currentProforma
        currentProforma = { manoObra: [], repuestos: [] };
        renderProformas();
    });
}

function addItemToProforma(type) {
    const prefix = type === 'manoObra' ? 'mo' : 'rep';
    const desc = document.getElementById(`${prefix}-desc`).value;
    const cant = document.getElementById(`${prefix}-cant`).value;
    const precio = document.getElementById(`${prefix}-precio`).value;
    const descPct = document.getElementById(`${prefix}-desc-pct`).value || 0;

    if(!desc || !precio) return alert("Llena descripción y precio.");

    const importe = (cant * precio) - descPct;
    currentProforma[type].push({ desc, cant, precio: parseFloat(precio), descuento: parseFloat(descPct), importe });
    
    document.getElementById(`${prefix}-desc`).value = '';
    document.getElementById(`${prefix}-precio`).value = '';
    window.renderItemsPreviewTable();
}

function buildProformaObject() {
    return {
        cliente: document.getElementById('prof-cliente').value || "CLIENTE NO DEFINIDO",
        ruc: document.getElementById('prof-ruc').value || "0",
        placa: document.getElementById('prof-placa').value || "---",
        correo: document.getElementById('prof-correo').value || "---",
        direccion: document.getElementById('prof-direccion').value || "---",
        celular: document.getElementById('prof-celular').value || "---",
        fecha: document.getElementById('prof-fecha').value || new Date().toISOString().split('T')[0],
        validez: document.getElementById('prof-validez').value || "---",
        moneda: document.getElementById('prof-moneda').value || "SOLES",
        observaciones: document.getElementById('prof-observaciones').value || "",
        manoObra: JSON.parse(JSON.stringify(currentProforma.manoObra)),
        repuestos: JSON.parse(JSON.stringify(currentProforma.repuestos))
    };
}

window.renderItemsPreviewTable = function() {
    const container = document.getElementById('preview-items-container');
    if(!container) return;
    
    const items = [
        ...currentProforma.manoObra.map((x, i) => ({...x, _rIndex: i, _type: 'manoObra'})),
        ...currentProforma.repuestos.map((x, i) => ({...x, _rIndex: i, _type: 'repuestos'}))
    ];
    
    if(items.length === 0) {
        container.innerHTML = "<p style='color:var(--text-muted); font-size:0.9rem; font-style:italic;'>Aún no hay ítems en esta proforma.</p>";
        return;
    }
    
    container.innerHTML = `
        <table class="table-glass" style="margin-top:0;">
            <thead>
                <tr>
                    <th>Tipo</th>
                    <th>Descripción</th>
                    <th>Cant.</th>
                    <th>Precio</th>
                    <th>Descuento</th>
                    <th>Importe</th>
                    <th style="width:80px">Acción</th>
                </tr>
            </thead>
            <tbody>
                ${items.map(it => `
                    <tr>
                        <td><span class="badge ${it._type==='manoObra' ? 'badge-income' : 'badge-expense'}" style="font-size:0.7rem">${it._type==='manoObra' ? 'Mano de Obra' : 'Repuesto'}</span></td>
                        <td>${it.desc}</td>
                        <td>${it.cant}</td>
                        <td>S/ ${it.precio.toFixed(2)}</td>
                        <td>${it.descuento || 0}</td>
                        <td style="color:var(--accent-cyan)">S/ ${it.importe.toFixed(2)}</td>
                        <td>
                            <div style="display:flex; gap:0.5rem">
                                <button title="Editar" onclick="editProformaItem('${it._type}', ${it._rIndex})" style="background:transparent; border:none; cursor:pointer; color:var(--success); font-size:1.2rem;"><i class="ph ph-pencil-simple"></i></button>
                                <button title="Borrar" onclick="delProformaItem('${it._type}', ${it._rIndex})" style="background:transparent; border:none; cursor:pointer; color:var(--danger); font-size:1.2rem;"><i class="ph ph-trash"></i></button>
                            </div>
                        </td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;
};

window.delProformaItem = function(type, index) {
    if(confirm("¿Quitar este ítem de la proforma?")) {
        currentProforma[type].splice(index, 1);
        window.renderItemsPreviewTable();
    }
}

window.editProformaItem = function(type, index) {
    const item = currentProforma[type][index];
    const prefix = type === 'manoObra' ? 'mo' : 'rep';
    document.getElementById(`${prefix}-desc`).value = item.desc;
    document.getElementById(`${prefix}-cant`).value = item.cant;
    document.getElementById(`${prefix}-precio`).value = item.precio;
    document.getElementById(`${prefix}-desc-pct`).value = item.descuento || 0;
    
    currentProforma[type].splice(index, 1);
    window.renderItemsPreviewTable();
}

window.reprintProforma = function(id) {
    const p = appData.proformas.find(x => x.id === id);
    if(p) {
        if(!p.manoObra) {
            alert("Esta es una proforma muy antigua y no contiene detalles internos guardados, solo metadata.");
            return;
        }
        launchPrintView(p);
    }
}

function launchPrintView(data) {
    const signo = data.moneda.toUpperCase().includes("DOL") ? "$" : "S/";

    let htmlManoObra = '';
    let subtotalMO = 0;
    (data.manoObra || []).forEach((i, idx) => {
        htmlManoObra += `<tr><td>${idx+1}</td><td>${i.desc}</td><td>${i.cant}</td><td>${signo} ${i.precio?.toFixed(2)||0}</td><td>${i.descuento>0?i.descuento:''}</td><td class="text-right">${signo} ${i.importe?.toFixed(2)||0}</td></tr>`;
        subtotalMO += parseFloat(i.importe || 0);
    });

    let htmlRepuestos = '';
    let subtotalRep = 0;
    (data.repuestos || []).forEach((i, idx) => {
        htmlRepuestos += `<tr><td>${idx+1}</td><td>${i.desc}</td><td>${i.cant}</td><td>UND</td><td>${signo} ${i.precio?.toFixed(2)||0}</td><td>${i.descuento>0?i.descuento:''}</td><td class="text-right">${signo} ${i.importe?.toFixed(2)||0}</td></tr>`;
        subtotalRep += parseFloat(i.importe || 0);
    });

    data.totalAPagar = subtotalMO + subtotalRep;
    const igv = data.totalAPagar * 0.18;
    const subtotalBruto = data.totalAPagar - igv;

    // Emulación del Logo Taller Mécanico "El Chino"
    const logoHtml = `
        <div style="display:flex; align-items:center;">
            <div style="margin-right:15px; text-align:center;">
                <!-- Simplemente coloca la imagen en la misma carpeta del proyecto bajo el nombre logo.png o logo.jpg -->
                <img src="logo.png" alt="" style="max-width: 500px; max-height: 120px; object-fit: contain;">
            </div>
        </div>
    `;

    printView.innerHTML = `
        <div class="proforma-header" style="border-bottom: 2px dashed #ff0000;">
            ${logoHtml}
            <div class="proforma-ruc" style="border-radius:15px; min-width: 250px;">
                <h3 style="font-size:18px">RUC: 10463232763</h3>
                <h3 style="margin-top:10px;">COTIZACIÓN N° ${data.nroCotizacion}</h3>
            </div>
        </div>
        
        <div class="proforma-client-info">
            <div><strong>Sr. (a):</strong> &nbsp;&nbsp;&nbsp;${data.cliente}</div>
            <div style="text-align:right"><strong>RUC:</strong> &nbsp;&nbsp;&nbsp;${data.ruc}</div>
            <div><strong>PLACA:</strong> &nbsp;&nbsp;&nbsp;${data.placa}</div>
            <div style="text-align:right"><strong>CORREO:</strong> &nbsp;&nbsp;&nbsp;${data.correo}</div>
            <div><strong>DIRECCIÓN:</strong> &nbsp;&nbsp;&nbsp;${data.direccion}</div>
            <div style="text-align:right"><strong>FECHA EMISIÓN:</strong> &nbsp;&nbsp;&nbsp;${data.fecha}</div>
            <div><strong>CELULAR:</strong> &nbsp;&nbsp;&nbsp;${data.celular}</div>
            <div style="text-align:right"><strong>VÁLIDO HASTA:</strong> &nbsp;&nbsp;&nbsp;${data.validez}</div>
            <div><strong>MONEDA:</strong> &nbsp;&nbsp;&nbsp;${data.moneda}</div>
        </div>

        <h4 style="margin:5px 0;">MANO DE OBRA</h4>
        <table class="pro-table">
            <thead>
                <tr><th style="width:50px">ITEM</th><th>DESCRIPCIÓN SERVICIO</th><th style="width:60px">CANT.</th><th style="width:100px">PRECIO<br>UNITARIO</th><th style="width:100px">DESCUENTO</th><th style="width:120px">IMPORTE<br>TOTAL</th></tr>
            </thead>
            <tbody>${htmlManoObra || '<tr><td colspan="6" style="text-align:center">Sin items</td></tr>'}</tbody>
        </table>

        <h4 style="margin:5px 0;">REPUESTOS</h4>
        <table class="pro-table">
            <thead>
                <tr><th style="width:50px">ITEM</th><th>DESCRIPCIÓN DEL PRODUCTO</th><th style="width:60px">CANT.</th><th style="width:50px">UND</th><th style="width:100px">PRECIO<br>UNITARIO</th><th style="width:100px">DESCUENTO</th><th style="width:120px">IMPORTE<br>TOTAL</th></tr>
            </thead>
            <tbody>${htmlRepuestos || '<tr><td colspan="7" style="text-align:center">Sin items</td></tr>'}</tbody>
        </table>

        <div style="display:flex; justify-content:space-between; margin-top:10px;">
            <div style="border:1px solid #000; padding:10px; width:58%;">
                <strong>Observaciones:</strong><br>${data.observaciones ? String(data.observaciones).replace(/\n/g, '<br>') : ''}<br>
            </div>
            <div class="pro-totals">
                <table style="width:100%; text-align:right; border-collapse:collapse; font-size:13px;">
                    <tr><td style="padding:3px;">Sub Total</td><td style="width:50px">${signo}</td><td>${subtotalBruto.toFixed(2)}</td></tr>
                    <tr><td style="padding:3px;">I.G.V</td><td>${signo}</td><td>${igv.toFixed(2)}</td></tr>
                    <tr><td style="padding:3px;">Descuento</td><td>${signo}</td><td>-</td></tr>
                    <tr><td style="padding:3px; border-top:1px solid #000; font-size:14px;"><strong>Total a Pagar</strong></td><td style="border-top:1px solid #000; font-size:14px;"><strong>${signo}</strong></td><td style="border-top:1px solid #000; font-size:14px;"><strong>${data.totalAPagar.toFixed(2)}</strong></td></tr>
                </table>
            </div>
        </div>

        <div style="margin-top:20px; font-size:10px; border:1px solid #000; padding:10px;">
            <strong style="font-size:11px;">Términos y condiciones</strong><br>
            1. Validez: La presente cotización tendrá una vigencia de 10 días calendario, desde su emisión.<br>
            2. Precios: Sujetos a cambio sin previo aviso según disponibilidad y mercado.<br>
            3. Pagos: Se realizan a las cuentas bancarias indicadas tras confirmar el pedido.<br>
            4. Mano de obra y repuestos: Puede variar según diagnóstico y condiciones del servicio.
        </div>

        <table class="pro-table bank-accounts">
            <thead>
                <tr><th>BANCO</th><th>MONEDA</th><th>NRO. CUENTA</th><th>NRO. CUENTA CCI</th></tr>
            </thead>
            <tbody>
                ${BANCOS_PROFORMA.map(b => `
                    <tr>
                        <td style="text-align:center">${b.banco}</td>
                        <td style="text-align:center">${b.moneda}</td>
                        <td style="text-align:center">${b.cuenta}</td>
                        <td style="text-align:center">${b.cci}</td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;

    setTimeout(() => {
        window.print();
    }, 200);
}

window.delProforma = function(id) {
    if(confirm("¿Seguro de borrar esta proforma?")) {
        appData.proformas = appData.proformas.filter(p => p.id !== id);
        DB.save(appData);
        renderProformas();
    }
}

window.editProforma = function(id) {
    const p = appData.proformas.find(x => x.id === id);
    if(p) {
        if(!p.manoObra) {
            alert("Esta es una proforma muy antigua y no contiene detalles internos guardados, solo metadata.");
            return;
        }
        document.getElementById('prof-cliente').value = p.cliente || "";
        document.getElementById('prof-ruc').value = p.ruc || "";
        document.getElementById('prof-placa').value = p.placa || "";
        document.getElementById('prof-correo').value = p.correo || "";
        document.getElementById('prof-direccion').value = p.direccion || "";
        document.getElementById('prof-celular').value = p.celular || "";
        document.getElementById('prof-fecha').value = p.fecha || "";
        document.getElementById('prof-validez').value = p.validez || "";
        document.getElementById('prof-moneda').value = p.moneda || "SOLES";
        document.getElementById('prof-observaciones').value = p.observaciones || "";
        document.getElementById('prof-nro').value = p.nroCotizacion || "";
        
        currentProforma.manoObra = JSON.parse(JSON.stringify(p.manoObra));
        currentProforma.repuestos = JSON.parse(JSON.stringify(p.repuestos));
        
        appData.proformas = appData.proformas.filter(x => x.id !== id);
        document.getElementById('proformas-history-panel').classList.add('hidden');
        DB.save(appData); // Fix offset when pushing back later
        window.renderItemsPreviewTable();
    }
}

// Inicializar la App
setView('dashboard');
