function toggleUtilities() {
  const menu = document.getElementById('utilitiesMenu');
  menu.classList.toggle('show');
}

document.addEventListener('click', (e) => {
  const dropdown = document.querySelector('.utilities-dropdown');
  const menu = document.getElementById('utilitiesMenu');
  if (dropdown && !dropdown.contains(e.target)) {
    menu.classList.remove('show');
  }
});

function openPlanillasNewTab() {
  document.getElementById('utilitiesMenu').classList.remove('show');
  window.open('about:blank', '_blank');
  toast('📋 Abre la funcionalidad de Planillas en otra pestaña');
}


// Filtros MEJORADOS con ordenamiento
const Filters = {
  catec: false,
  ftth: false,
  excludeFTTH: false,
  nodoEstado: '',
  cmts: '',
  days: 3,
  territorio: '',
  sistema: '',
  alarma: '',
  quickSearch: '',
  showAllStates: false,
  ordenarPorIngreso: 'desc', // 'desc', 'asc', o ''
  equipoModelo: '',
  equipoMarca: '',

  apply(rows) {
    const queries = this.quickSearch
      ? this.quickSearch.split(';').map(q => q.trim()).filter(Boolean)
      : [];

    return rows.filter(r => {
      const meta = r.__meta || {};

      if (meta.daysFromToday > this.days) return false;

      if (!this.showAllStates) {
        if (!meta.estado || !meta.estadoValido) return false;
        if (CONFIG.estadosOcultosPorDefecto.includes(meta.estado)) return false;
        if (!CONFIG.estadosPermitidos.includes(meta.estado)) return false;
      }

      if (this.catec && !meta.tipoTrabajo.includes('CATEC')) return false;
      if (this.ftth && !meta.esFTTH) return false;
      if (this.excludeFTTH && meta.esFTTH) return false;
      if (this.territorio && meta.territorio !== this.territorio) return false;
      if (this.sistema && meta.sistema !== this.sistema) return false;

      if (queries.length) {
        if (!meta.searchable) return false;
        if (queries.length === 1) {
          if (!TextUtils.matches(meta.searchable, queries[0])) return false;
        } else {
          if (!TextUtils.matchesMultiple(meta.searchable, queries)) return false;
        }
      }

      return true;
    });
  },

  applyToZones(zones) {
    let filtered = zones.slice();

    if (this.nodoEstado) {
      filtered = filtered.filter(z => {
        if (this.nodoEstado === 'up') return z.nodoEstado === 'up';
        if (this.nodoEstado === 'down') return z.nodoEstado === 'down' || z.nodoEstado === 'critical';
        if (this.nodoEstado === 'critical') return z.nodoEstado === 'critical';
        return true;
      });
    }

    if (this.cmts) {
      filtered = filtered.filter(z => z.cmts === this.cmts);
    }

    if (this.alarma) {
      if (this.alarma === 'con-alarma') {
        filtered = filtered.filter(z => z.tieneAlarma);
      } else if (this.alarma === 'sin-alarma') {
        filtered = filtered.filter(z => !z.tieneAlarma);
      }
    }

    // NUEVO: Aplicar ordenamiento por ingreso
    if (this.ordenarPorIngreso === 'desc') {
      filtered.sort((a, b) => b.totalOTs - a.totalOTs);
    } else if (this.ordenarPorIngreso === 'asc') {
      filtered.sort((a, b) => a.totalOTs - b.totalOTs);
    }

    return filtered;
  }
};



function applyEquiposFilters() {
  Filters.equipoModelo = document.getElementById('filterEquipoModelo').value;
  Filters.equipoMarca = document.getElementById('filterEquipoMarca').value;
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
}

function clearEquiposFilters() {
  Filters.equipoModelo = '';
  Filters.equipoMarca = '';
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
}

function toggleEquiposGrupo(zona){
  if (!window.equiposOpen) window.equiposOpen = new Set();
  if (window.equiposOpen.has(zona)) window.equiposOpen.delete(zona);
  else window.equiposOpen.add(zona);
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
}

function exportEquiposGrupoExcel(zona, useFiltered = false){
  const source = useFiltered && window.equiposPorZona ? window.equiposPorZona : window.equiposPorZonaCompleto;

  if (!source) return toast('No hay datos de equipos');

  const arr = source.get(zona)||[];
  if (!arr.length) return toast('No hay equipos en esa zona');

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(arr);
  XLSX.utils.book_append_sheet(wb, ws, `Equipos_${zona||'NA'}`);

  const filterInfo = useFiltered && (Filters.equipoModelo || Filters.equipoMarca) 
    ? `_${Filters.equipoModelo || 'TodosModelos'}_${Filters.equipoMarca || 'TodasMarcas'}`
    : '';

  XLSX.writeFile(wb, `Equipos_${zona||'NA'}${filterInfo}_${new Date().toISOString().slice(0,10)}.xlsx`);
  toast(`✅ Exportados ${arr.length} equipos de ${zona}`);
}

let currentData = null;
let allZones = [];
let allCMTS = [];
let currentZone = null;
let selectedOrders = new Set();
let baseZoneGroups = [];
let equipmentCache = new Map();
window.equipmentCache = equipmentCache;

function ensureOrderRowKey(order, index = 0, zona = '') {
  if (!order) return '';
  if (order.__rowKey) return order.__rowKey;

  const candidateFields = [
    'Número de cita',
    'Número del caso',
    'Caso Externo',
    'Caso externo',
    'External Case Id',
    'CASO EXTERNO',
    'external_case_id'
  ];

  for (const field of candidateFields) {
    const value = order[field];
    if (value !== undefined && value !== null) {
      const trimmed = String(value).trim();
      if (trimmed) {
        order.__rowKey = trimmed;
        return order.__rowKey;
      }
    }
  }

  const zonaSegment = zona ? String(zona).replace(/[^A-Za-z0-9]/g, '').slice(0, 10) : 'zona';
  order.__rowKey = `row_${zonaSegment}_${Date.now()}_${index}`;
  return order.__rowKey;
}

function setZoneModalFooter() {
  const footer = document.getElementById('modalFooter');
  if (!footer) return;

  const disabledAttr = selectedOrders.size > 0 ? '' : 'disabled';
  footer.innerHTML = `
    <div class="selection-info" id="selectionInfo">${selectedOrders.size} órdenes seleccionadas</div>
    <div class="modal-footer-actions">
      <button class="btn btn-primary" onclick="exportModalDetalleExcel()">
        📄 Exportar detalle
      </button>
      <button class="btn btn-primary" onclick="exportZoneOrdersExcel()">
        📥 Exportar Excel (Zona)
      </button>
      <button class="btn btn-warning" id="btnExportBEFAN" ${disabledAttr} onclick="exportBEFAN()">
        📤 Exportar BEFAN (TXT)
      </button>
      <button class="btn btn-secondary" onclick="closeModal()">
        Cerrar
      </button>
    </div>
  `;
}

document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
});

function setupEventListeners() {
  document.getElementById('fileConsolidado1').addEventListener('change', e => loadFile(e, 1));
  document.getElementById('fileConsolidado2').addEventListener('change', e => loadFile(e, 2));
  document.getElementById('fileNodos').addEventListener('change', e => loadFile(e, 3));
  document.getElementById('fileFMS').addEventListener('change', e => loadFile(e, 4));

  document.getElementById('filterCATEC').addEventListener('change', applyFilters);
  document.getElementById('showAllStates').addEventListener('change', applyFilters);
  document.getElementById('filterFTTH').addEventListener('change', e => {
    if (e.target.checked) document.getElementById('filterExcludeFTTH').checked = false;
    applyFilters();
  });
  document.getElementById('filterExcludeFTTH').addEventListener('change', e => {
    if (e.target.checked) document.getElementById('filterFTTH').checked = false;
    applyFilters();
  });
  document.getElementById('filterNodoEstado').addEventListener('change', applyFilters);
  document.getElementById('filterCMTS').addEventListener('change', applyFilters);
  document.getElementById('daysWindow').addEventListener('change', applyFilters);
  document.getElementById('filterTerritorio').addEventListener('change', applyFilters);
  document.getElementById('filterSistema').addEventListener('change', applyFilters);
  document.getElementById('filterAlarma').addEventListener('change', applyFilters);
  document.getElementById('quickSearch').addEventListener('input', debounce(applyFilters, 300));
  document.getElementById('ordenarPorIngreso').addEventListener('change', applyFilters);
}

async function loadFile(e, tipo) {
  const file = e.target.files[0];
  if (!file) return;

  const statusEl = document.getElementById(`status${tipo}`);
  statusEl.textContent = 'Cargando...';
  statusEl.classList.remove('loaded');

  let result;

  const fileExtension = file.name.split('.').pop()?.toLowerCase();

  if (tipo === 4 && fileExtension === 'csv') {
    result = await dataProcessor.loadCSV(file);
  } else {
    result = await dataProcessor.loadExcel(file, tipo);
  }

  if (result.success) {
    statusEl.textContent = `✓ ${result.rows} filas cargadas`;
    statusEl.classList.add('loaded');

    const nombres = ['Consolidado 1', 'Consolidado 2', 'Nodos UP/DOWN', 'Alarmas FMS'];
    toast(`${nombres[tipo-1]} cargado: ${result.rows} registros`);

    if ((dataProcessor.consolidado1 || dataProcessor.consolidado2)) {
      document.getElementById('mergeStatus').style.display = 'flex';
      processData();
    }
  } else {
    statusEl.textContent = `✗ Error`;
    toast(`Error al cargar archivo: ${result.error}`);
  }
}

function processData() {
  const merged = dataProcessor.merge();
  if (!merged.length) return;

  const daysWindow = parseInt(document.getElementById('daysWindow').value);

  const prepared = dataProcessor.prepareOrders(merged);
  baseZoneGroups = dataProcessor.processZones(prepared);
  const zonesForAnalysis = baseZoneGroups.map(z => ({...z, ordenes: z.ordenes.slice()}));

  allZones = dataProcessor.analyzeZones(zonesForAnalysis, daysWindow);
  allCMTS = dataProcessor.analyzeCMTS(allZones);

  currentData = {
    ordenes: prepared,
    zonas: baseZoneGroups
  };

  equipmentCache = buildEquipmentCache(prepared);
  window.equipmentCache = equipmentCache;

  populateFilters();

  const stats = {
    total: allZones.reduce((sum, z) => sum + z.totalOTs, 0),
    zonas: allZones.length,
    criticas: allZones.filter(z => z.criticidad === 'CRÍTICO').length,
    ftth: allZones.filter(z => z.tipo === 'FTTH').length,
    conAlarmas: allZones.filter(z => z.tieneAlarma).length,
    nodosCriticos: allZones.filter(z => z.nodoEstado === 'critical').length
  };

  document.getElementById('statsGrid').innerHTML = UIRenderer.renderStats(stats);
  document.getElementById('btnExportExcel').disabled = false;
  document.getElementById('btnExportExcelZonas').disabled = false;

  let statusText = `✓ ${merged.length} órdenes procesadas (deduplicadas)`;
  if (dataProcessor.nodosData) {
    statusText += ` • ${dataProcessor.nodosMap.size} nodos integrados`;
  }
  if (dataProcessor.fmsData) {
    statusText += ` • ${dataProcessor.fmsMap.size} zonas con alarmas`;
  }
  statusText += ` • Ventana: ${daysWindow} días`;
  document.getElementById('mergeStatusText').textContent = statusText;

  applyFilters();
}

function buildEquipmentCache(orders) {
  const cache = new Map();

  orders.forEach(order => {
    const meta = order.__meta || {};
    if (!meta.zonaPrincipal || !meta.dispositivos || !meta.dispositivos.length) return;

    if (!cache.has(meta.zonaPrincipal)) {
      cache.set(meta.zonaPrincipal, []);
    }

    const target = cache.get(meta.zonaPrincipal);
    meta.dispositivos.forEach(device => {
      target.push({
        zona: meta.zonaPrincipal,
        numCaso: meta.numeroCaso || order['Número de cita'] || '',
        sistema: meta.sistema || '',
        serialNumber: device.serialNumber || '',
        macAddress: device.macAddress || '',
        tipo: device.description || device.type || '',
        marca: device.category || '',
        modelo: device.model || ''
      });
    });
  });

  console.log(`✅ Cache de equipos construida para ${cache.size} zonas`);
  return cache;
}

function populateFilters() {
  const territorios = [...new Set(currentData.ordenes.map(o => o['Territorio de servicio: Nombre']).filter(Boolean))];
  const terrSelect = document.getElementById('filterTerritorio');
  terrSelect.innerHTML = '<option value="">Todos</option>' + territorios.sort().map(t => `<option value="${t}">${t}</option>`).join('');

  const cmtsList = [...new Set(allZones.map(z => z.cmts).filter(Boolean))];
  const cmtsSelect = document.getElementById('filterCMTS');
  cmtsSelect.innerHTML = '<option value="">Todos</option>' + cmtsList.sort().map(c => `<option value="${c}">${c}</option>`).join('');
}

function applyFilters() {
  if (!currentData) return;

  Filters.catec = document.getElementById('filterCATEC').checked;
  Filters.showAllStates = document.getElementById('showAllStates').checked;
  Filters.ftth = document.getElementById('filterFTTH').checked;
  Filters.excludeFTTH = document.getElementById('filterExcludeFTTH').checked;
  Filters.nodoEstado = document.getElementById('filterNodoEstado').value;
  Filters.cmts = document.getElementById('filterCMTS').value;
  Filters.days = parseInt(document.getElementById('daysWindow').value);
  Filters.territorio = document.getElementById('filterTerritorio').value;
  Filters.sistema = document.getElementById('filterSistema').value;
  Filters.alarma = document.getElementById('filterAlarma').value;
  Filters.quickSearch = document.getElementById('quickSearch').value;
  Filters.ordenarPorIngreso = document.getElementById('ordenarPorIngreso').value;

  console.log(`✅ Aplicando filtros - Ventana: ${Filters.days} días`);
  console.log(`📊 Estados activos: ${CONFIG.estadosPermitidos.join(', ')}`);
  console.log(`🔥 Ordenamiento: ${Filters.ordenarPorIngreso || 'Por score'}`);

  const filtered = Filters.apply(currentData.ordenes);
  console.log(`Órdenes después de filtros: ${filtered.length}`);

  const filteredSet = new Set(filtered);
  const zones = baseZoneGroups
    .map(zone => {
      const matched = zone.ordenes.filter(order => filteredSet.has(order));
      if (!matched.length) return null;
      return {
        ...zone,
        ordenes: matched
      };
    })
    .filter(Boolean);

  let analyzed = dataProcessor.analyzeZones(zones, Filters.days);

  analyzed = Filters.applyToZones(analyzed);

  const cmtsFiltered = dataProcessor.analyzeCMTS(analyzed);

  const stats = {
    total: analyzed.reduce((sum, z) => sum + z.totalOTs, 0),
    zonas: analyzed.length,
    criticas: analyzed.filter(z => z.criticidad === 'CRÍTICO').length,
    ftth: analyzed.filter(z => z.tipo === 'FTTH').length,
    conAlarmas: analyzed.filter(z => z.tieneAlarma).length,
    nodosCriticos: analyzed.filter(z => z.nodoEstado === 'critical').length
  };
  document.getElementById('statsGrid').innerHTML = UIRenderer.renderStats(stats);

  window.lastFilteredOrders = filtered;
  window.currentAnalyzedZones = analyzed;
  window.currentCMTSData = cmtsFiltered;
  allZones = analyzed;
  allCMTS = cmtsFiltered;

  document.getElementById('btnExportExcel').disabled = analyzed.length===0 && filtered.length===0;
  document.getElementById('btnExportExcelZonas').disabled = analyzed.length===0;

  document.getElementById('zonasPanel').innerHTML = UIRenderer.renderZonas(analyzed);
  document.getElementById('cmtsPanel').innerHTML = UIRenderer.renderCMTS(cmtsFiltered);
  document.getElementById('edificiosPanel').innerHTML = UIRenderer.renderEdificios(filtered);
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(filtered);
}

function clearFilters() {
  document.getElementById('filterCATEC').checked = false;
  document.getElementById('showAllStates').checked = false;
  document.getElementById('filterFTTH').checked = false;
  document.getElementById('filterExcludeFTTH').checked = false;
  document.getElementById('filterNodoEstado').value = '';
  document.getElementById('filterCMTS').value = '';
  document.getElementById('daysWindow').value = '3';
  document.getElementById('filterTerritorio').value = '';
  document.getElementById('filterSistema').value = '';
  document.getElementById('filterAlarma').value = '';
  document.getElementById('quickSearch').value = '';
  document.getElementById('ordenarPorIngreso').value = 'desc';
  applyFilters();
}

function switchTab(tabName) {
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(panel => panel.classList.remove('active'));

  document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
  document.getElementById(`${tabName}Panel`).classList.add('active');
}

function showAlarmaInfo(zoneIdx) {
  const zonaData = window.currentAnalyzedZones[zoneIdx];
  if (!zonaData || !zonaData.alarmas.length) return;

  let html = '<div class="alarma-info-box">';
  html += `<h4>🚨 Alarmas en Zona: ${zonaData.zona}</h4>`;

  zonaData.alarmas.forEach((a, idx) => {
    html += `<div style="margin-bottom: 20px; padding-bottom: 20px; border-bottom: 1px solid var(--border-subtle);">`;
    html += `<h5 style="color: var(--win-blue); margin-bottom: 10px;">Alarma ${idx + 1} ${a.isActive ? '<span class="badge badge-alarma-activa">ACTIVO</span>' : '<span class="badge">CERRADO</span>'}</h5>`;
    html += `<div class="alarma-info-grid">`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Tipo Alarma</span><span class="alarma-info-value">${a.type}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Elemento</span><span class="alarma-info-value">${FMS_TIPOS[a.elementType] || a.elementType}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Código Elemento</span><span class="alarma-info-value">${a.elementCode}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Fecha Creación</span><span class="alarma-info-value">${a.creationDate}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Fecha Cierre</span><span class="alarma-info-value">${a.recoveryDate || 'ABIERTO'}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Incidente</span><span class="alarma-info-value">${a.damage}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Clasificación</span><span class="alarma-info-value">${a.incidentClassification}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Reclamos</span><span class="alarma-info-value">${a.claims}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">CM Count</span><span class="alarma-info-value">${a.cmCount}</span></div>`;
    html += `</div></div>`;
  });

  html += '</div>';

  document.getElementById('alarmaModalBody').innerHTML = html;
  document.getElementById('alarmaBackdrop').classList.add('show');
  document.body.classList.add('modal-open');
}

function closeAlarmaModal() {
  document.getElementById('alarmaBackdrop').classList.remove('show');
  document.body.classList.remove('modal-open');
}

function showEdificioDetail(idx) {
  const edificio = window.edificiosData[idx];
  if (!edificio) return;

  document.getElementById('edificioModalTitle').textContent = `🏢 Edificio: ${edificio.direccion}`;

  let html = `<div style="margin-bottom: 16px; padding: 12px; background: var(--bg-tertiary); border-radius: 8px;">
    <strong>Zona:</strong> ${edificio.zona}<br>
    <strong>Territorio:</strong> ${edificio.territorio}<br>
    <strong>Total OTs:</strong> ${edificio.casos.length}
  </div>`;

  html += '<div class="table-container"><div class="table-wrapper"><table class="detail-table"><thead><tr>';
  html += '<th>Número de Caso</th><th>Estado</th><th>Diagnóstico</th><th>Fecha Creación</th>';
  html += '</tr></thead><tbody>';

  edificio.casos.forEach(o => {
    const numCaso = o['Número del caso'] || o['Caso Externo'] || '';
    const sistema = TextUtils.detectarSistema(numCaso);
    let badgeSistema = '';
    if (sistema === 'OPEN') {
      badgeSistema = '<span class="badge badge-open">OPEN</span>';
    } else if (sistema === 'FAN') {
      badgeSistema = '<span class="badge badge-fan">FAN</span>';
    }

    const estado = o['Estado'] || o['Estado.1'] || o['Estado.2'] || '';
    const diag = o['Diagnostico Tecnico'] || o['Diagnóstico Técnico'] || '-';
    const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || '';

    html += `<tr>
      <td>${numCaso} ${badgeSistema}</td>
      <td>${estado}</td>
      <td>${diag.substring(0, 100)}${diag.length > 100 ? '...' : ''}</td>
      <td>${fecha}</td>
    </tr>`;
  });

  html += '</tbody></table></div></div>';

  document.getElementById('edificioModalBody').innerHTML = html;
  document.getElementById('edificioBackdrop').classList.add('show');
  document.body.classList.add('modal-open');
}

function closeEdificioModal() {
  document.getElementById('edificioBackdrop').classList.remove('show');
  document.body.classList.remove('modal-open');
}

function openModal(zoneIdx) {
  const zonaData = window.currentAnalyzedZones[zoneIdx];
  if (!zonaData) return;

  currentZone = zonaData;
  selectedOrders.clear();

  setZoneModalFooter();

  document.getElementById('modalTitle').textContent = `Detalle: ${currentZone.zona}`;

  const ordenesParaModal = currentZone.ordenesOriginales || currentZone.ordenes;
  currentZone.ordenes = ordenesParaModal;
  currentZone.ordenes.forEach((order, idx) => ensureOrderRowKey(order, idx, currentZone.zona));

  const modalFilters = document.getElementById('modalFilters');
  if (modalFilters) {
    modalFilters.style.display = 'flex';
  }

  const dias = [...new Set(ordenesParaModal.map(o => {
    const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || o['Fecha de inicio'];
    const dt = DateUtils.parse(fecha);
    return dt ? DateUtils.format(dt) : null;
  }).filter(Boolean))].sort();

  const diaSelect = document.getElementById('filterDia');
  diaSelect.innerHTML = '<option value="">Todos los días</option>' + dias.map(d => `<option value="${d}">${d}</option>`).join('');

  renderModalContent();

  document.getElementById('modalBackdrop').classList.add('show');
  document.body.classList.add('modal-open');
}

function openCMTSDetail(cmts) {
  console.log('=== openCMTSDetail llamado ===');
  console.log('CMTS solicitado:', cmts);

  const cmtsDataSource = window.currentCMTSData || allCMTS;

  if (!cmtsDataSource || !cmtsDataSource.length) {
    console.error('No hay datos de CMTS disponibles');
    toast('❌ No hay datos de CMTS disponibles');
    return;
  }

  const cmtsData = cmtsDataSource.find(c => c.cmts === cmts);

  if (!cmtsData) {
    console.error('CMTS no encontrado:', cmts);
    toast('❌ No se encontró el CMTS: ' + cmts);
    return;
  }

  let html = '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
  html += '<th>Zona</th><th>Tipo</th><th>Total OTs</th><th>N</th><th>N-1</th><th>Estado Nodo</th>';
  html += '</tr></thead><tbody>';

  cmtsData.zonas.forEach(z => {
    const badgeTipo = z.tipo === 'FTTH' ? '<span class="badge badge-ftth">FTTH</span>' : '<span class="badge badge-hfc">HFC</span>';

    let badgeNodo = '<span class="badge">Sin datos</span>';
    if (z.nodoEstado === 'up') {
      badgeNodo = `<span class="badge badge-up">✓ UP</span>`;
    } else if (z.nodoEstado === 'critical') {
      badgeNodo = `<span class="badge badge-critical">⚠ CRÍTICO</span>`;
    } else if (z.nodoEstado === 'down') {
      badgeNodo = `<span class="badge badge-down">↓ DOWN</span>`;
    }

    html += `<tr>
      <td><strong>${z.zona}</strong></td>
      <td>${badgeTipo}</td>
      <td class="number">${z.totalOTs}</td>
      <td class="number">${z.ingresoN}</td>
      <td class="number">${z.ingresoN1}</td>
      <td>${badgeNodo}</td>
    </tr>`;
  });

  html += '</tbody></table></div></div>';

  document.getElementById('modalTitle').textContent = `Zonas del CMTS: ${cmts}`;
  document.getElementById('modalBody').innerHTML = html;
  document.getElementById('modalFilters').style.display = 'none';

  document.getElementById('modalFooter').innerHTML = `
    <div></div>
    <div style="display:flex; gap:8px;">
      <button class="btn btn-primary" onclick="exportModalDetalleExcel()">📥 Exportar detalle</button>
      <button class="btn btn-secondary" onclick="closeModal()">Cerrar</button>
    </div>`;

  document.getElementById('modalBackdrop').classList.add('show');
  document.body.classList.add('modal-open');
}

function closeModal() {
  document.getElementById('modalBackdrop').classList.remove('show');
  document.body.classList.remove('modal-open');
  selectedOrders.clear();
  document.getElementById('filterHorario').value = '';
  document.getElementById('filterDia').value = '';

  const modalFilters = document.getElementById('modalFilters');
  if (modalFilters) {
    modalFilters.style.display = 'flex';
  }
  setZoneModalFooter();
  updateSelectionInfo();
  document.getElementById('modalFilters').style.display = 'flex';
  document.getElementById('modalFooter').innerHTML = `
    <div class="selection-info" id="selectionInfo">0 órdenes seleccionadas</div>
    <div style="display: flex; gap: 8px;">
      <button class="btn btn-primary" onclick="exportModalDetalleExcel()">
        📥 Exportar detalle
      </button>
      <button class="btn btn-warning" id="btnExportBEFAN" disabled onclick="exportBEFAN()">
        📤 Exportar BEFAN (TXT)
      </button>
      <button class="btn btn-secondary" onclick="closeModal()">Cerrar</button>
    </div>
  `;
}

function applyModalFilters() {
  renderModalContent();
}

function getFilteredModalOrders(horario, dia) {
  if (!currentZone || !Array.isArray(currentZone.ordenes)) {
    return [];
  }

  let ordenes = currentZone.ordenes.slice();

  if (horario) {
    const [start, end] = horario.split('-').map(Number);
    ordenes = ordenes.filter(o => {
      const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || o['Fecha de inicio'];
      const dt = DateUtils.parse(fecha);
      if (!dt) return false;
      const hour = DateUtils.getHour(dt);
      return hour >= start && hour < end;
    });
  }

  if (dia) {
    ordenes = ordenes.filter(o => {
      const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || o['Fecha de inicio'];
      const dt = DateUtils.parse(fecha);
      return dt && DateUtils.format(dt) === dia;
    });
  }

  return ordenes;
}

function renderModalContent() {
  const horario = document.getElementById('filterHorario').value;
  const dia = document.getElementById('filterDia').value;

  const ordenes = getFilteredModalOrders(horario, dia);

  let html = '<div class="chart-container">';
  html += '<div class="chart-title">📊 Distribución de Ingresos (7 días)</div>';
  html += '<div style="display: flex; justify-content: space-around; align-items: flex-end; height: 150px; padding: 10px;">';

  currentZone.last7DaysCounts.forEach((count, i) => {
    const height = currentZone.last7DaysCounts.length > 0 ? (count / Math.max(...currentZone.last7DaysCounts, 1)) * 120 : 0;
    html += `<div style="text-align: center;">
      <div style="width: 60px; height: ${height}px; background: linear-gradient(180deg, #0078D4 0%, #005A9E 100%); border-radius: 4px 4px 0 0; margin: 0 auto;"></div>
      <div style="font-size: 0.75rem; font-weight: 700; margin-top: 4px;">${count}</div>
      <div style="font-size: 0.6875rem; color: var(--text-secondary);">${currentZone.last7Days[i]}</div>
    </div>`;
  });

  html += '</div></div>';

  html += `<div style="margin-bottom: 16px; padding: 12px; background: var(--bg-tertiary); border-radius: 8px;">
    <strong>Total órdenes mostradas:</strong> ${ordenes.length} de ${currentZone.ordenes.length}
  </div>`;

  html += '<div class="table-container"><div style="max-height: 400px; overflow-y: auto;"><table class="detail-table"><thead><tr>';
  html += '<th style="position: sticky; top: 0; z-index: 10;"><input type="checkbox" id="selectAllCheckbox" onchange="toggleSelectAll()"></th>';

  const allCols = new Set();
  ordenes.forEach(o => {
    Object.keys(o).forEach(k => {
      if (!k.startsWith('_')) allCols.add(k);
    });
  });

  const columns = Array.from(allCols);
  columns.forEach(col => {
    html += `<th style="position: sticky; top: 0; z-index: 10; white-space: nowrap;">${col}</th>`;
  });

  html += '</tr></thead><tbody>';

  ordenes.forEach((o, idx) => {
    const rowKey = ensureOrderRowKey(o, idx, currentZone?.zona);
    const checked = selectedOrders.has(rowKey) ? 'checked' : '';
    const safeRowKey = String(rowKey).replace(/"/g, '&quot;');
    html += `<tr>`;
    html += `<td><input type="checkbox" class="order-checkbox" data-cita="${safeRowKey}" ${checked} onchange="updateSelection()"></td>`;

    columns.forEach(col => {
      let value = o[col] || '';

      if (col.includes('Número del caso') || col.includes('Caso')) {
        const sistema = TextUtils.detectarSistema(value);
        if (sistema) {
          const badgeClass = sistema === 'OPEN' ? 'badge-open' : 'badge-fan';
          value = `${value} <span class="badge ${badgeClass}">${sistema}</span>`;
        }
      }

      if (typeof value === 'string' && value.length > 100) {
        value = value.substring(0, 100) + '...';
      }

      html += `<td title="${value}">${value}</td>`;
    });

    html += '</tr>';
  });

  html += '</tbody></table></div></div>';

  document.getElementById('modalBody').innerHTML = html;
  updateSelectionInfo();
}

function toggleSelectAll() {
  const checked = document.getElementById('selectAllOrders')?.checked || 
                 document.getElementById('selectAllCheckbox')?.checked || false;

  document.querySelectorAll('.order-checkbox').forEach(cb => {
    cb.checked = checked;
    const cita = cb.getAttribute('data-cita');
    if (checked) {
      selectedOrders.add(cita);
    } else {
      selectedOrders.delete(cita);
    }
  });

  updateSelectionInfo();
}

function updateSelection() {
  selectedOrders.clear();
  document.querySelectorAll('.order-checkbox:checked').forEach(cb => {
    selectedOrders.add(cb.getAttribute('data-cita'));
  });
  updateSelectionInfo();
}

function updateSelectionInfo() {
  const count = selectedOrders.size;
  const infoEl = document.getElementById('selectionInfo');
  if (infoEl) {
    infoEl.textContent = `${count} órdenes seleccionadas`;
  }
  const btnEl = document.getElementById('btnExportBEFAN');
  if (btnEl) {
    btnEl.disabled = count === 0;
  }
}

function exportBEFAN() {
  if (selectedOrders.size === 0) {
    toast('No hay órdenes seleccionadas');
    return;
  }

  const ordenes = currentZone.ordenes.filter((o, idx) => selectedOrders.has(ensureOrderRowKey(o, idx, currentZone?.zona)));

  let txt = '';
  let contadorExitoso = 0;

  ordenes.forEach(o => {
    const casoExterno = o['Caso Externo'] || 
                       o['External Case Id'] || 
                       o['Número del caso'] ||
                       o['Caso externo'] ||
                       o['CASO EXTERNO'] ||
                       o['external_case_id'] || '';

    if (casoExterno && String(casoExterno).trim() !== '') {
      txt += `${String(casoExterno).trim()}|${CONFIG.codigo_befan}\n`;
      contadorExitoso++;
    }
  });

  if (!txt || contadorExitoso === 0) {
    toast('⚠️ Las órdenes seleccionadas no tienen Caso Externo válido');
    return;
  }

  try {
    const blob = new Blob([txt], {type: 'text/plain;charset=utf-8'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `BEFAN_${currentZone.zona}_${new Date().toISOString().slice(0,10)}.txt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    toast(`✅ Exportadas ${contadorExitoso} órdenes a BEFAN`);
  } catch (error) {
    console.error('Error al descargar BEFAN:', error);
    toast('❌ Error al descargar archivo BEFAN');
  }
}

function exportZoneOrdersExcel() {
  if (!currentZone || !Array.isArray(currentZone.ordenes) || !currentZone.ordenes.length) {
    toast('❌ No hay una zona activa para exportar');
    return;
  }

  const horario = document.getElementById('filterHorario')?.value || '';
  const dia = document.getElementById('filterDia')?.value || '';

  let ordenes = getFilteredModalOrders(horario, dia);

  if (!ordenes.length) {
    toast('⚠️ No hay órdenes para exportar con los filtros aplicados');
    return;
  }

  if (selectedOrders.size > 0) {
    ordenes = ordenes.filter((o, idx) => selectedOrders.has(ensureOrderRowKey(o, idx, currentZone?.zona)));
    if (!ordenes.length) {
      toast('⚠️ La selección actual no tiene órdenes disponibles para exportar');
      return;
    }
  }

  const sanitized = ordenes.map(order => {
    const clean = {};
    Object.keys(order).forEach(key => {
      if (!key.startsWith('_')) {
        clean[key] = order[key];
      }
    });
    return clean;
  });

  if (!sanitized.length) {
    toast('⚠️ No hay datos exportables para la zona seleccionada');
    return;
  }

  const wb = XLSX.utils.book_new();
  const rawSheetName = currentZone.zona || 'Zona';
  const sheetName = rawSheetName.toString().slice(0, 28) || 'Zona';
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sanitized), sheetName);

  const fecha = new Date().toISOString().slice(0, 10);
  const fileSafeZone = sheetName.replace(/[^A-Za-z0-9_-]/g, '_') || 'Zona';
  XLSX.writeFile(wb, `Zona_${fileSafeZone}_${fecha}.xlsx`);

  toast(`✅ Exportadas ${sanitized.length} órdenes de ${rawSheetName}`);
}

function exportExcelVista(){
  const wb = XLSX.utils.book_new();

  if (Array.isArray(window.currentAnalyzedZones) && window.currentAnalyzedZones.length){
    const zonasData = window.currentAnalyzedZones.map(z => ({
      Zona: z.zona,
      Tipo: z.tipo,
      Red: z.tipo === 'FTTH' ? z.zonaHFC : '',
      CMTS: z.cmts,
      Tiene_Alarma: z.tieneAlarma ? 'SÍ' : 'NO',
      Alarmas_Activas: z.alarmasActivas,
      Total_OTs: z.totalOTs,
      Ingreso_N: z.ingresoN,
      Ingreso_N1: z.ingresoN1,
      Max_Dia: z.maxDia
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(zonasData), 'Zonas');
  }

  if (Array.isArray(window.currentCMTSData) && window.currentCMTSData.length){
    const cmtsData = window.currentCMTSData.map(c => ({
      CMTS: c.cmts,
      Zonas: c.zonas.length,
      Total_OTs: c.totalOTs,
      Zonas_UP: c.zonasUp,
      Zonas_DOWN: c.zonasDown,
      Zonas_Criticas: c.zonasCriticas
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cmtsData), 'CMTS');
  }

  if (window.edificiosData && window.edificiosData.length){
    const edi = window.edificiosData.map(e => ({
      Direccion: e.direccion,
      Zona: e.zona,
      Territorio: e.territorio,
      Total_OTs: e.casos.length
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(edi), 'Edificios');
  }

  if (window.equiposPorZona && window.equiposPorZona.size){
    const todos = [];
    window.equiposPorZona.forEach((arr, zona)=>{
      arr.forEach(it=>todos.push({...it, zona}));
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(todos), 'Equipos');
  }

  const fecha = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `Vista_Filtrada_${fecha}.xlsx`);
  toast('✓ Vista filtrada exportada');
}

function exportExcelZonasCrudo(){
  if (!Array.isArray(window.currentAnalyzedZones) || !window.currentAnalyzedZones.length) {
    return toast('No hay zonas para exportar');
  }
  const wb = XLSX.utils.book_new();
  const zonasData = window.currentAnalyzedZones.map(z => ({
    Zona: z.zona,
    Tipo: z.tipo,
    Red: z.tipo === 'FTTH' ? z.zonaHFC : '',
    CMTS: z.cmts,
    Nodo_Estado: z.nodoEstado,
    Nodo_UP: z.nodoUp,
    Nodo_DOWN: z.nodoDown,
    Tiene_Alarma: z.tieneAlarma ? 'SÍ' : 'NO',
    Alarmas_Activas: z.alarmasActivas,
    Territorio: z.territorio,
    Total_OTs: z.totalOTs,
    Ingreso_N: z.ingresoN,
    Ingreso_N1: z.ingresoN1,
    Max_Dia: z.maxDia
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(zonasData), 'Zonas');
  const fecha = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `Zonas_Crudo_${fecha}.xlsx`);
}

function exportModalDetalleExcel(){
  const title = document.getElementById('modalTitle').textContent || '';

  if (title.startsWith('Detalle: ') && window.currentZone){
    exportZoneOrdersExcel();
    return;
  }

  const wb = XLSX.utils.book_new();

  if (title.startsWith('Zonas del CMTS: ')){
    const cmts = title.replace('Zonas del CMTS: ','').trim();
    const data = (window.currentCMTSData||[]).find(c=>c.cmts===cmts);
    if (!data) return toast('Sin datos de CMTS');
    const zonasFlat = data.zonas.map(z=>({
      Zona: z.zona, Tipo: z.tipo, Total_OTs: z.totalOTs, Ingreso_N: z.ingresoN, Ingreso_N1: z.ingresoN1, Estado_Nodo: z.nodoEstado
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(zonasFlat), `CMTS_${cmts.slice(0,20)}`);
  } else {
    const edTitle = document.getElementById('edificioModalTitle')?.textContent||'';
    if (edTitle.startsWith('🏢 Edificio: ') && window.edificiosData){
      const dir = edTitle.replace('🏢 Edificio: ','').trim();
      const e = window.edificiosData.find(x=>x.direccion===dir);
      if (!e) return toast('Sin datos del edificio');
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(e.casos), 'Edificio_OTs');
    } else {
      return toast('No hay detalle para exportar');
    }
  }

  const fecha = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `Detalle_${fecha}.xlsx`);
}

console.log('✅ Panel v4.9 MEJORADO inicializado');
console.log('🔥 NUEVAS FUNCIONALIDADES:');
console.log('   • Ordenamiento de zonas por ingreso (mayor/menor)');
console.log('   • Filtros de equipos por modelo y marca');
console.log('   • Exportación filtrada de equipos');
console.log('📊 Estados activos:', CONFIG.estadosPermitidos);

window.toggleUtilities = toggleUtilities;
window.openPlanillasNewTab = openPlanillasNewTab;
window.applyEquiposFilters = applyEquiposFilters;
window.clearEquiposFilters = clearEquiposFilters;
window.toggleEquiposGrupo = toggleEquiposGrupo;
window.exportEquiposGrupoExcel = exportEquiposGrupoExcel;
window.clearFilters = clearFilters;
window.switchTab = switchTab;
window.showAlarmaInfo = showAlarmaInfo;
window.closeAlarmaModal = closeAlarmaModal;
window.showEdificioDetail = showEdificioDetail;
window.closeEdificioModal = closeEdificioModal;
window.openModal = openModal;
window.openCMTSDetail = openCMTSDetail;
window.closeModal = closeModal;
window.applyModalFilters = applyModalFilters;
window.toggleSelectAll = toggleSelectAll;
window.exportBEFAN = exportBEFAN;
window.exportZoneOrdersExcel = exportZoneOrdersExcel;
window.exportExcelVista = exportExcelVista;
window.exportExcelZonasCrudo = exportExcelZonasCrudo;
window.exportModalDetalleExcel = exportModalDetalleExcel;
