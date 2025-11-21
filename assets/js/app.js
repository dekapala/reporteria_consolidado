console.log('🚀 Panel v5.0 CONTROLADOR - Lógica y Eventos');

// === FUNCIONES GLOBALES DE UTILIDAD ===

async function readFileAsUint8Array(file) {
  console.log(`📂 Leyendo archivo: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`);
  try {
    if (file.arrayBuffer) {
      const buf = await file.arrayBuffer();
      return new Uint8Array(buf);
    }
    return await new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => resolve(new Uint8Array(fr.result));
      fr.onerror = () => reject(new Error('Error al leer archivo con FileReader'));
      fr.readAsArrayBuffer(file);
    });
  } catch (error) {
    console.error(`❌ Error leyendo ${file.name}:`, error);
    throw error;
  }
}

function on(id, evt, fn) {
  const el = document.getElementById(id);
  if (el) {
    el.addEventListener(evt, fn);
    return true;
  }
  console.warn(`⚠️ Elemento #${id} no encontrado, evento '${evt}' no registrado`);
  return false;
}

// Usamos configuración global si existe (desde utils.js)
const CONFIG = window.CONFIG || {
  estadosPermitidos: ['NUEVA', 'EN PROGRESO'],
  estadosOcultosPorDefecto: ['CANCELADA']
};

// Usamos FMS_TIPOS globales si existen
const FMS_TIPOS = window.FMS_TIPOS || {
  'ED': 'Edificio', 'CDO': 'CDO', 'CMTS': 'CMTS' // Fallback simple
};

const TerritorioUtils = {
  normalizar(territorio) {
    if (!territorio) return '';
    return territorio.replace(/\s*flex\s*\d+/gi, ' Flex').replace(/\s{2,}/g, ' ').trim();
  },
  extraerFlex(territorio) {
    if (!territorio) return null;
    const match = territorio.match(/flex\s*(\d+)/i);
    return match ? parseInt(match[1]) : null;
  },
  tieneFlex(territorio) {
    if (!territorio) return false;
    return /flex\s*\d+/i.test(territorio);
  }
};

function normalizeEstado(estado) {
  if (!estado) return '';
  return String(estado).toUpperCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

// === MANEJO DE ESTADO GLOBAL ===
const zoneFilterState = {
  allOptions: [],
  filteredOptions: [],
  selected: new Set(),
  search: '',
  open: false
};

const Filters = {
  catec: false,
  excludeCatec: false,
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
  ordenarPorIngreso: 'desc',
  equipoModelo: [],
  equipoMarca: '',
  equipoTerritorio: '',
  criticidad: '',
  selectedZonas: [],
  
  apply(rows) {
    let filtered = rows.slice();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    filtered = filtered.filter(r => {
      const fecha = r['Fecha de creación'] || r['Fecha/Hora de apertura'] || r['Fecha de inicio'];
      const dt = DateUtils.parse(fecha);
      if (!dt) return false;
      const daysAgo = DateUtils.daysBetween(today, dt);
      return daysAgo <= this.days;
    });
    
    if (!this.showAllStates) {
      const estadosOcultosNorm = (CONFIG.estadosOcultosPorDefecto || []).map(normalizeEstado);
      const estadosPermitidosNorm = (CONFIG.estadosPermitidos || []).map(normalizeEstado);

      filtered = filtered.filter(r => {
        const estRaw = r['Estado.1'] || r['Estado'] || r['Estado.2'] || '';
        const est = normalizeEstado(estRaw);
        if (!est) return false;
        if (estadosOcultosNorm.includes(est)) return false;
        return estadosPermitidosNorm.includes(est);
      });
    }
    
    // Lógica CATEC
    const catecActive = this.catec;
    const excludeCatecActive = !catecActive && this.excludeCatec;
    if (catecActive) {
      filtered = filtered.filter(r => String(r['Tipo de trabajo: Nombre de tipo de trabajo'] || '').toUpperCase().includes('CATEC'));
    }
    if (excludeCatecActive) {
      filtered = filtered.filter(r => !String(r['Tipo de trabajo: Nombre de tipo de trabajo'] || '').toUpperCase().includes('CATEC'));
    }
    
    // Lógica FTTH
    if (this.ftth) filtered = filtered.filter(r => /9\d{2}/.test(r['Zona Tecnica HFC'] || ''));
    if (this.excludeFTTH) filtered = filtered.filter(r => !/9\d{2}/.test(r['Zona Tecnica HFC'] || ''));
    
    // Otros filtros
    if (this.territorio) filtered = filtered.filter(r => (r['Territorio de servicio: Nombre'] || '') === this.territorio);
    if (this.sistema) filtered = filtered.filter(r => TextUtils.detectarSistema(r['Número del caso'] || r['Caso Externo'] || '') === this.sistema);
    
    if (this.quickSearch) {
      const queries = this.quickSearch.split(';').map(q => q.trim()).filter(Boolean);
      filtered = filtered.filter(r => {
        const searchable = [r['Zona Tecnica HFC'], r['Zona Tecnica FTTH'], r['Calle'], r['Caso Externo'], r['External Case Id'], r['Número del caso']].join(' ');
        return queries.length === 1 ? TextUtils.matches(searchable, queries[0]) : TextUtils.matchesMultiple(searchable, queries);
      });
    }

    if (this.selectedZonas && this.selectedZonas.length) {
      const zonesSet = new Set(this.selectedZonas);
      filtered = filtered.filter(r => {
        const { zonaPrincipal } = dataProcessor.getZonaPrincipal(r);
        return zonaPrincipal && zonesSet.has(zonaPrincipal);
      });
    }
    return filtered;
  },

  applyToZones(zones) {
    let filtered = zones.slice();

    if (this.criticidad === 'critico') filtered = filtered.filter(z => z.criticidad === 'CRÍTICO');
    
    if (this.nodoEstado) {
      filtered = filtered.filter(z => {
        if (this.nodoEstado === 'up') return z.nodoEstado === 'up';
        if (this.nodoEstado === 'down') return z.nodoEstado === 'down' || z.nodoEstado === 'critical';
        if (this.nodoEstado === 'critical') return z.nodoEstado === 'critical';
        return true;
      });
    }
    
    if (this.cmts) filtered = filtered.filter(z => z.cmts === this.cmts);
    
    if (this.alarma) {
      if (this.alarma === 'con-alarma') filtered = filtered.filter(z => z.tieneAlarma);
      else if (this.alarma === 'sin-alarma') filtered = filtered.filter(z => !z.tieneAlarma);
    }

    if (this.selectedZonas && this.selectedZonas.length) {
      const zonesSet = new Set(this.selectedZonas);
      filtered = filtered.filter(z => zonesSet.has(z.zona));
    }

    if (this.ordenarPorIngreso === 'desc') filtered.sort((a, b) => b.totalOTs - a.totalOTs);
    else if (this.ordenarPorIngreso === 'asc') filtered.sort((a, b) => a.totalOTs - b.totalOTs);
    
    return filtered;
  }
};

// Exponer Filtros globalmente para que UI Renderer pueda leerlos
window.Filters = Filters;

// === LÓGICA DE FILTROS DE ZONAS ===
// ... (Funciones auxiliares de filtros de zona: initializeZoneFilterOptions, renderZoneFilterOptions, etc. se mantienen igual)
function initializeZoneFilterOptions(zones) {
  const unique = Array.from(new Set((zones || []).filter(Boolean)));
  unique.sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));

  zoneFilterState.allOptions = unique;
  zoneFilterState.filteredOptions = unique.slice();
  zoneFilterState.selected = new Set();
  zoneFilterState.search = '';

  const searchInput = document.getElementById('zoneFilterSearch');
  if (searchInput) searchInput.value = '';

  renderZoneFilterOptions();
  updateZoneFilterSummary();
  Filters.selectedZonas = [];
}

function renderZoneFilterOptions() {
  const container = document.getElementById('zoneFilterOptions');
  if (!container) return;
  if (!zoneFilterState.allOptions.length) { container.innerHTML = '<div class="zone-multiselect__empty">Cargá los reportes para habilitar el filtro por zonas</div>'; return; }
  if (!zoneFilterState.filteredOptions.length) { container.innerHTML = '<div class="zone-multiselect__empty">No hay coincidencias para la búsqueda actual</div>'; return; }
  const optionsHtml = zoneFilterState.filteredOptions.map(z => {
    const checked = zoneFilterState.selected.has(z) ? 'checked' : '';
    return `<label class="zone-option"><input type="checkbox" value="${z}" ${checked}> <span>${z}</span></label>`;
  }).join('');
  container.innerHTML = optionsHtml;
}

function updateZoneFilterSummary() {
  const summaryEl = document.getElementById('zoneFilterSummary');
  const trigger = document.getElementById('zoneFilterTrigger');
  if (!summaryEl) return;
  const hasOptions = zoneFilterState.allOptions.length > 0;
  if (trigger) trigger.disabled = !hasOptions;
  if (!hasOptions) { summaryEl.textContent = 'Sin datos (carga reportes)'; summaryEl.title = 'Carga los reportes para habilitar el filtro por zonas'; return; }
  const count = zoneFilterState.selected.size;
  if (count === 0) { summaryEl.textContent = 'Todas las zonas'; summaryEl.title = 'Mostrar todas las zonas disponibles'; }
  else if (count === 1) { const [only] = Array.from(zoneFilterState.selected); summaryEl.textContent = only; summaryEl.title = only; }
  else { summaryEl.textContent = `${count} zonas seleccionadas`; const listPreview = Array.from(zoneFilterState.selected).slice(0, 8).join(', '); summaryEl.title = listPreview; }
}

function toggleZoneFilter(event) {
  if (event) { event.preventDefault(); event.stopPropagation(); }
  const container = document.getElementById('zoneFilter');
  if (!container) return;
  const willOpen = !container.classList.contains('open');
  if (willOpen) {
    container.classList.add('open'); zoneFilterState.open = true; renderZoneFilterOptions();
    setTimeout(() => { const searchInput = document.getElementById('zoneFilterSearch'); if (searchInput) { searchInput.focus(); if (zoneFilterState.search) { const pos = zoneFilterState.search.length; searchInput.setSelectionRange(pos, pos); } } }, 0);
  } else { closeZoneFilter(); }
}
function closeZoneFilter() { const container = document.getElementById('zoneFilter'); if (!container) return; container.classList.remove('open'); zoneFilterState.open = false; }
function handleZoneFilterOutsideClick(event) { const container = document.getElementById('zoneFilter'); if (!container) return; if (!container.contains(event.target)) { closeZoneFilter(); } }
function onZoneFilterSearch(event) {
  const value = event.target.value || ''; zoneFilterState.search = value;
  if (!zoneFilterState.allOptions.length) { renderZoneFilterOptions(); return; }
  const normalizedTerm = TextUtils.normalize(value);
  if (!normalizedTerm) { zoneFilterState.filteredOptions = zoneFilterState.allOptions.slice(); }
  else { zoneFilterState.filteredOptions = zoneFilterState.allOptions.filter(z => TextUtils.normalize(z).includes(normalizedTerm)); }
  renderZoneFilterOptions();
}
function onZoneOptionChange(event) {
  const target = event.target; if (!target || target.type !== 'checkbox') return; const zone = target.value; if (!zone) return;
  if (target.checked) { zoneFilterState.selected.add(zone); } else { zoneFilterState.selected.delete(zone); }
  Filters.selectedZonas = Array.from(zoneFilterState.selected); updateZoneFilterSummary(); applyFilters();
}
function selectAllZones(event) {
  if (event) { event.preventDefault(); event.stopPropagation(); } if (!zoneFilterState.allOptions.length) return;
  const hasSearch = Boolean(zoneFilterState.search); const source = hasSearch ? zoneFilterState.filteredOptions : zoneFilterState.allOptions; if (!source.length) return;
  source.forEach(z => zoneFilterState.selected.add(z)); Filters.selectedZonas = Array.from(zoneFilterState.selected); renderZoneFilterOptions(); updateZoneFilterSummary(); applyFilters();
}
function clearZoneSelection(event) { if (event) { event.preventDefault(); event.stopPropagation(); } resetZoneFilterState({ apply: true, keepSearch: true }); }
function resetZoneFilterState(options = {}) {
  const { apply = false, keepSearch = false } = options; zoneFilterState.selected = new Set();
  if (!keepSearch) { zoneFilterState.search = ''; zoneFilterState.filteredOptions = zoneFilterState.allOptions.slice(); const searchInput = document.getElementById('zoneFilterSearch'); if (searchInput) searchInput.value = ''; }
  else { const normalizedTerm = TextUtils.normalize(zoneFilterState.search); if (!normalizedTerm) { zoneFilterState.filteredOptions = zoneFilterState.allOptions.slice(); } else { zoneFilterState.filteredOptions = zoneFilterState.allOptions.filter(z => TextUtils.normalize(z).includes(normalizedTerm)); } }
  renderZoneFilterOptions(); updateZoneFilterSummary(); Filters.selectedZonas = []; if (apply) { applyFilters(); }
}

// === FILTROS EQUIPOS ===
function applyEquiposFilters() {
  const select = document.getElementById('filterEquipoModelo');
  Filters.equipoModelo = Array.from(select.selectedOptions).map(opt => opt.value);
  Filters.equipoMarca = document.getElementById('filterEquipoMarca').value;
  Filters.equipoTerritorio = document.getElementById('filterEquipoTerritorio').value;
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
}
function clearEquiposFilters() {
  Filters.equipoModelo = []; Filters.equipoMarca = ''; Filters.equipoTerritorio = '';
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
}
function toggleEquiposGrupo(zona){
  if (!window.equiposOpen) window.equiposOpen = new Set();
  if (window.equiposOpen.has(zona)) window.equiposOpen.delete(zona); else window.equiposOpen.add(zona);
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
}

// === FUNCIONES DE EXPORTACIÓN (Mantenidas igual, son lógica de negocio/controlador) ===
const ZONE_EXPORT_HEADERS = ['Fecha','ZonaHFC','ZonaFTTH','Territorio','Ubicacion','Caso','NumeroOrden','NumeroOTuca','Diagnostico','Tipo','TipoTrabajo','Estado1','Estado2','Estado3','MAC'];
const ORDER_FIELD_KEYS = {
  fecha: ['Fecha de creación','Fecha/Hora de apertura','Fecha de inicio','Fecha'],
  zonaHFC: ['Zona Tecnica HFC','Zona Tecnica','Zona HFC','Zona'],
  zonaFTTH: ['Zona Tecnica FTTH','Zona FTTH'],
  territorio: ['Territorio de servicio: Nombre','Territorio','Territorio servicio'],
  ubicacionCalle: ['Calle','Dirección','Direccion','Domicilio','Dirección del servicio','Direccion del servicio','Dirección Servicio'],
  ubicacionAltura: ['Altura','Altura (N°)','Altura domicilio','Número','Numero','Nro','Altura Domicilio'],
  ubicacionLocalidad: ['Localidad','Localidad Instalación','Localidad Instalacion','Ciudad','Partido','Distrito','Provincia'],
  caso: ['Número del caso','Numero del caso','Caso Externo','Caso externo','External Case Id','Nro Caso','Número de Caso','Numero Caso'],
  numeroOrden: ['Número de cita','Numero de cita','Número de Orden','Numero de Orden','Número Orden','Numero Orden','Orden','Orden de trabajo','N° Orden','Nro Orden','Número de orden de trabajo','Numero de orden de trabajo','N° OT'],
  numeroOTuca: ['Número OT UCA','Numero OT UCA','Número de OT UCA','Numero de OT UCA','NumeroOTUca','OT UCA','OT_UCA','Nro OT UCA'],
  diagnostico: ['Diagnostico Tecnico','Diagnóstico Técnico','Diagnostico','Diagnóstico','Diagnostico Cliente','Diagnostico tecnico'],
  tipo: ['Tipo','Tipo de caso','Tipo caso','Tipo OPEN','Tipo FAN','Tipo de OT','Tipo Servicio'],
  tipoTrabajo: ['Tipo de trabajo: Nombre de tipo de trabajo','Tipo de trabajo','Tipo Trabajo','TipoTrabajo'],
  estado1: ['Estado.1','Estado','Estado inicial'],
  estado2: ['Estado.2','Estado final','Estado actual'],
  estado3: ['Estado.3','Estado final','Estado detalle','Estado gestion','Estado gestión','Estado Gestión'],
  mac: ['MAC','Mac','Mac Address','Dirección MAC','Direccion MAC','MAC Address']
};

function pickFirstValue(obj, keys){
  if (!obj || !keys) return '';
  for (const key of keys){ if (!key) continue; const value = obj[key]; if (value !== null && value !== undefined){ const str = String(value).trim(); if (str) return str; } }
  return '';
}
function buildUbicacion(order){
  const calle = pickFirstValue(order, ORDER_FIELD_KEYS.ubicacionCalle); const altura = pickFirstValue(order, ORDER_FIELD_KEYS.ubicacionAltura); const localidad = pickFirstValue(order, ORDER_FIELD_KEYS.ubicacionLocalidad);
  const parts = []; if (calle && altura){ parts.push(`${calle} ${altura}`.trim()); } else if (calle){ parts.push(calle); } else if (altura){ parts.push(altura); } if (localidad){ parts.push(localidad); } return parts.join(', ');
}
function buildOrderExportRow(order, zoneInfo){
  if (!order) return null; const meta = order.__meta || {};
  let fecha = pickFirstValue(order, ORDER_FIELD_KEYS.fecha); if (!fecha && meta.fecha){ fecha = DateUtils.format(meta.fecha); }
  const zonaHFC = meta.zonaHFC || zoneInfo?.zonaHFC || pickFirstValue(order, ORDER_FIELD_KEYS.zonaHFC) || zoneInfo?.zona || '';
  const zonaFTTH = meta.zonaFTTH || zoneInfo?.zonaFTTH || pickFirstValue(order, ORDER_FIELD_KEYS.zonaFTTH);
  const territorio = meta.territorio || pickFirstValue(order, ORDER_FIELD_KEYS.territorio);
  const ubicacion = buildUbicacion(order);
  const caso = meta.numeroCaso || pickFirstValue(order, ORDER_FIELD_KEYS.caso);
  const numeroOrden = pickFirstValue(order, ORDER_FIELD_KEYS.numeroOrden);
  const numeroOTuca = pickFirstValue(order, ORDER_FIELD_KEYS.numeroOTuca);
  const diagnostico = pickFirstValue(order, ORDER_FIELD_KEYS.diagnostico);
  const tipo = pickFirstValue(order, ORDER_FIELD_KEYS.tipo) || zoneInfo?.tipo || '';
  const tipoTrabajo = pickFirstValue(order, ORDER_FIELD_KEYS.tipoTrabajo);
  const estado1 = pickFirstValue(order, ORDER_FIELD_KEYS.estado1);
  let estado2 = pickFirstValue(order, ORDER_FIELD_KEYS.estado2);
  let estado3 = pickFirstValue(order, ORDER_FIELD_KEYS.estado3);
  if (estado3 && estado2 && estado3 === estado2){ const alternativas = ['Estado final', 'Estado gestión', 'Estado Gestion', 'Estado detalle']; const altern = pickFirstValue(order, alternativas); if (altern && altern !== estado2){ estado3 = altern; } }
  let mac = pickFirstValue(order, ORDER_FIELD_KEYS.mac); if (!mac && Array.isArray(meta.dispositivos) && meta.dispositivos.length){ mac = String(meta.dispositivos[0].macAddress || meta.dispositivos[0].mac || '').trim(); }
  return { Fecha: fecha || '', ZonaHFC: zonaHFC || '', ZonaFTTH: zonaFTTH || '', Territorio: territorio || '', Ubicacion: ubicacion || '', Caso: caso || '', NumeroOrden: numeroOrden || '', NumeroOTuca: numeroOTuca || '', Diagnostico: diagnostico || '', Tipo: tipo || '', TipoTrabajo: tipoTrabajo || '', Estado1: estado1 || '', Estado2: estado2 || '', Estado3: estado3 || '', MAC: mac || '' };
}
function buildZoneExportRows(zoneData){
  if (!zoneData) return []; const source = zoneData.ordenesOriginales || zoneData.ordenes || [];
  return source.map(order => buildOrderExportRow(order, zoneData)).filter(row => row && ZONE_EXPORT_HEADERS.some(header => (row[header] || '').toString().trim().length));
}
function createWorksheetFromRows(rows, headers){ if (!rows || !rows.length) return null; const data = rows.map(row => headers.map(header => row[header] || '')); return XLSX.utils.aoa_to_sheet([headers, ...data]); }
function sanitizeSheetName(name){ const fallback = 'Hoja'; if (!name) return fallback; const invalidChars = /[\/?*:[\]]/g; const cleaned = name.toString().replace(invalidChars, ' ').replace(/[\u0000-\u001f]/g, ' ').trim(); const truncated = cleaned.substring(0, 31); return truncated || fallback; }
function appendSheet(workbook, worksheet, desiredName, usedNames){
  if (!worksheet) return false; const base = sanitizeSheetName(desiredName); let name = base; let counter = 1;
  while (usedNames.has(name)){ counter += 1; const suffix = `_${counter}`; const baseTrim = base.substring(0, Math.max(31 - suffix.length, 1)); name = `${baseTrim}${suffix}`; }
  usedNames.add(name); XLSX.utils.book_append_sheet(workbook, worksheet, name); return true;
}
function exportEquiposGrupoExcel(zona, useFiltered = false){
  const source = useFiltered && window.equiposPorZona ? window.equiposPorZona : window.equiposPorZonaCompleto;
  if (!source) return toast('No hay datos de equipos');
  const arr = source.get(zona) || []; if (!arr.length) return toast('No hay equipos en esa zona');
  const wb = XLSX.utils.book_new(); const ws = XLSX.utils.json_to_sheet(arr); XLSX.utils.book_append_sheet(wb, ws, `Equipos_${zona || 'NA'}`);
  const filterInfo = useFiltered && (Filters.equipoModelo.length > 0 || Filters.equipoMarca || Filters.equipoTerritorio) ? `_filtrado` : '';
  XLSX.writeFile(wb, `Equipos_${zona || 'NA'}${filterInfo}_${new Date().toISOString().slice(0, 10)}.xlsx`); toast(`✅ Exportados ${arr.length} equipos de ${zona}`);
}
function exportZonaExcel(zoneIdx) {
  const zonaData = window.currentAnalyzedZones?.[zoneIdx]; if (!zonaData) return toast('No hay datos de la zona');
  const rows = buildZoneExportRows(zonaData); if (!rows.length) { toast('No hay órdenes para exportar en esta zona'); return; }
  const wb = XLSX.utils.book_new(); const usedNames = new Set(); const sheet = createWorksheetFromRows(rows, ZONE_EXPORT_HEADERS); appendSheet(wb, sheet, `Zona_${zonaData.zona}`, usedNames);
  const fecha = new Date().toISOString().slice(0, 10); XLSX.writeFile(wb, `Zona_${zonaData.zona}_${fecha}.xlsx`); toast(`✅ Exportada zona ${zonaData.zona} (${rows.length} órdenes)`);
}
function exportCMTSExcel(cmts) {
  const cmtsData = (window.currentCMTSData || []).find(c => c.cmts === cmts); if (!cmtsData) return toast('No hay datos del CMTS');
  const wb = XLSX.utils.book_new(); const zonasFlat = cmtsData.zonas.map(z => ({ Zona: z.zona, Tipo: z.tipo, Total_OTs: z.totalOTs, Ingreso_N: z.ingresoN, Ingreso_N1: z.ingresoN1, Estado_Nodo: z.nodoEstado }));
  const ws = XLSX.utils.json_to_sheet(zonasFlat); XLSX.utils.book_append_sheet(wb, ws, `CMTS_${cmts.slice(0, 20)}`);
  XLSX.writeFile(wb, `CMTS_${cmts.slice(0, 20)}_${new Date().toISOString().slice(0, 10)}.xlsx`); toast(`✅ Exportado CMTS ${cmts}`);
}
function exportExcelVista(){
  const hasZones = Array.isArray(window.currentAnalyzedZones) && window.currentAnalyzedZones.length; const hasOrders = Array.isArray(window.lastFilteredOrders) && window.lastFilteredOrders.length;
  if (!hasZones && !hasOrders) { toast('No hay datos para exportar'); return; }
  const wb = XLSX.utils.book_new(); const usedNames = new Set();
  if (hasZones){
    const zonasData = window.currentAnalyzedZones.map(z => ({ Zona: z.zona, Tipo: z.tipo, Red: z.tipo === 'FTTH' ? z.zonaHFC : '', CMTS: z.cmts, Tiene_Alarma: z.tieneAlarma ? 'SÍ' : 'NO', Alarmas_Activas: z.alarmasActivas, Total_OTs: z.totalOTs, Ingreso_N: z.ingresoN, Ingreso_N1: z.ingresoN1, Max_Dia: z.maxDia }));
    const zonasSheet = XLSX.utils.json_to_sheet(zonasData); appendSheet(wb, zonasSheet, 'Zonas', usedNames);
  }
  if (Array.isArray(window.currentCMTSData) && window.currentCMTSData.length){
    const cmtsData = window.currentCMTSData.map(c => ({ CMTS: c.cmts, Zonas: c.zonas.length, Total_OTs: c.totalOTs, Zonas_UP: c.zonasUp, Zonas_DOWN: c.zonasDown, Zonas_Criticas: c.zonasCriticas }));
    const cmtsSheet = XLSX.utils.json_to_sheet(cmtsData); appendSheet(wb, cmtsSheet, 'CMTS', usedNames);
  }
  if (window.edificiosData && window.edificiosData.length){
    const edi = window.edificiosData.map(e => ({ Direccion: e.direccion, Zona: e.zona, Territorio: e.territorio, Total_OTs: e.casos.length }));
    const ediSheet = XLSX.utils.json_to_sheet(edi); appendSheet(wb, ediSheet, 'Edificios', usedNames);
  }
  if (window.equiposPorZona && window.equiposPorZona.size){
    const todos = []; window.equiposPorZona.forEach((arr, zona) => { arr.forEach(it => todos.push({ ...it, zona })); });
    const equiposSheet = XLSX.utils.json_to_sheet(todos); appendSheet(wb, equiposSheet, 'Equipos', usedNames);
  }
  if (hasZones){
    const detalleRows = []; window.currentAnalyzedZones.forEach(z => { const rows = buildZoneExportRows(z); if (rows && rows.length) { detalleRows.push(...rows); } });
    const detalleSheet = createWorksheetFromRows(detalleRows, ZONE_EXPORT_HEADERS); appendSheet(wb, detalleSheet, 'Detalle_Zonas', usedNames);
  }
  if (!wb.SheetNames || wb.SheetNames.length === 0){ toast('No hay datos para exportar'); return; }
  const fecha = new Date().toISOString().slice(0, 10); XLSX.writeFile(wb, `Vista_Filtrada_${fecha}.xlsx`); toast('✓ Vista detallada exportada');
}
function exportExcelZonasCrudo(){
  if (!Array.isArray(window.currentAnalyzedZones) || !window.currentAnalyzedZones.length) { return toast('No hay zonas para exportar'); }
  const wb = XLSX.utils.book_new();
  const zonasData = window.currentAnalyzedZones.map(z => ({ Zona: z.zona, Tipo: z.tipo, Red: z.tipo === 'FTTH' ? z.zonaHFC : '', CMTS: z.cmts, Nodo_Estado: z.nodoEstado, Nodo_UP: z.nodoUp, Nodo_DOWN: z.nodoDown, Tiene_Alarma: z.tieneAlarma ? 'SÍ' : 'NO', Alarmas_Activas: z.alarmasActivas, Territorio: z.territorio, Total_OTs: z.totalOTs, Ingreso_N: z.ingresoN, Ingreso_N1: z.ingresoN1, Max_Dia: z.maxDia }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(zonasData), 'Zonas');
  const fecha = new Date().toISOString().slice(0, 10); XLSX.writeFile(wb, `Zonas_Crudo_${fecha}.xlsx`); toast('✓ Zonas exportadas');
}
function exportModalDetalleExcel(){
  const title = document.getElementById('modalTitle').textContent || ''; const wb = XLSX.utils.book_new();
  if (title.startsWith('Detalle: ') && window.currentZone){ const ordenes = window.currentZone.ordenes || []; XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ordenes), `Zona_${window.currentZone.zona || 'NA'}`); }
  else if (title.startsWith('Zonas del CMTS: ')){
    const cmts = title.replace('Zonas del CMTS: ', '').trim(); const data = (window.currentCMTSData || []).find(c => c.cmts === cmts); if (!data) return toast('Sin datos de CMTS');
    const zonasFlat = data.zonas.map(z => ({ Zona: z.zona, Tipo: z.tipo, Total_OTs: z.totalOTs, Ingreso_N: z.ingresoN, Ingreso_N1: z.ingresoN1, Estado_Nodo: z.nodoEstado }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(zonasFlat), `CMTS_${cmts.slice(0, 20)}`);
  } else {
    const edTitle = document.getElementById('edificioModalTitle')?.textContent || '';
    if (edTitle.startsWith('🏢 Edificio: ') && window.edificiosData){ const dir = edTitle.replace('🏢 Edificio: ', '').trim(); const e = window.edificiosData.find(x => x.direccion === dir); if (!e) return toast('Sin datos del edificio'); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(e.casos), 'Edificio_OTs'); }
    else { return toast('No hay detalle para exportar'); }
  }
  const fecha = new Date().toISOString().slice(0, 10); XLSX.writeFile(wb, `Detalle_${fecha}.xlsx`); toast('✓ Detalle exportado');
}

// === CONTROLADOR PRINCIPAL Y EVENTOS ===
let currentData = null;
let allZones = [];
let allCMTS = [];
let currentZone = null;
let selectedOrders = new Set();

document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
});

function setupEventListeners() {
  on('fileConsolidado1', 'change', e => loadFile(e, 1));
  on('fileConsolidado2', 'change', e => loadFile(e, 2));
  on('fileNodos', 'change', e => loadFile(e, 3));
  on('fileFMS', 'change', e => loadFile(e, 4));
  
  const filterCatec = document.getElementById('filterCATEC');
  const filterExcludeCatec = document.getElementById('filterExcludeCATEC');
  if (filterCatec && filterExcludeCatec) {
    filterCatec.addEventListener('change', e => { if (e.target.checked) filterExcludeCatec.checked = false; applyFilters(); });
    filterExcludeCatec.addEventListener('change', e => { if (e.target.checked) filterCatec.checked = false; applyFilters(); });
  }
  
  on('showAllStates', 'change', applyFilters);
  on('filterFTTH', 'change', e => { if (e.target.checked) { const el = document.getElementById('filterExcludeFTTH'); if (el) el.checked = false; } applyFilters(); });
  on('filterExcludeFTTH', 'change', e => { if (e.target.checked) { const el = document.getElementById('filterFTTH'); if (el) el.checked = false; } applyFilters(); });
  
  on('filterNodoEstado', 'change', applyFilters);
  on('filterCMTS', 'change', applyFilters);
  on('daysWindow', 'change', applyFilters);
  on('filterTerritorio', 'change', applyFilters);
  on('filterSistema', 'change', applyFilters);
  on('filterAlarma', 'change', applyFilters);
  on('quickSearch', 'input', debounce(applyFilters, 300));
  on('ordenarPorIngreso', 'change', applyFilters);
  on('zoneFilterSearch', 'input', onZoneFilterSearch);
  on('zoneFilterOptions', 'change', onZoneOptionChange);

  document.addEventListener('click', handleZoneFilterOutsideClick);
  console.log('✅ Event listeners configurados');
}

async function loadFile(e, tipo) {
  const file = e.target.files[0]; if (!file) return;
  const statusEl = document.getElementById(`status${tipo}`);
  if (statusEl) { statusEl.textContent = 'Cargando...'; statusEl.classList.remove('loaded'); }
  console.log(`📁 Cargando archivo: ${file.name} (${(file.size / 1024).toFixed(1)} KB) - Tipo: ${tipo}`);
  
  let result;
  try {
    if (tipo === 4 && file.name.endsWith('.csv')) { result = await dataProcessor.loadCSV(file); }
    else { result = await dataProcessor.loadExcel(file, tipo); }
    
    if (result.success) {
      if (statusEl) { statusEl.textContent = `✓ ${result.rows} filas cargadas`; statusEl.classList.add('loaded'); }
      const nombres = ['Consolidado 1', 'Consolidado 2', 'Nodos UP/DOWN', 'Alarmas FMS'];
      toast(`${nombres[tipo - 1]} cargado: ${result.rows} registros`);
      console.log(`✅ ${nombres[tipo - 1]} cargado exitosamente: ${result.rows} registros`);
      
      if ((dataProcessor.consolidado1 || dataProcessor.consolidado2)) {
        const mergeStatus = document.getElementById('mergeStatus'); if (mergeStatus) mergeStatus.style.display = 'flex'; processData();
      }
    } else { if (statusEl) statusEl.textContent = `✗ Error`; toast(`Error al cargar archivo: ${result.error}`); }
  } catch (error) { console.error(`❌ Error en loadFile:`, error); if (statusEl) statusEl.textContent = `✗ Error`; toast(`Error inesperado: ${error.message}`); }
}

function processData() {
  const merged = dataProcessor.merge(); if (!merged.length) return;
  const daysWindow = parseInt(document.getElementById('daysWindow').value);
  
  const zones = dataProcessor.processZones(merged);
  allZones = dataProcessor.analyzeZones(zones, daysWindow);
  allCMTS = dataProcessor.analyzeCMTS(allZones);
  
  const territoriosAnalisis = dataProcessor.analyzeTerritorios(allZones);
  window.territoriosData = territoriosAnalisis;
  
  currentData = { ordenes: merged, zonas: allZones };
  
  populateFilters();
  
  const stats = {
    total: allZones.reduce((sum, z) => sum + z.totalOTs, 0),
    zonas: allZones.length,
    territoriosCriticos: territoriosAnalisis.filter(t => t.esCritico).length,
    ftth: allZones.filter(z => z.tipo === 'FTTH').length,
    conAlarmas: allZones.filter(z => z.tieneAlarma).length,
    nodosCriticos: allZones.filter(z => z.nodoEstado === 'critical').length
  };
  
  if (window.UIRenderer) document.getElementById('statsGrid').innerHTML = UIRenderer.renderStats(stats);
  document.getElementById('btnExportExcel').disabled = false; document.getElementById('btnExportExcelZonas').disabled = false;
  
  let statusText = `✓ ${merged.length} órdenes procesadas (deduplicadas)`;
  if (dataProcessor.nodosData) { statusText += ` • ${dataProcessor.nodosMap.size} nodos integrados`; }
  if (dataProcessor.fmsData) { statusText += ` • ${dataProcessor.fmsMap.size} zonas con alarmas`; }
  statusText += ` • Ventana: ${daysWindow} días`;
  document.getElementById('mergeStatusText').textContent = statusText;
  
  applyFilters();
}

function populateFilters() {
  const territorios = [...new Set(currentData.ordenes.map(o => o['Territorio de servicio: Nombre']).filter(Boolean))];
  document.getElementById('filterTerritorio').innerHTML = '<option value="">Todos</option>' + territorios.sort().map(t => `<option value="${t}">${t}</option>`).join('');

  const cmtsList = [...new Set(allZones.map(z => z.cmts).filter(Boolean))];
  document.getElementById('filterCMTS').innerHTML = '<option value="">Todos</option>' + cmtsList.sort().map(c => `<option value="${c}">${c}</option>`).join('');

  const zoneNames = [...new Set((currentData.zonas || []).map(z => z.zona).filter(Boolean))];
  initializeZoneFilterOptions(zoneNames);
}

function applyFilters() {
  if (!currentData) return;
  
  // Sincronizar UI a Filtros
  const catecEl = document.getElementById('filterCATEC');
  const excludeCatecEl = document.getElementById('filterExcludeCATEC');
  Filters.catec = catecEl ? catecEl.checked : false;
  Filters.excludeCatec = excludeCatecEl ? excludeCatecEl.checked : false;
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
  Filters.selectedZonas = Array.from(zoneFilterState.selected);

  // 1. Filtrar Órdenes base
  const filtered = Filters.apply(currentData.ordenes);
  
  // 2. Reprocesar zonas con órdenes filtradas
  const zones = dataProcessor.processZones(filtered);
  let analyzed = dataProcessor.analyzeZones(zones, Filters.days);
  
  // 3. Aplicar filtros directos sobre zonas (ej. estado nodo)
  analyzed = Filters.applyToZones(analyzed);
  
  const cmtsFiltered = dataProcessor.analyzeCMTS(analyzed);
  const territoriosAnalisis = dataProcessor.analyzeTerritorios(analyzed);
  window.territoriosData = territoriosAnalisis;
  
  const stats = {
    total: analyzed.reduce((sum, z) => sum + z.totalOTs, 0),
    zonas: analyzed.length,
    territoriosCriticos: territoriosAnalisis.filter(t => t.esCritico).length,
    ftth: analyzed.filter(z => z.tipo === 'FTTH').length,
    conAlarmas: analyzed.filter(z => z.tieneAlarma).length,
    nodosCriticos: analyzed.filter(z => z.nodoEstado === 'critical').length
  };
  
  // Actualizar Referencias globales para exports
  window.lastFilteredOrders = filtered;
  window.currentAnalyzedZones = analyzed;
  window.currentCMTSData = cmtsFiltered;
  allZones = analyzed;
  allCMTS = cmtsFiltered;
  
  // Renderizar UI si UIRenderer está disponible
  if (window.UIRenderer) {
      document.getElementById('statsGrid').innerHTML = UIRenderer.renderStats(stats);
      document.getElementById('zonasPanel').innerHTML = UIRenderer.renderZonas(analyzed);
      document.getElementById('cmtsPanel').innerHTML = UIRenderer.renderCMTS(cmtsFiltered);
      document.getElementById('edificiosPanel').innerHTML = UIRenderer.renderEdificios(filtered);
      document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(filtered);

      // === RENDERIZADO FMS (Conectado) ===
      const allAlarmas = analyzed.flatMap(z => (z.alarmas || []).map(a => ({ ...a, zonaTecnica: z.zona })));
      let alarmasFiltradas = allAlarmas;
      if (Filters.alarma === 'con-alarma') alarmasFiltradas = allAlarmas.filter(a => a.isActive);
      else if (Filters.alarma === 'sin-alarma') alarmasFiltradas = [];
      document.getElementById('fmsPanel').innerHTML = UIRenderer.renderFMS(alarmasFiltradas);
  } else {
      console.error("❌ UIRenderer no encontrado. Verifica que ui-renderer.js esté cargado.");
  }

  document.getElementById('btnExportExcel').disabled = analyzed.length === 0 && filtered.length === 0;
  document.getElementById('btnExportExcelZonas').disabled = analyzed.length === 0;
}

function resetFiltersState() {
  const catecEl = document.getElementById('filterCATEC'); const excludeCatecEl = document.getElementById('filterExcludeCATEC');
  if (catecEl) catecEl.checked = false; if (excludeCatecEl) excludeCatecEl.checked = false;
  document.getElementById('showAllStates').checked = false; document.getElementById('filterFTTH').checked = false; document.getElementById('filterExcludeFTTH').checked = false;
  document.getElementById('filterNodoEstado').value = ''; document.getElementById('filterCMTS').value = ''; document.getElementById('daysWindow').value = '3';
  document.getElementById('filterTerritorio').value = ''; document.getElementById('filterSistema').value = ''; document.getElementById('filterAlarma').value = '';
  document.getElementById('quickSearch').value = ''; document.getElementById('ordenarPorIngreso').value = 'desc';
  Filters.criticidad = ''; resetZoneFilterState(); closeZoneFilter();
}
function clearFilters() { resetFiltersState(); applyFilters(); }

function filterByStat(statType) {
  resetFiltersState();
  switch (statType) {
    case 'total': case 'zonas': toast('🔄 Mostrando todas las zonas (filtros reseteados)'); break;
    case 'territoriosCriticos':
      if (!window.territoriosData) { toast('❌ No hay datos de territorios disponibles'); break; }
      const territoriosCriticos = window.territoriosData.filter(t => t.esCritico);
      if (territoriosCriticos.length === 0) { toast('✅ No hay territorios críticos en este momento'); break; }
      showTerritoriosCriticosModal(territoriosCriticos); return;
    case 'ftth': document.getElementById('filterFTTH').checked = true; document.getElementById('filterExcludeFTTH').checked = false; toast('🔌 Mostrando solo zonas FTTH'); break;
    case 'alarmas': document.getElementById('filterAlarma').value = 'con-alarma'; toast('🚨 Mostrando zonas con alarmas activas'); break;
    case 'nodosCriticos': document.getElementById('filterNodoEstado').value = 'critical'; toast('🟥 Mostrando zonas con NODOS CRÍTICOS'); break;
  }
  applyFilters();
}

// === MODALES Y PANELES EXTRA (Utilizan UIRenderer si es necesario, o lógica directa) ===

function showTerritoriosCriticosModal(territoriosCriticos) {
  let html = '<div class="territorios-criticos-container"><h3 style="color: #D13438; margin-bottom: 20px;">🚨 Territorios Críticos Detectados</h3>';
  territoriosCriticos.forEach((t, idx) => {
    html += `<div class="territorio-critico-card" style="background: white; border: 2px solid #D13438; border-radius: 8px; padding: 16px; margin-bottom: 16px; box-shadow: 0 4px 8px rgba(209, 52, 56, 0.15);">`;
    html += `<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;"><h4 style="color: #0078D4; margin: 0; font-size: 1.1rem;">${t.territorio}</h4><span class="badge badge-critical" style="font-size: 0.85rem;">CRÍTICO</span></div>`;
    html += `<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin-bottom: 12px;"><div><strong>Total OTs:</strong> ${t.totalOTs}</div><div><strong>Zonas Críticas:</strong> ${t.zonasCriticas}</div><div><strong>Con Alarmas:</strong> ${t.zonasConAlarma}</div><div><strong>Total Zonas:</strong> ${t.zonas.length}</div></div>`;
    if (t.flexSubdivisiones.size > 0) {
      html += `<div style="margin-top: 12px; padding-top: 12px; border-top: 1px solid #E1DFDD;"><strong style="color: #605E5C; font-size: 0.9rem;">📍 Subdivisiones:</strong><div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)); gap: 8px; margin-top: 8px;">`;
      Array.from(t.flexSubdivisiones.values()).forEach(flex => { const badgeColor = flex.zonasCriticas > 0 ? '#D13438' : '#107C10'; html += `<div style="background: ${badgeColor}15; border: 1px solid ${badgeColor}; border-radius: 6px; padding: 8px; text-align: center;"><div style="font-weight: 600; color: ${badgeColor};">${flex.nombre}</div><div style="font-size: 0.75rem; color: #605E5C;">${flex.zonas.length} zonas</div><div style="font-size: 0.75rem; color: #605E5C;">${flex.totalOTs} OTs</div>${flex.zonasCriticas > 0 ? `<div style="font-size: 0.75rem; color: ${badgeColor};">⚠️ ${flex.zonasCriticas} críticas</div>` : ''}</div>`; }); html += `</div></div>`;
    }
    html += `<div style="margin-top: 12px;"><button class="btn btn-primary" style="font-size: 0.85rem; padding: 8px 16px;" onclick="filtrarPorTerritorio('${t.territorio.replace(/'/g, "\\'")}')">🔍 Ver Zonas de ${t.territorio}</button></div></div>`;
  });
  html += `<div style="text-align: center; margin-top: 20px;"><button class="btn btn-secondary" onclick="closeTerritoriosCriticosModal()">Cerrar</button></div></div>`;
  document.getElementById('modalTitle').textContent = `📊 Territorios Críticos (${territoriosCriticos.length})`; document.getElementById('modalBody').innerHTML = html; document.getElementById('modalFilters').style.display = 'none'; document.getElementById('modalFooter').innerHTML = ''; document.getElementById('modalBackdrop').classList.add('show'); document.body.classList.add('modal-open');
}
function closeTerritoriosCriticosModal() { closeModal(); }
function filtrarPorTerritorio(territorioNormalizado) {
  closeModal(); const territoriosOriginales = [...new Set(currentData.ordenes.map(o => o['Territorio de servicio: Nombre']).filter(t => t && TerritorioUtils.normalizar(t) === territorioNormalizado))];
  if (territoriosOriginales.length === 0) { toast('❌ No se encontraron zonas para este territorio'); return; }
  document.getElementById('filterTerritorio').value = territoriosOriginales[0]; toast(`🔍 Filtrando por territorio: ${territorioNormalizado}`); applyFilters(); switchTab('zonas');
}

function switchTab(tabName) {
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active')); document.querySelectorAll('.tab-panel').forEach(panel => panel.classList.remove('active'));
  document.querySelector(`[data-tab="${tabName}"]`).classList.add('active'); document.getElementById(`${tabName}Panel`).classList.add('active');
}

// ... Resto de funciones de modal (showAlarmaInfo, showEdificioDetail, openModal, etc.)
// Como estas funciones usan lógica de renderizado interno específico para modales y datos globales, pueden quedarse aquí
// o moverse a UI Renderer si se desea refactorizar más a fondo. Por ahora, las mantendremos aquí para minimizar cambios riesgosos.

function showAlarmaInfo(zoneIdx) {
  const zonaData = window.currentAnalyzedZones[zoneIdx]; if (!zonaData || !zonaData.alarmas.length) return;
  let html = '<div class="alarma-info-box">'; html += `<h4>🚨 Alarmas en Zona: ${zonaData.zona}</h4>`;
  zonaData.alarmas.forEach((a, idx) => {
    html += `<div style="margin-bottom: 20px; padding-bottom: 20px; border-bottom: 1px solid var(--border-subtle);"><h5 style="color: var(--win-blue); margin-bottom: 10px;">Alarma ${idx + 1} ${a.isActive ? '<span class="badge badge-alarma-activa">ACTIVO</span>' : '<span class="badge">CERRADO</span>'}</h5><div class="alarma-info-grid">`;
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
  html += '</div>'; document.getElementById('alarmaModalBody').innerHTML = html; document.getElementById('alarmaBackdrop').classList.add('show'); document.body.classList.add('modal-open');
}
function closeAlarmaModal() { document.getElementById('alarmaBackdrop').classList.remove('show'); document.body.classList.remove('modal-open'); }

function showEdificioDetail(idx) {
  const edificio = window.edificiosData[idx]; if (!edificio) return;
  document.getElementById('edificioModalTitle').textContent = `🏢 Edificio: ${edificio.direccion}`;
  let html = `<div style="margin-bottom: 16px; padding: 12px; background: var(--bg-tertiary); border-radius: 8px;"><strong>Zona:</strong> ${edificio.zona}<br><strong>Territorio:</strong> ${edificio.territorio}<br><strong>Total OTs:</strong> ${edificio.casos.length}</div>`;
  html += '<div class="table-container"><div class="table-wrapper"><table class="detail-table"><thead><tr><th>Número de Caso</th><th>Estado</th><th>Diagnóstico</th><th>Fecha Creación</th></tr></thead><tbody>';
  edificio.casos.forEach(o => {
    const numCaso = o['Número del caso'] || o['Caso Externo'] || ''; const sistema = TextUtils.detectarSistema(numCaso); let badgeSistema = ''; if (sistema === 'OPEN') { badgeSistema = '<span class="badge badge-open">OPEN</span>'; } else if (sistema === 'FAN') { badgeSistema = '<span class="badge badge-fan">FAN</span>'; }
    const estado = o['Estado'] || o['Estado.1'] || o['Estado.2'] || ''; const diag = o['Diagnostico Tecnico'] || o['Diagnóstico Técnico'] || '-'; const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || '';
    html += `<tr><td>${numCaso} ${badgeSistema}</td><td>${estado}</td><td>${diag.substring(0, 100)}${diag.length > 100 ? '...' : ''}</td><td>${fecha}</td></tr>`;
  });
  html += '</tbody></table></div></div>'; document.getElementById('edificioModalBody').innerHTML = html; document.getElementById('edificioBackdrop').classList.add('show'); document.body.classList.add('modal-open');
}
function closeEdificioModal() { document.getElementById('edificioBackdrop').classList.remove('show'); document.body.classList.remove('modal-open'); }
function showEdificioAlarmas(idx) {
  const edificio = window.edificiosData[idx]; if (!edificio) return; if (!edificio.alarmasDetalle || edificio.alarmasDetalle.length === 0) { toast('No hay alarmas FMS para este edificio'); return; }
  document.getElementById('alarmaModalBody').innerHTML = '';
  let html = `<div style="margin-bottom: 16px; padding: 12px; background: var(--bg-tertiary); border-radius: 8px;"><strong>🏢 Edificio:</strong> ${edificio.direccion}<br><strong>Zonas:</strong> ${Array.from(edificio.zonas).join(', ')}<br><strong>Total Alarmas:</strong> ${edificio.alarmasTotal} (${edificio.alarmasActivas} activas)</div>`;
  edificio.alarmasDetalle.forEach(detalle => {
    html += `<div style="margin-bottom: 16px;"><h4 style="margin: 0 0 8px 0; color: var(--text-primary);">📍 Zona: ${detalle.zona} (${detalle.activas} activas de ${detalle.alarmas.length})</h4><div class="table-container"><div class="table-wrapper"><table class="detail-table"><thead><tr><th>ID</th><th>Tipo</th><th>Elemento</th><th>Daño</th><th>Descripción</th><th>Fecha Creación</th><th>Estado</th></tr></thead><tbody>`;
    detalle.alarmas.forEach(alarma => {
      const estadoClass = alarma.isActive ? 'status-critical' : 'status-ok'; const estadoText = alarma.isActive ? '🔴 Activa' : '🟢 Recuperada';
      html += `<tr><td>${alarma.eventId || '-'}</td><td>${alarma.type || alarma.elementType || '-'}</td><td>${alarma.elementCode || '-'}</td><td>${alarma.damage || alarma.damageClassification || '-'}</td><td title="${alarma.description || ''}">${(alarma.description || '-').substring(0, 50)}${(alarma.description || '').length > 50 ? '...' : ''}</td><td>${alarma.creationDate || '-'}</td><td class="${estadoClass}">${estadoText}</td></tr>`;
    });
    html += '</tbody></table></div></div></div>';
  });
  document.getElementById('alarmaModalBody').innerHTML = html; document.getElementById('alarmaBackdrop').classList.add('show'); document.body.classList.add('modal-open');
}

function openModal(zoneIdx) {
  const zonaData = window.currentAnalyzedZones[zoneIdx]; if (!zonaData) return;
  currentZone = zonaData; selectedOrders.clear(); document.getElementById('modalTitle').textContent = `Detalle: ${currentZone.zona}`;
  const ordenesParaModal = currentZone.ordenesOriginales || currentZone.ordenes; currentZone.ordenes = ordenesParaModal;
  const dias = [...new Set(ordenesParaModal.map(o => { const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || o['Fecha de inicio']; const dt = DateUtils.parse(fecha); return dt ? DateUtils.format(dt) : null; }).filter(Boolean))].sort();
  document.getElementById('filterDia').innerHTML = '<option value="">Todos los días</option>' + dias.map(d => `<option value="${d}">${d}</option>`).join('');
  renderModalContent(); document.getElementById('modalBackdrop').classList.add('show'); document.body.classList.add('modal-open');
}
function openCMTSDetail(cmts) {
  const cmtsDataSource = window.currentCMTSData || allCMTS; if (!cmtsDataSource || !cmtsDataSource.length) { toast('❌ No hay datos de CMTS disponibles'); return; }
  const cmtsData = cmtsDataSource.find(c => c.cmts === cmts); if (!cmtsData) { toast('❌ No se encontró el CMTS: ' + cmts); return; }
  let html = '<div class="table-container"><div class="table-wrapper"><table><thead><tr><th>Zona</th><th>Tipo</th><th>Total OTs</th><th>N</th><th>N-1</th><th>Estado Nodo</th></tr></thead><tbody>';
  cmtsData.zonas.forEach(z => {
    const badgeTipo = z.tipo === 'FTTH' ? '<span class="badge badge-ftth">FTTH</span>' : '<span class="badge badge-hfc">HFC</span>';
    let badgeNodo = '<span class="badge">Sin datos</span>'; if (z.nodoEstado === 'up') badgeNodo = `<span class="badge badge-up">✓ UP</span>`; else if (z.nodoEstado === 'critical') badgeNodo = `<span class="badge badge-critical">⚠ CRÍTICO</span>`; else if (z.nodoEstado === 'down') badgeNodo = `<span class="badge badge-down">↓ DOWN</span>`;
    html += `<tr><td><strong>${z.zona}</strong></td><td>${badgeTipo}</td><td class="number">${z.totalOTs}</td><td class="number">${z.ingresoN}</td><td class="number">${z.ingresoN1}</td><td>${badgeNodo}</td></tr>`;
  });
  html += '</tbody></table></div></div>'; document.getElementById('modalTitle').textContent = `Zonas del CMTS: ${cmts}`; document.getElementById('modalBody').innerHTML = html; document.getElementById('modalFilters').style.display = 'none';
  document.getElementById('modalFooter').innerHTML = `<div></div><div style="display:flex; gap:8px;"><button class="btn btn-primary" onclick="exportModalDetalleExcel()">📥 Exportar detalle</button><button class="btn btn-secondary" onclick="closeModal()">Cerrar</button></div>`;
  document.getElementById('modalBackdrop').classList.add('show'); document.body.classList.add('modal-open');
}
function closeModal() {
  document.getElementById('modalBackdrop').classList.remove('show'); document.body.classList.remove('modal-open'); selectedOrders.clear(); document.getElementById('filterHorario').value = ''; document.getElementById('filterDia').value = '';
  document.getElementById('modalFilters').style.display = 'flex'; document.getElementById('modalFooter').innerHTML = `<div class="selection-info" id="selectionInfo">0 órdenes seleccionadas</div><div style="display: flex; gap: 8px;"><button class="btn btn-warning" id="btnExportBEFAN" disabled onclick="exportBEFAN()">📤 Exportar BEFAN (TXT)</button><button class="btn btn-secondary" onclick="closeModal()">Cerrar</button></div>`;
}
function applyModalFilters() { renderModalContent(); }
function renderModalContent() {
  const horario = document.getElementById('filterHorario').value; const dia = document.getElementById('filterDia').value; let ordenes = currentZone.ordenes.slice();
  if (horario) { const [start, end] = horario.split('-').map(Number); ordenes = ordenes.filter(o => { const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || o['Fecha de inicio']; const dt = DateUtils.parse(fecha); if (!dt) return false; const hour = DateUtils.getHour(dt); return hour >= start && hour < end; }); }
  if (dia) { ordenes = ordenes.filter(o => { const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || o['Fecha de inicio']; const dt = DateUtils.parse(fecha); return dt && DateUtils.format(dt) === dia; }); }
  const chartCounts = UIRenderer.normalizeCounts(currentZone.last7DaysCounts); const chartLabels = Array.isArray(currentZone.last7Days) ? currentZone.last7Days : []; const maxChartValue = Math.max(...chartCounts, 1);
  let html = '<div class="chart-container"><div class="chart-title">📊 Distribución de Ingresos (7 días)</div>';
  if (!chartCounts.length) { html += '<div class="sparkline-placeholder" style="height: 150px; display: flex; align-items: center; justify-content: center;">Sin datos</div>'; } else { html += '<div style="display: flex; justify-content: space-around; align-items: flex-end; height: 150px; padding: 10px;">'; chartCounts.forEach((count, i) => { const barHeight = maxChartValue > 0 ? (count / maxChartValue) * 120 : 0; const label = chartLabels[i] || ''; html += `<div style="text-align: center;"><div style="width: 60px; height: ${barHeight}px; background: linear-gradient(180deg, #0078D4 0%, #005A9E 100%); border-radius: 4px 4px 0 0; margin: 0 auto;"></div><div style="font-size: 0.75rem; font-weight: 700; margin-top: 4px;">${count}</div><div style="font-size: 0.6875rem; color: var(--text-secondary);">${label}</div></div>`; }); html += '</div>'; } html += '</div>';
  html += `<div style="margin-bottom: 16px; padding: 12px; background: var(--bg-tertiary); border-radius: 8px;"><strong>Total órdenes mostradas:</strong> ${ordenes.length} de ${currentZone.ordenes.length}</div>`;
  html += '<div class="table-container"><div style="max-height: 400px; overflow-y: auto;"><table class="detail-table"><thead><tr>';
  html += '<th style="position: sticky; top: 0; z-index: 10;"><input type="checkbox" id="selectAllCheckbox" onchange="toggleSelectAll()"></th>';
  const allCols = new Set(); ordenes.forEach(o => { Object.keys(o).forEach(k => { if (!k.startsWith('_')) allCols.add(k); }); });
  const columns = Array.from(allCols); columns.forEach(col => { html += `<th style="position: sticky; top: 0; z-index: 10; white-space: nowrap;">${col}</th>`; }); html += '</tr></thead><tbody>';
  ordenes.forEach((o, idx) => {
    const cita = o['Número de cita'] || `row_${idx}`; const checked = selectedOrders.has(cita) ? 'checked' : '';
    html += `<tr><td><input type="checkbox" class="order-checkbox" data-cita="${cita}" ${checked} onchange="updateSelection()"></td>`;
    columns.forEach(col => { let value = o[col] || ''; if (col.includes('Número del caso') || col.includes('Caso')) { const sistema = TextUtils.detectarSistema(value); if (sistema) { const badgeClass = sistema === 'OPEN' ? 'badge-open' : 'badge-fan'; value = `${value} <span class="badge ${badgeClass}">${sistema}</span>`; } } if (typeof value === 'string' && value.length > 100) { value = value.substring(0, 100) + '...'; } html += `<td title="${value}">${value}</td>`; });
    html += '</tr>';
  });
  html += '</tbody></table></div></div>'; document.getElementById('modalBody').innerHTML = html; updateSelectionInfo();
}

function toggleSelectAll() {
  const checked = document.getElementById('selectAllOrders')?.checked || document.getElementById('selectAllCheckbox')?.checked || false;
  document.querySelectorAll('.order-checkbox').forEach(cb => { cb.checked = checked; const cita = cb.getAttribute('data-cita'); if (checked) { selectedOrders.add(cita); } else { selectedOrders.delete(cita); } });
  updateSelectionInfo();
}
function updateSelection() { selectedOrders.clear(); document.querySelectorAll('.order-checkbox:checked').forEach(cb => { selectedOrders.add(cb.getAttribute('data-cita')); }); updateSelectionInfo(); }
function updateSelectionInfo() { const count = selectedOrders.size; const infoEl = document.getElementById('selectionInfo'); if (infoEl) { infoEl.textContent = `${count} órdenes seleccionadas`; } const btnEl = document.getElementById('btnExportBEFAN'); if (btnEl) { btnEl.disabled = count === 0; } }
function exportBEFAN() {
  if (selectedOrders.size === 0) { toast('No hay órdenes seleccionadas'); return; }
  const ordenes = currentZone.ordenes.filter(o => { const cita = o['Número de cita']; return selectedOrders.has(cita); });
  let txt = ''; let contadorExitoso = 0;
  ordenes.forEach(o => { const casoExterno = o['Caso Externo'] || o['External Case Id'] || o['Número del caso'] || o['Caso externo'] || o['CASO EXTERNO'] || o['external_case_id'] || ''; if (casoExterno && String(casoExterno).trim() !== '') { txt += `${String(casoExterno).trim()}|${CONFIG.codigo_befan}\n`; contadorExitoso++; } });
  if (!txt || contadorExitoso === 0) { toast('⚠️ Las órdenes seleccionadas no tienen Caso Externo válido'); return; }
  try { const blob = new Blob([txt], {type: 'text/plain;charset=utf-8'}); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `BEFAN_${currentZone.zona}_${new Date().toISOString().slice(0, 10)}.txt`; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url); toast(`✅ Exportadas ${contadorExitoso} órdenes a BEFAN`); }
  catch (error) { console.error('Error al descargar BEFAN:', error); toast('❌ Error al descargar archivo BEFAN'); }
}
function debounce(func, wait) { let timeout; return function (...args) { clearTimeout(timeout); timeout = setTimeout(() => func.apply(this, args), wait); }; }
}

{
type: uploaded file
fileName: index.php
fullContent:
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML>

<HEAD>
	<?php

	include "../recursos/recursos.php";
	include "../recursos/encabezado.php";
	session_start();
	$herramienta = 'panel_fulfillment';
	validar_sesion($herramienta);
	headerBasico();
	headerBootstrap(1);
	headerDatatables();
	// headerHightcharts(1);
	$aleatorio = rand(1, 100000);
	log_visitas_a_web($con_w);
	echo "<script type='text/javascript' src='funciones.js?$aleatorio'></script>";
	echo "<script type='text/javascript' src='../recursos/js.js?$aleatorio'></script>";
	echo "<link rel='stylesheet' type='text/css' href='style.css?$aleatorio'>";
	
	// SCRIPTS PANEL FULFILLMENT
	echo "<script src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'></script>";
	echo "<link rel='stylesheet' href='assets/css/styles.css?$aleatorio'>";

	?>
	<TITLE>Panel Fulfillment v5.0</TITLE>
</HEAD>

<body>

	<div class="card border-0">
		<div class="card-heading">
			<nav class="navbar navbar-expand-lg bg-dark navbar-dark " >
				<h4 class="navbar-brand m-0" href="#">📊 Panel Fulfillment v5.0</h4>
				
				<div class="collapse navbar-collapse" id="navbarSupportedContent">
					<ul class="navbar-nav mr-auto">
					</ul>

					<ul class="navbar-nav">
						<li class="nav-item">
							<span class="navbar-text">
								<span class="fa fa-user-circle" aria-hidden="true"></span>
								<?php echo utf8_encode("Bienvenido: " . $_SESSION['user']); ?>
							</span>
						</li>
						<li class="nav-item">
							<a href="../recursos/sesion/desconectar.php?pag=panel_fulfillment" class="nav-link"><span class="fa fa-sign-out" aria-hidden="true"></span>CERRAR SESION</a>
						</li>
					</ul>
				</div>
			</nav>
		</div>
		
		<div>
			<div class="app-container">
    
				<header class="app-header">
					<div class="header-content">
						<div class="header-title">
							<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
								<path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path>
								<polyline points="3.27 6.96 12 12.01 20.73 6.96"></polyline>
								<line x1="12" y1="22.08" x2="12" y2="12"></line>
							</svg>
							<div>
								<h1>📊 Panel Fulfillment <span class="version-badge">v6.0 </span></h1>
								<p>✨ Filtro selectivo de zonas • Sparklines normalizadas • Históricos consistentes • Alarmas FMS • Equipos • Filtro anti-CATEC</p>
							</div>
						</div>
						<div class="header-actions">
							<div class="utilities-dropdown">
								<button class="btn-utilities" onclick="toggleUtilities()">
									🔧 Utilidades
								</button>
								<div class="utilities-menu" id="utilitiesMenu">
									<div class="utilities-menu-group">
										<div class="utilities-menu-header">Recursos rápidos</div>
										<div class="utilities-menu-item" onclick="openPlanillasNewTab()">
											📋 Planillas
										</div>
										<div class="utilities-menu-item" onclick="openUsefulLinks()">
											🔗 Links útiles
										</div>
									</div>
								</div>
							</div>
							<button class="btn-help" onclick="toast('💡 v5.0 Completo: Filtro selectivo de zonas estilo Excel + Links útiles renovados, sparklines normalizadas y filtro anti-CATEC (últimas 24h)')">
								💡 Info
							</button>
						</div>
					</div>
				</header>

				<section class="load-panel">
					<div class="load-section">
						<div class="section-header">📁 CARGAR REPORTES</div>
						
						<div class="load-buttons">
							<div class="load-item">
								<label for="fileConsolidado1" class="btn btn-primary">📄 Consolidado 1</label>
								<input type="file" id="fileConsolidado1" accept=".xlsx,.xls,.csv" hidden>
								<span class="file-status" id="status1">Sin cargar</span>
							</div>
							
							<div class="load-item">
								<label for="fileConsolidado2" class="btn btn-primary">📄 Consolidado 2</label>
								<input type="file" id="fileConsolidado2" accept=".xlsx,.xls,.csv" hidden>
								<span class="file-status" id="status2">Sin cargar</span>
							</div>
							
							<div class="load-item">
								<label for="fileNodos" class="btn btn-primary">🌐 Nodos UP/DOWN</label>
								<input type="file" id="fileNodos" accept=".xlsx,.xls,.csv" hidden>
								<span class="file-status" id="status3">Sin cargar</span>
							</div>
							
							<div class="load-item">
								<label for="fileFMS" class="btn btn-primary">🚨 Alarmas FMS</label>
								<input type="file" id="fileFMS" accept=".xlsx,.xls,.csv" hidden>
								<span class="file-status" id="status4">Sin cargar</span>
							</div>
							
							<div class="merge-status" id="mergeStatus" style="display: none;">
								✓ <span id="mergeStatusText">Procesando...</span>
							</div>
						</div>
					</div>
				</section>

				<nav class="filters-panel">
					<div class="filter-group">
						<div class="filter-label">🔥 Ordenar Zonas</div>
						<select id="ordenarPorIngreso" class="input-select">
							<option value="desc">Mayor a Menor Ingreso</option>
							<option value="asc">Menor a Mayor Ingreso</option>
							<option value="">Sin ordenar (Score)</option>
						</select>
					</div>

					<div class="filter-group">
						<div class="filter-label">Filtros Base</div>
						<label class="checkbox-label">
							<input type="checkbox" id="filterCATEC">
							<span>Solo CATEC</span>
						</label>
						<label class="toggle-control">
							<input type="checkbox" id="filterExcludeCATEC">
							<span class="toggle-track">
								<span class="toggle-thumb"></span>
							</span>
							<span class="toggle-text">Excluir CATEC</span>
						</label>
						<label class="checkbox-label">
							<input type="checkbox" id="showAllStates">
							<span>Ver todos los estados</span>
						</label>
					</div>

					<div class="filter-group">
						<div class="filter-label">Tecnología</div>
						<label class="checkbox-label">
							<input type="checkbox" id="filterFTTH">
							<span>Solo FTTH (9xx)</span>
						</label>
						<label class="checkbox-label">
							<input type="checkbox" id="filterExcludeFTTH">
							<span>Excluir FTTH</span>
						</label>
					</div>
					
					<div class="filter-group">
						<div class="filter-label">Estado Nodo</div>
						<select id="filterNodoEstado" class="input-select">
							<option value="">Todos</option>
							<option value="up">Solo UP</option>
							<option value="down">Solo DOWN</option>
							<option value="critical">Críticos (>50 DOWN)</option>
						</select>
					</div>
					
					<div class="filter-group">
						<div class="filter-label">CMTS</div>
						<select id="filterCMTS" class="input-select">
							<option value="">Todos</option>
						</select>
					</div>
					
					<div class="filter-group">
						<div class="filter-label">⭐ Ventana Temporal</div>
						<select id="daysWindow" class="input-select">
							<option value="3" selected>3 días</option>
							<option value="7">7 días</option>
							<option value="14">14 días</option>
							<option value="30">30 días</option>
						</select>
					</div>
					
					<div class="filter-group">
						<div class="filter-label">Territorio</div>
						<select id="filterTerritorio" class="input-select">
							<option value="">Todos</option>
						</select>
					</div>
					
					<div class="filter-group">
						<div class="filter-label">Sistema</div>
						<select id="filterSistema" class="input-select">
							<option value="">Todos</option>
							<option value="OPEN">OPEN (8...)</option>
							<option value="FAN">FAN (3...)</option>
						</select>
					</div>
					
					<div class="filter-group">
						<div class="filter-label">Alarmas</div>
						<select id="filterAlarma" class="input-select">
							<option value="">Todas</option>
							<option value="con-alarma">Con alarma</option>
							<option value="sin-alarma">Sin alarma</option>
						</select>
					</div>
					
					<div class="filter-group filter-search">
						<div class="filter-label">🔍 Búsqueda Múltiple</div>
						<input type="text" id="quickSearch" class="input-field" placeholder="Ej: VLU901;CON376D;RSC397B">
						<div class="search-help">Usa ; para múltiples zonas • Sin acentos • Selecciona zonas específicas con el filtro múltiple</div>
					</div>

					<div class="filter-group filter-zonas">
						<div class="filter-label">📍 Filtro selectivo de zonas</div>
						<div class="zone-multiselect" id="zoneFilter">
							<button type="button" class="zone-multiselect__trigger" id="zoneFilterTrigger" onclick="toggleZoneFilter(event)">
								<span id="zoneFilterSummary">Sin datos (carga reportes)</span>
								<span class="zone-multiselect__chevron">▾</span>
							</button>
							<div class="zone-multiselect__dropdown" id="zoneFilterDropdown">
								<div class="zone-multiselect__search">
									<input type="text" id="zoneFilterSearch" placeholder="Buscar zona..." autocomplete="off">
									<div class="zone-multiselect__hint">Escribe las primeras letras para ver coincidencias y marca las zonas a mostrar</div>
								</div>
								<div class="zone-multiselect__options" id="zoneFilterOptions">
									<div class="zone-multiselect__empty">Cargá los reportes para habilitar el filtro por zonas</div>
								</div>
								<div class="zone-multiselect__actions">
									<button type="button" class="zone-action" onclick="selectAllZones(event)">Seleccionar todo</button>
									<button type="button" class="zone-action" onclick="clearZoneSelection(event)">Limpiar</button>
								</div>
							</div>
						</div>
					</div>

					<div class="filter-actions">
						<button class="btn btn-secondary" onclick="clearFilters()">
							🔄 Limpiar
						</button>
					</div>
				</nav>

				<main class="main-content">
					
					<section class="stats-grid" id="statsGrid">
						<div class="loading-message">
							<div class="spinner"></div>
							<p>Carga los archivos para comenzar</p>
						</div>
					</section>

					<section class="content-tabs">
						<div class="tabs-nav">
							<button class="tab-btn active" data-tab="zonas" onclick="switchTab('zonas')">🗺️ Zonas</button>
							<button class="tab-btn" data-tab="cmts" onclick="switchTab('cmts')">🧭 CMTS</button>
							<button class="tab-btn" data-tab="edificios" onclick="switchTab('edificios')">🏢 Edificios</button>
							<button class="tab-btn" data-tab="equipos" onclick="switchTab('equipos')">🔧 Equipos</button>
							<button class="tab-btn" data-tab="fms" onclick="switchTab('fms')">🚨 FMS</button>
						</div>

						<div class="tabs-content">
							<div id="zonasPanel" class="tab-panel active">
								<div class="loading-message"><p>Los datos de zonas aparecerán aquí</p></div>
							</div>
							<div id="cmtsPanel" class="tab-panel">
								<div class="loading-message"><p>Los datos agrupados por CMTS aparecerán aquí</p></div>
							</div>
							<div id="edificiosPanel" class="tab-panel">
								<div class="loading-message"><p>Los datos de edificios aparecerán aquí</p></div>
							</div>
							<div id="equiposPanel" class="tab-panel">
								<div class="loading-message"><p>Los equipos por zona aparecerán aquí</p></div>
							</div>
							<div id="fmsPanel" class="tab-panel">
								<div class="loading-message"><p>Los datos FMS aparecerán aquí</p></div>
							</div>
						</div>
					</section>

					<div class="export-section">
						<button class="btn btn-success" id="btnExportExcel" disabled onclick="exportExcelVista()">
							📥 Exportar vista (Excel)
						</button>
						<button class="btn btn-secondary" id="btnExportExcelZonas" disabled onclick="exportExcelZonasCrudo()">
							📥 Exportar zonas (crudo)
						</button>
					</div>
				</main>

				<footer class="app-footer">
					<p>Panel Fulfillment v6.0 © 2025</p>
					<p class="credits">Creado por Gabriel Palazzini </p>
				</footer>
			</div>

			<div id="modalBackdrop" class="modal-backdrop" onclick="closeModal()">
				<div class="modal" onclick="event.stopPropagation()">
					<div class="modal-header">
						<div class="modal-title" id="modalTitle">Detalle de Zona</div>
						<button class="modal-close" onclick="closeModal()">×</button>
					</div>
					
					<div class="modal-filters" id="modalFilters">
						<div class="filter-group" style="min-width: 200px;">
							<div class="filter-label">Filtrar por Horario</div>
							<select id="filterHorario" class="input-select" onchange="applyModalFilters()">
								<option value="">Todos los horarios</option>
								<option value="0-6">00:00 - 06:00</option>
								<option value="6-12">06:00 - 12:00</option>
								<option value="12-18">12:00 - 18:00</option>
								<option value="18-24">18:00 - 24:00</option>
							</select>
						</div>
						
						<div class="filter-group" style="min-width: 200px;">
							<div class="filter-label">Filtrar por Día</div>
							<select id="filterDia" class="input-select" onchange="applyModalFilters()">
								<option value="">Todos los días</option>
							</select>
						</div>
						
						<div class="filter-group">
							<label class="checkbox-label">
								<input type="checkbox" id="selectAllOrders" onchange="toggleSelectAll()">
								<span>Seleccionar Todas</span>
							</label>
						</div>
					</div>
					
					<div class="modal-body" id="modalBody">
						<div class="loading-message">
							<div class="spinner"></div>
							<p>Cargando detalles...</p>
						</div>
					</div>
					
					<div class="modal-footer" id="modalFooter">
						<div class="selection-info" id="selectionInfo">
							0 órdenes seleccionadas
						</div>
						<div style="display: flex; gap: 8px;">
							<button class="btn btn-warning" id="btnExportBEFAN" disabled onclick="exportBEFAN()">
								📤 Exportar BEFAN (TXT)
							</button>
							<button class="btn btn-secondary" onclick="closeModal()">
								Cerrar
							</button>
						</div>
					</div>
				</div>
			</div>

			<div id="alarmaBackdrop" class="modal-backdrop" onclick="closeAlarmaModal()">
				<div class="modal" onclick="event.stopPropagation()" style="max-width: 1000px;">
					<div class="modal-header">
						<div class="modal-title">🚨 Información de Alarmas</div>
						<button class="modal-close" onclick="closeAlarmaModal()">×</button>
					</div>
					
					<div class="modal-body" id="alarmaModalBody">
						<div class="loading-message">
							<div class="spinner"></div>
							<p>Cargando datos de alarmas...</p>
						</div>
					</div>
					
					<div class="modal-footer">
						<button class="btn btn-primary" onclick="closeAlarmaModal()">
							Cerrar
						</button>
					</div>
				</div>
			</div>

			<div id="edificioBackdrop" class="modal-backdrop" onclick="closeEdificioModal()">
				<div class="modal" onclick="event.stopPropagation()" style="max-width: 1200px;">
					<div class="modal-header">
						<div class="modal-title" id="edificioModalTitle">Detalle de Edificio</div>
						<button class="modal-close" onclick="closeEdificioModal()">×</button>
					</div>
					
					<div class="modal-body" id="edificioModalBody">
						<div class="loading-message">
							<div class="spinner"></div>
							<p>Cargando órdenes del edificio...</p>
						</div>
					</div>
					
					<div class="modal-footer">
						<button class="btn btn-secondary" onclick="exportEdificioDetalle()">
							📥 Exportar órdenes
						</button>
						<button class="btn btn-primary" onclick="closeEdificioModal()">
							Cerrar
						</button>
					</div>
				</div>
			</div>
			</div>

		<div class="card-footer">
			<label>Performance Eficiencia y Mejora</label> <a href="mailto:dataofficepem@teco.com.ar" target="_top">dataofficepem@teco.com.ar</a>
		</div>





		<iframe class="d-none" id="miframe" name="miframe">

	</div>

	<?php 
	echo "<script src='assets/js/data-processor.js?$aleatorio'></script>";
	echo "<script src='assets/js/ui-renderer.js?$aleatorio'></script>";
	echo "<script src='assets/js/app.js?$aleatorio'></script>"; 
	?>

</body>

</HTML>
}
