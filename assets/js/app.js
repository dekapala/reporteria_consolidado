console.log('🚀 Panel v5.0 COMPLETO - Territorios Críticos + Filtros Equipos + Stats Clickeables');

const CONFIG = {
  codigo_befan: 'FR461',
  defaultDays: 3,
  estadosPermitidos: [
    'NUEVA',
    'EN PROGRESO',
    'PENDIENTE DE ACCION',
    'PROGRAMADA',
    'PENDIENTE DE CONTACTO',
    'EN ESPERA DE EJECUCION'
  ],
  estadosOcultosPorDefecto: ['CANCELADA', 'CERRADA']
};

const FMS_TIPOS = {
  'ED': 'Edificio',
  'CDO': 'CDO',
  'PNO': 'Puerto Nodo Óptico',
  'FA': 'Fuente Alimentación',
  'NO': 'Nodo Óptico',
  'LE': 'Line Extender',
  'MB': 'Mini Bridge',
  'SR_HUB': 'Sitio Red Hub',
  'CMTS': 'CMTS',
  'SITIO_MOVIL': 'Sitio Móvil',
  'TR': 'Troncal',
  'EDF': 'Edificio FTTH',
  'CE': 'Caja Empalme',
  'NAP': 'NAP'
};

const DIAGNOSTICO_TECNICO_KEYS = [
  'Diagnostico Tecnico',
  'Diagnóstico Técnico',
  'Diagnostico tecnico',
  'Diagnóstico tecnico',
  'Diagnostico técnico',
  'Diagnóstico Técnico '
];

const TerritorioUtils = {
  normalizar(territorio) {
    if (!territorio) return '';
    return territorio
      .replace(/\s*flex\s*\d+/gi, ' Flex')
      .replace(/\s{2,}/g, ' ')
      .trim();
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

// ═══════════════════════════════════════════════════════════════════
// TEXTUTILS - Utilidades para parsear texto y JSON de dispositivos
// ═══════════════════════════════════════════════════════════════════
const TextUtils = {
  /**
   * Parsea el JSON de dispositivos desde la columna "Informacion Dispositivos"
   * Formato esperado: {"openResp":[{"model":"BCEG670","macAddress":"1494488F450A","type":"Cablemodem"...}]}
   * @param {string|object} raw - Texto JSON o objeto ya parseado
   * @returns {Array} Array de dispositivos normalizados
   */
  parseDispositivosJSON(raw) {
    if (!raw) return [];
    
    try {
      let parsed;
      
      // Si ya es un objeto, usarlo directamente
      if (typeof raw === 'object' && raw !== null) {
        parsed = raw;
      } else {
        // Intentar parsear como JSON
        const str = String(raw).trim();
        if (!str) return [];
        
        // Limpiar el string si tiene caracteres extraños
        const cleanStr = str
          .replace(/^\s*[;,]\s*/, '') // Eliminar ; o , al inicio
          .replace(/\s*[;,]\s*$/, ''); // Eliminar ; o , al final
        
        if (!cleanStr) return [];
        
        parsed = JSON.parse(cleanStr);
      }
      
      // Extraer dispositivos del objeto parseado
      let dispositivos = [];
      
      // Formato 1: {"openResp": [{...}]}
      if (parsed.openResp && Array.isArray(parsed.openResp)) {
        dispositivos = parsed.openResp;
      }
      // Formato 2: Array directo [{...}]
      else if (Array.isArray(parsed)) {
        dispositivos = parsed;
      }
      // Formato 3: Objeto único {...}
      else if (typeof parsed === 'object') {
        dispositivos = [parsed];
      }
      
      // Normalizar cada dispositivo
      return dispositivos.map(dev => {
        if (!dev || typeof dev !== 'object') return null;
        
        return {
          // Datos principales
          model: dev.model || dev.modelo || '',
          macAddress: dev.macAddress || dev.mac || dev.mac_address || dev.MAC || '',
          type: dev.type || dev.tipo || '',
          
          // Datos adicionales
          category: dev.category || dev.brand || dev.marca || '',
          brand: dev.brand || dev.category || dev.marca || '',
          technology: dev.technology || dev.description || dev.detalle || '',
          description: dev.description || dev.detalle || '',
          serialNumber: dev.serialNumber || dev.serial || dev.serie || '',
          lifecycleState: dev.lifecycleState || dev.estado || '',
          publicIdentifier: dev.publicIdentifier || dev.id || '',
          esbType: dev.esbType || ''
        };
      }).filter(dev => {
        // Filtrar dispositivos que tengan al menos uno de los campos principales
        return dev && (dev.model || dev.macAddress || dev.type);
      });
      
    } catch (error) {
      console.warn('Error al parsear Informacion Dispositivos:', error.message, 'Raw:', raw);
      return [];
    }
  },

  /**
   * Extrae información resumida de dispositivos para mostrar en tablas
   * @param {Array} dispositivos - Array de dispositivos normalizados
   * @returns {string} String formateado con info de dispositivos
   */
  formatDispositivosResumen(dispositivos) {
    if (!Array.isArray(dispositivos) || !dispositivos.length) {
      return 'Sin información';
    }
    
    const resumen = dispositivos.map(dev => {
      const parts = [];
      if (dev.model) parts.push(dev.model);
      if (dev.type) parts.push(`(${dev.type})`);
      if (dev.macAddress) parts.push(dev.macAddress.substring(0, 17));
      return parts.join(' ');
    });
    
    return resumen.join(' | ');
  }
};

function stripAccents(s=''){return s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');}

function extractDiagnosticoTecnico(order){
  if (!order) return '';
  const metaValue = order.__meta?.diagnosticoTecnico;
  if (metaValue) {
    const str = String(metaValue).trim();
    if (str) return str;
  }
  for (const key of DIAGNOSTICO_TECNICO_KEYS){
    if (!key) continue;
    const value = order[key];
    if (value !== null && value !== undefined){
      const str = String(value).trim();
      if (str) return str;
    }
  }
  return '';
}

function findDispositivosColumn(rowObj){
  const keys = Object.keys(rowObj||{});
  for (const k of keys){
    const nk = stripAccents(k.toLowerCase());
    if (nk.includes('informacion') && nk.includes('dispositivo')) return k;
  }
  return rowObj['Informacion Dispositivos'] ? 'Informacion Dispositivos'
       : rowObj['Información Dispositivos'] ? 'Información Dispositivos'
       : null;
}

async function readFileAsUint8Array(file){
  if (!file) return new Uint8Array();
  if (file.arrayBuffer) {
    const buf = await file.arrayBuffer();
    return new Uint8Array(buf);
  }
  return await new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(new Uint8Array(fr.result));
    fr.onerror = reject;
    fr.readAsArrayBuffer(file);
  });
}

function on(id, evt, handler){
  const el = document.getElementById(id);
  if (!el){
    console.warn(`⚠️ No se encontró el elemento con id "${id}" para adjuntar ${evt}`);
    return null;
  }
  el.addEventListener(evt, handler);
  return el;
}

function normalizeMac(mac=''){
  const clean = String(mac).trim().toUpperCase().replace(/[^0-9A-F]/g,'');
  if (clean.length !== 12) return '';
  return clean.match(/.{1,2}/g).join(':');
}

function pickPreferredDevice(devs=[]){
  if (!Array.isArray(devs) || !devs.length) return null;

  const mapValue = (value) => {
    if (value === null || value === undefined) return '';
    const str = String(value).trim();
    return str;
  };

  const normalizedDevices = devs.map(dev => {
    if (!dev || typeof dev !== 'object') return null;
    const brand = mapValue(dev.brand || dev.category || dev.marca);
    const technology = mapValue(dev.technology || dev.description || dev.detalle);
    const model = mapValue(dev.model || dev.modelo);
    const macRaw = dev.mac || dev.macAddress || dev.mac_address || dev.MAC;
    const mac = normalizeMac(macRaw) || mapValue(macRaw);
    const type = mapValue(dev.type || dev.tipo);

    if (!(brand || technology || model || mac || type)) return null;

    return {
      brand,
      technology,
      model,
      mac,
      type
    };
  }).filter(Boolean);

  if (!normalizedDevices.length) return null;

  const withMac = normalizedDevices.find(d => d.mac);
  if (withMac) return withMac;

  const withModel = normalizedDevices.find(d => d.model);
  if (withModel) return withModel;

  return normalizedDevices[0];
}

function ensureOrderDeviceMeta(order){
  if (!order || typeof order !== 'object') return null;

  if (!order.__meta) order.__meta = {};
  if (order.__meta.device) return order.__meta.device;

  let dispositivos = [];

  if (Array.isArray(order.__meta.dispositivos) && order.__meta.dispositivos.length) {
    dispositivos = order.__meta.dispositivos;
  } else {
    const colInfo = findDispositivosColumn(order);
    if (colInfo && order[colInfo]) {
      // AQUÍ SE USA LA NUEVA FUNCIÓN TextUtils.parseDispositivosJSON
      dispositivos = TextUtils.parseDispositivosJSON(order[colInfo]);
      console.log('📱 Dispositivos extraídos:', dispositivos.length, 'de columna:', colInfo);
    }
  }

  const normalizedDevices = Array.isArray(dispositivos)
    ? dispositivos.map(dev => {
        if (!dev || typeof dev !== 'object') return null;
        const mapValue = (val) => val === null || val === undefined ? '' : String(val).trim();
        
        // Normalizar MAC address
        const macAddress = normalizeMac(dev.macAddress || dev.mac || dev.mac_address || dev.MAC) 
                          || mapValue(dev.macAddress || dev.mac);
        
        const serialNumber = mapValue(dev.serialNumber || dev.serial || dev.serie);
        const model = mapValue(dev.model || dev.modelo);
        const category = mapValue(dev.category || dev.brand || dev.marca);
        const technology = mapValue(dev.technology || dev.description || dev.detalle);
        const type = mapValue(dev.type || dev.tipo);
        const brand = mapValue(dev.brand || dev.category || dev.marca || category);

        if (!(macAddress || serialNumber || model || category || technology || type || brand)) {
          return null;
        }

        return {
          macAddress,
          serialNumber,
          model,
          category,
          technology,
          type,
          brand,
          mac: macAddress
        };
      }).filter(Boolean)
    : [];

  order.__meta.dispositivos = normalizedDevices;
  dispositivos = normalizedDevices;

  const device = pickPreferredDevice(dispositivos) || null;
  if (device && Array.isArray(dispositivos) && dispositivos.length) {
    const [first] = dispositivos;
    if (device !== first) {
      order.__meta.dispositivos = [device, ...dispositivos.filter(d => d !== device)];
    }
  }
  if (device) {
    order.__meta.device = device;
    console.log('✅ Dispositivo seleccionado:', device.model || device.type || 'sin modelo', 
                'MAC:', device.mac || 'sin MAC');
    return device;
  }

  order.__meta.device = null;
  return null;
}

const DateUtils = {
  parse(str) {
    if (!str) return null;
    const match = String(str).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (match) return new Date(match[3], match[2] - 1, match[1]);
    return new Date(str);
  },
  format(date) {
    if (!date) return '';
    const d = String(date.getDate()).padStart(2,'0');
    const m = String(date.getMonth()+1).padStart(2,'0');
    return `${d}/${m}/${date.getFullYear()}`;
  },
  formatWithTime(date) {
    if (!date) return '';
    const d = String(date.getDate()).padStart(2,'0');
    const m = String(date.getMonth()+1).padStart(2,'0');
    const h = String(date.getHours()).padStart(2,'0');
    const min = String(date.getMinutes()).padStart(2,'0');
    return `${d}/${m}/${date.getFullYear()} ${h}:${min}`;
  },
  toDayKey(date) {
    const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    return d.getTime();
  },
  getHour(date) {
    return date.getHours();
  },
  daysBetween(date1, date2) {
    const oneDay = 24 * 60 * 60 * 1000;
    const diff = date2 - date1;
    return Math.round(diff / oneDay);
  }
};

let ordersData = null;
let nodosData = null;
let alarmasData = null;
let window_currentAnalyzedZones = [];
let window_lastFilteredOrders = [];
let window_currentCMTSData = [];
let window_edificiosData = [];
let window_equiposPorZona = new Map();

// Estado de filtros
let activeFilters = {
  tipoRed: [],
  estadosOT: [],
  cmts: [],
  equipos: {
    modelos: [],
    tipos: [],
    marcas: []
  },
  territorios: [],
  dias: CONFIG.defaultDays
};

const selectedOrders = new Set();
let currentZone = null;

const ZONE_EXPORT_HEADERS = [
  'Número de cita',
  'Caso Externo',
  'Fecha de creación',
  'Hora de creación',
  'Fecha Finalizacion',
  'Hora Finalizacion',
  'Días desde creación',
  'Estado',
  'Motivo',
  'Territorio',
  'Zona',
  'Tipo de Zona',
  'CMTS',
  'Direccion',
  'Nodo',
  'Estado Nodo',
  'Tiene Alarma',
  'Alarmas Activas',
  'Modelo Equipo',
  'MAC Equipo',
  'Tipo Equipo',
  'Diagnostico Tecnico'
];

function toast(msg){
  const t = document.createElement('div');
  t.textContent = msg;
  t.style.cssText = `
    position:fixed; bottom:20px; right:20px; 
    background:#333; color:#fff; padding:12px 20px; 
    border-radius:6px; z-index:99999; 
    box-shadow:0 4px 12px rgba(0,0,0,0.3);
    font-size:14px; font-family:system-ui;
  `;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 3500);
}

async function onFileUpload(evt){
  const file = evt.target.files[0];
  if (!file) return;
  
  const fileName = file.name.toLowerCase();
  try {
    const data = await readFileAsUint8Array(file);
    const wb = XLSX.read(data, {type:'array', cellDates:true, cellNF:false, cellText:false});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, {raw:false, defval:''});
    
    if (fileName.includes('consolidado')) {
      ordersData = rows;
      toast(`✓ ${rows.length} órdenes cargadas`);
    } else if (fileName.includes('nodos')) {
      nodosData = rows;
      toast(`✓ ${rows.length} nodos cargados`);
    } else if (fileName.includes('fms') || fileName.includes('alarma')) {
      alarmasData = rows;
      toast(`✓ ${rows.length} alarmas cargadas`);
    } else {
      toast('⚠️ Archivo no reconocido');
      return;
    }
    
    updateButtonStates();
  } catch (error) {
    console.error('Error cargando archivo:', error);
    toast(`❌ Error: ${error.message}`);
  }
}

function updateButtonStates(){
  const hasOrders = ordersData && ordersData.length > 0;
  const hasNodos = nodosData && nodosData.length > 0;
  const hasAlarmas = alarmasData && alarmasData.length > 0;
  
  const btnAnalizar = document.getElementById('btnAnalizar');
  if (btnAnalizar) btnAnalizar.disabled = !hasOrders;
}

function renderInitialUI(){
  const container = document.getElementById('app');
  if (!container) return;
  
  container.innerHTML = `
    <div style="max-width:1400px; margin:0 auto; padding:20px; font-family:system-ui,-apple-system,sans-serif;">
      <div style="background:linear-gradient(135deg,#667eea 0%,#764ba2 100%); padding:30px; border-radius:12px; margin-bottom:30px; box-shadow:0 8px 32px rgba(0,0,0,0.1);">
        <h1 style="color:#fff; margin:0 0 10px 0; font-size:32px; font-weight:700;">Panel Fulfillment v5.0</h1>
        <p style="color:rgba(255,255,255,0.9); margin:0; font-size:16px;">Territorios Críticos • Filtros Equipos • Exportación BEFAN</p>
      </div>
      
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(300px, 1fr)); gap:20px; margin-bottom:30px;">
        <div style="background:#fff; padding:24px; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,0.08);">
          <h3 style="margin:0 0 16px 0; font-size:18px; color:#333; font-weight:600;">📊 Consolidado</h3>
          <input type="file" id="fileConsolidado" accept=".xlsx,.xls" style="width:100%; padding:10px; border:2px dashed #ddd; border-radius:8px; cursor:pointer; font-size:14px;">
        </div>
        
        <div style="background:#fff; padding:24px; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,0.08);">
          <h3 style="margin:0 0 16px 0; font-size:18px; color:#333; font-weight:600;">🌐 Nodos (opcional)</h3>
          <input type="file" id="fileNodos" accept=".xlsx,.xls" style="width:100%; padding:10px; border:2px dashed #ddd; border-radius:8px; cursor:pointer; font-size:14px;">
        </div>
        
        <div style="background:#fff; padding:24px; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,0.08);">
          <h3 style="margin:0 0 16px 0; font-size:18px; color:#333; font-weight:600;">🚨 FMS/Alarmas (opcional)</h3>
          <input type="file" id="fileFMS" accept=".xlsx,.xls" style="width:100%; padding:10px; border:2px dashed #ddd; border-radius:8px; cursor:pointer; font-size:14px;">
        </div>
      </div>
      
      <div style="text-align:center;">
        <button id="btnAnalizar" disabled style="background:#667eea; color:#fff; border:none; padding:16px 40px; border-radius:8px; font-size:16px; font-weight:600; cursor:pointer; box-shadow:0 4px 12px rgba(102,126,234,0.3); transition:all 0.3s;">
          🚀 Analizar Datos
        </button>
      </div>
      
      <div id="results" style="margin-top:30px;"></div>
    </div>
    
    <!-- Modal Detalle -->
    <div id="detailModal" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:9999; overflow:auto;">
      <div style="background:#fff; max-width:95%; margin:20px auto; border-radius:12px; box-shadow:0 8px 32px rgba(0,0,0,0.2); max-height:90vh; display:flex; flex-direction:column;">
        <div style="padding:20px; border-bottom:1px solid #eee; display:flex; justify-content:space-between; align-items:center;">
          <h2 id="modalTitle" style="margin:0; font-size:24px; color:#333;"></h2>
          <button onclick="closeDetailModal()" style="background:#f44336; color:#fff; border:none; padding:8px 16px; border-radius:6px; cursor:pointer; font-weight:600;">Cerrar</button>
        </div>
        <div id="modalBody" style="padding:20px; overflow-y:auto; flex:1;"></div>
        <div style="padding:20px; border-top:1px solid #eee; display:flex; justify-content:space-between; align-items:center; background:#f9f9f9;">
          <span id="selectionInfo" style="font-weight:600; color:#666;">0 órdenes seleccionadas</span>
          <div style="display:flex; gap:10px;">
            <button id="btnExportBEFAN" onclick="exportBEFAN()" disabled style="background:#4CAF50; color:#fff; border:none; padding:10px 20px; border-radius:6px; cursor:pointer; font-weight:600;">📤 Exportar BEFAN</button>
            <button onclick="exportModalDetalleExcel()" style="background:#2196F3; color:#fff; border:none; padding:10px 20px; border-radius:6px; cursor:pointer; font-weight:600;">📊 Exportar Excel</button>
          </div>
        </div>
      </div>
    </div>
  `;
  
  on('fileConsolidado', 'change', onFileUpload);
  on('fileNodos', 'change', onFileUpload);
  on('fileFMS', 'change', onFileUpload);
  on('btnAnalizar', 'click', analyzeData);
}

function analyzeData(){
  if (!ordersData || !ordersData.length) {
    toast('⚠️ Primero carga el archivo Consolidado');
    return;
  }
  
  console.log('🔍 Iniciando análisis...');
  console.log('📊 Órdenes:', ordersData.length);
  
  // Procesar cada orden para extraer información de dispositivos
  ordersData.forEach((order, idx) => {
    if (idx < 5) console.log(`📝 Procesando orden ${idx + 1}...`);
    ensureOrderDeviceMeta(order);
  });
  
  toast('✅ Análisis completado');
  
  // Aquí continuaría el resto de tu lógica de análisis...
  // Por ahora solo muestro que se procesaron los dispositivos
  
  const resultsDiv = document.getElementById('results');
  if (resultsDiv) {
    resultsDiv.innerHTML = `
      <div style="background:#fff; padding:24px; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,0.08);">
        <h3 style="margin:0 0 16px 0; color:#333;">✅ Análisis Completado</h3>
        <p>Se procesaron ${ordersData.length} órdenes</p>
        <p>Revisa la consola para ver los dispositivos extraídos</p>
      </div>
    `;
  }
}

function closeDetailModal(){
  document.getElementById('detailModal').style.display = 'none';
}

function appendSheet(wb, sheet, name, usedNames) {
  let finalName = name;
  let counter = 1;
  while (usedNames.has(finalName)) {
    finalName = `${name}_${counter}`;
    counter++;
  }
  usedNames.add(finalName);
  XLSX.utils.book_append_sheet(wb, sheet, finalName);
}

function createWorksheetFromRows(rows, headers) {
  if (!rows || !rows.length) {
    return XLSX.utils.aoa_to_sheet([headers]);
  }
  const ws = XLSX.utils.json_to_sheet(rows, {header: headers});
  return ws;
}

function buildZoneExportRows(zone) {
  if (!zone || !zone.ordenes) return [];
  return zone.ordenes.map(o => {
    const device = ensureOrderDeviceMeta(o);
    return {
      'Número de cita': o['Número de cita'] || '',
      'Caso Externo': o['Caso Externo'] || o['External Case Id'] || '',
      'Fecha de creación': o['Fecha de creación'] || '',
      'Hora de creación': o['Hora de creación'] || '',
      'Fecha Finalizacion': o['Fecha Finalizacion'] || '',
      'Hora Finalizacion': o['Hora Finalizacion'] || '',
      'Días desde creación': '', // Calcular si es necesario
      'Estado': o['Estado'] || '',
      'Motivo': o['Motivo'] || '',
      'Territorio': o['Territorio'] || '',
      'Zona': zone.zona || '',
      'Tipo de Zona': zone.tipo || '',
      'CMTS': zone.cmts || '',
      'Direccion': o['Dirección de instalación'] || o['Direccion'] || '',
      'Nodo': zone.nodo || '',
      'Estado Nodo': zone.nodoEstado || '',
      'Tiene Alarma': zone.tieneAlarma ? 'SÍ' : 'NO',
      'Alarmas Activas': zone.alarmasActivas || 0,
      'Modelo Equipo': device?.model || '',
      'MAC Equipo': device?.mac || device?.macAddress || '',
      'Tipo Equipo': device?.type || '',
      'Diagnostico Tecnico': extractDiagnosticoTecnico(o)
    };
  });
}

function exportBEFAN() {
  if (selectedOrders.size === 0) {
    toast('No hay órdenes seleccionadas');
    return;
  }
  
  const ordenes = currentZone.ordenes.filter(o => {
    const cita = o['Número de cita'];
    return selectedOrders.has(cita);
  });
  
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
    a.download = `BEFAN_${currentZone.zona}_${new Date().toISOString().slice(0, 10)}.txt`;
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

function exportModalDetalleExcel(){
  const title = document.getElementById('modalTitle').textContent || '';
  const wb = XLSX.utils.book_new();

  if (title.startsWith('Detalle: ') && window.currentZone){
    const ordenes = window.currentZone.ordenes || [];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ordenes), `Zona_${window.currentZone.zona || 'NA'}`);
  } else {
    return toast('No hay detalle para exportar');
  }

  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Detalle_${fecha}.xlsx`);
  toast('✓ Detalle exportado');
}

function toggleSelectAll() {
  const checked = document.getElementById('selectAllCheckbox')?.checked || false;
  
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

console.log('✅ Panel v5.0 inicializado con extracción mejorada de dispositivos');
console.log('📱 TextUtils.parseDispositivosJSON disponible para extraer:');
console.log('   • model: Modelo del equipo');
console.log('   • macAddress: Dirección MAC');
console.log('   • type: Tipo de equipo');
console.log('   • category/brand: Marca/Categoría');
console.log('   • technology/description: Tecnología/Descripción');

// Inicializar la UI
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', renderInitialUI);
} else {
  renderInitialUI();
}
