// ---- SheetJS compat para ES modules ----
const XLSX = window.XLSX ?? null;
if (!XLSX) {
  console.error('❌ XLSX no está disponible. Verifica el <script> del CDN antes de app.js');
  alert('No se encontró la librería XLSX. Revisá el orden de los <script>.');
}

console.log('🚀 Panel v5.0 COMPLETO - Territorios Críticos + Filtros Equipos + Stats Clickeables');

function debounce(func, wait) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

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
  'CDO': 'Centro de Distribución Óptica',
  'CO': 'Central Office',
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

function formatFMSTypeLabel(tipo) {
  if (!tipo) return 'Sin tipo';
  const desc = FMS_TIPOS[tipo];
  return desc ? `${tipo} (${desc})` : tipo;
}

function formatAlarmaDamage(alarma = {}) {
  const damage = alarma.damage || alarma.damageClassification || alarma.incidentClassification;
  return damage && String(damage).trim() ? damage : 'Sin información';
}

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

function stripAccents(s=''){return s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');}

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
    return Math.round(Math.abs((date1 - date2) / oneDay));
  }
};

const NumberUtils = {
  format(num, dec=0) {
    if (typeof num !== 'number') return '0';
    return num.toFixed(dec).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  }
};

const TextUtils = {
  normalize(text) {
    return String(text||'')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g,'')
      .trim();
  },

    
  matches(text, query) {
    return this.normalize(text).includes(this.normalize(query));
  },
  matchesMultiple(text, queries) {
    const normalized = this.normalize(text);
    return queries.some(q => normalized.includes(this.normalize(q)));
  },
  parseDispositivosJSON(jsonStr) {
    try {
      if (!jsonStr) return [];
      
      const str = String(jsonStr).trim();
      if (!str) return [];
      
      // MÉTODO 1: Intentar parsear JSON completo
      try {
        const parsed = JSON.parse(str);
        const arr = parsed?.openResp || parsed?.open_response || parsed?.data || [];
        if (Array.isArray(arr) && arr.length > 0) {
          return arr.map(d => ({
            category: d.category || '',
            description: d.description || '',
            model: d.model || '',
            serialNumber: d.serialNumber || d.serial || '',
            macAddress: d.macAddress || d.mac || '',
            type: d.type || ''
          }));
        }
      } catch (e) {
        // JSON truncado o inválido, continuar con método alternativo
        console.log('JSON truncado, usando extracción por regex');
      }
      
      // MÉTODO 2: Extracción por expresiones regulares (para JSON truncado a 255 chars)
      const get = (field) => {
        // Buscar "field":"value"
        const re = new RegExp(`"${field}"\\s*:\\s*"([^"]*)"`, 'i');
        const match = str.match(re);
        return match ? match[1] : '';
      };
      
      // Verificar si tiene algún campo relevante
      if (/category|model|macAddress|serialNumber|type|description/i.test(str)) {
        const dispositivo = {
          category: get('category'),
          description: get('description'),
          model: get('model'),
          serialNumber: get('serialNumber'),
          macAddress: get('macAddress'),
          type: get('type')
        };
        
        // Solo retornar si al menos tiene model, mac o type
        if (dispositivo.model || dispositivo.macAddress || dispositivo.type) {
          return [dispositivo];
        }
      }
      
    } catch (err) {
      console.warn('Error parseando dispositivos:', err.message);
    }
    return [];
  },
  detectarSistema(numCaso) {
    const numStr = String(numCaso || '').trim();
    if (numStr.startsWith('8')) return 'OPEN';
    if (numStr.startsWith('3')) return 'FAN';
    return '';
  }
};

const OrderTypeClassifier = {
  // Categorías de tipos de orden
  categories: {
    SIN_SENAL: {
      name: 'Sin Señal',
      keywords: [
        'sin señal', 'sin senal', 'no recibe señal', 'sin registracion',
        'no registra', 'sin registro', 'offline', 'fuera de linea',
        'no conecta', 'sin conexion', 'no hay señal'
      ],
      icon: '📡',
      color: '#e74c3c'
    },
    EQUIPO_DAÑADO: {
      name: 'Equipo Dañado',
      keywords: [
        'equipo dañado', 'equipo danado', 'dispositivo dañado', 'dañado',
        'danado', 'roto', 'quemado', 'defectuoso', 'falla hardware',
        'hardware dañado'
      ],
      icon: '🔥',
      color: '#d35400'
    },
    PROBLEMAS_SEÑAL: {
      name: 'Problemas de Señal',
      keywords: [
        'problemas de señal', 'problemas señal', 'señal baja',
        'pixelado', 'imagen pixelada', 'nivel bajo', 'interferencia',
        'ruido', 'atenuacion', 'atenuación', 'mala señal'
      ],
      icon: '📶',
      color: '#f39c12'
    },
    CABLEMODEM: {
      name: 'Cablemódem',
      keywords: [
        'cablemodem', 'cable modem', 'cablemódem', 'cable módem',
        'cm ', 'modem', 'módem'
      ],
      icon: '📟',
      color: '#3498db'
    },
    INSTALACION: {
      name: 'Instalación',
      keywords: [
        'instalacion', 'instalación', 'nuevo cliente', 'alta',
        'activacion', 'activación', 'provision'
      ],
      icon: '🔧',
      color: '#27ae60'
    },
    CAMBIO_EQUIPO: {
      name: 'Cambio de Equipo',
      keywords: [
        'cambio equipo', 'cambio de equipo', 'reemplazo', 'swap',
        'sustitucion', 'sustitución'
      ],
      icon: '🔄',
      color: '#16a085'
    },
    RED: {
      name: 'Problema de Red',
      keywords: [
        'problema red', 'red', 'nodo', 'cmts', 'tap', 'amplificador',
        'optico', 'óptico', 'fibra', 'troncal'
      ],
      icon: '🌐',
      color: '#8e44ad'
    },
    CLIENTE: {
      name: 'Requiere Cliente',
      keywords: [
        'cliente', 'requiere', 'no quiere', 'rechaza', 'no acepta',
        'necesita', 'solicita cliente'
      ],
      icon: '👤',
      color: '#95a5a6'
    }
  },

  /**
   * Clasifica una orden según su diagnóstico técnico
   * @param {string} diagnostico - Diagnóstico técnico de la orden
   * @returns {Object} - {category, name, icon, color}
   */
  classify(diagnostico) {
    if (!diagnostico) {
      return {
        category: 'OTROS',
        name: 'Otros',
        icon: '📋',
        color: '#7f8c8d'
      };
    }

    const diagLower = String(diagnostico).toLowerCase().normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');

    // Buscar en cada categoría
    for (const [key, config] of Object.entries(this.categories)) {
      for (const keyword of config.keywords) {
        if (diagLower.includes(keyword.toLowerCase())) {
          return {
            category: key,
            name: config.name,
            icon: config.icon,
            color: config.color
          };
        }
      }
    }

    // Si no matchea nada, devolver "Otros"
    return {
      category: 'OTROS',
      name: 'Otros',
      icon: '📋',
      color: '#7f8c8d'
    };
  },

  /**
   * Agrega clasificación a una orden
   * @param {Object} order - Orden técnica
   */
  addClassification(order) {
    const diagnostico = order['Diagnostico Tecnico'] || 
                       order['Diagnóstico Técnico'] || 
                       order['Diagnostico tecnico'] || '';
    
    const classification = this.classify(diagnostico);
    
    if (!order.__meta) order.__meta = {};
    order.__meta.tipoOrden = classification;
    
    return classification;
  },

  /**
   * Obtiene estadísticas de tipos de orden de un array
   * @param {Array} orders - Array de órdenes
   * @returns {Object} - Estadísticas por tipo
   */
  getStats(orders) {
    const stats = {};
    
    orders.forEach(order => {
      const classification = this.addClassification(order);
      const cat = classification.category;
      
      if (!stats[cat]) {
        stats[cat] = {
          ...classification,
          count: 0,
          orders: []
        };
      }
      
      stats[cat].count++;
      stats[cat].orders.push(order);
    });
    
    return stats;
  }
};



function normalizeEstado(estado) {
  if (!estado) return '';
  return String(estado)
    .toUpperCase()
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function toast(msg) {
  const t = document.createElement('div');
  t.className = 'toast';
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 4000);
}

function toggleUtilities() {
  const menu = document.getElementById('utilitiesMenu');
  if (!menu) return;
  menu.classList.toggle('show');
}

function closeUtilitiesMenu() {
  const menu = document.getElementById('utilitiesMenu');
  if (!menu) return;
  menu.classList.remove('show');
}

document.addEventListener('click', (e) => {
  const dropdown = document.querySelector('.utilities-dropdown');
  if (dropdown && !dropdown.contains(e.target)) {
    closeUtilitiesMenu();
  }
});

function openPlanillasNewTab() {
  closeUtilitiesMenu();

  const newWindow = window.open('', '_blank', 'width=1000,height=800');
  if (!newWindow) {
    toast('🔒 Habilitá las ventanas emergentes para abrir las herramientas de Planillas.');
    return;
  }
  
  const planillasHTML = `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>📋 Planillas - Panel Fulfillment</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: 'Segoe UI', sans-serif;
      background: #F3F3F3;
      padding: 20px;
    }
    .container {
      max-width: 900px;
      margin: 0 auto;
      background: white;
      padding: 30px;
      border-radius: 12px;
      box-shadow: 0 8px 16px rgba(0,0,0,0.12);
    }
    h1 {
      color: #0078D4;
      margin-bottom: 30px;
    }
    .form-group {
      margin-bottom: 20px;
    }
    label {
      display: block;
      font-weight: 600;
      margin-bottom: 8px;
    }
    select, textarea {
      width: 100%;
      padding: 10px;
      border: 1px solid #E1DFDD;
      border-radius: 8px;
    }
    textarea {
      min-height: 100px;
      font-family: 'Courier New', monospace;
    }
    .btn {
      padding: 12px 24px;
      border: none;
      border-radius: 8px;
      font-weight: 600;
      cursor: pointer;
      margin-right: 10px;
    }
    .btn-primary { background: #0078D4; color: white; }
    .btn-success { background: #107C10; color: white; }
    .output {
      background: #FAFAFA;
      padding: 16px;
      border-radius: 8px;
      margin-top: 20px;
      font-family: 'Courier New', monospace;
      white-space: pre-wrap;
      display: none;
    }
    .output.show { display: block; }
  </style>
</head>
<body>
  <div class="container">
    <h1>📋 Generador de Planillas Técnicas</h1>
    <div class="form-group">
      <label>Tipo de Planilla</label>
      <select id="tipo" onchange="updateForm()">
        <option value="">Selecciona...</option>
        <option value="niveles">Niveles Fuera de Rango</option>
        <option value="sinSenal">Sin Señal</option>
        <option value="ruido">Ruido en Red</option>
      </select>
    </div>
    <div id="campos"></div>
    <button class="btn btn-primary" onclick="generar()">🔧 Generar</button>
    <button class="btn btn-success" onclick="copiar()">📋 Copiar</button>
    <div class="output" id="output"></div>
  </div>
  <script>
    function updateForm() {
      const tipo = document.getElementById('tipo').value;
      const container = document.getElementById('campos');
      if (!tipo) { container.innerHTML = ''; return; }
      container.innerHTML = '<div class="form-group"><label>Observaciones</label><textarea id="obs"></textarea></div>';
    }
    function generar() {
      const tipo = document.getElementById('tipo').value;
      if (!tipo) return alert('Selecciona un tipo');
      const obs = document.getElementById('obs')?.value || '';
      const output = document.getElementById('output');
      output.textContent = \`REPORTE TÉCNICO
Fecha: \${new Date().toLocaleDateString()}
Tipo: \${tipo.toUpperCase()}

Observaciones:
\${obs}

Generado desde Panel Fulfillment v5.0\`;
      output.classList.add('show');
    }
    function copiar() {
      const text = document.getElementById('output').textContent;
      if (!text) return alert('Genera primero una planilla');
      navigator.clipboard.writeText(text).then(() => alert('✓ Copiado'));
    }
  </script>
</body>
</html>`;

  newWindow.document.write(planillasHTML);
  newWindow.document.close();
}

function openUsefulLinks() {
  closeUtilitiesMenu();

  const win = window.open('', '_blank', 'width=1200,height=860,scrollbars=yes');
  if (!win) {
    toast('🔒 Habilitá las ventanas emergentes para ver los links útiles.');
    return;
  }

  const baseHref = new URL('.', window.location.href).href;
  const sections = [
    {
      title: 'Panel Fulfillment',
      description: 'Accesos clave al entorno de análisis y soporte del equipo.',
      links: [
        { label: 'Analisis', url: 'http://10.120.52.24/analisis/', note: 'Panel principal de gestión' },
        { label: 'planillas', url: 'http://10.120.52.24/pagmdr/', note: 'Plantillas de carga diaria' },
        { label: 'Herramientas', url: 'http://10.120.52.24/analisis/herramientas_internas/', note: 'Utilidades internas' }
      ]
    },
    {
      title: 'Gestión & Monitoreo',
      description: 'Aplicaciones operativas para seguimiento y control.',
      links: [
        { label: 'ICD', url: 'https://teco.sccd.ibmserviceengage.com/maximo_ICD/ui/login', note: 'Incidentes y órdenes' },
        { label: 'ftth 1', url: 'http://10.9.44.132/symphonica/v2_10/#/', note: 'Symphonica FTTH' },
        { label: 'CCIP', url: 'https://ccip/index.php/login', note: 'Capa de Control IP' },
        { label: 'iTracker', url: 'https://itracker.telecom.com.ar/?L=index&m=menu', note: 'Seguimiento tickets' },
        { label: 'PortalNOC', url: 'https://supervision/portalNOC/login.asp', note: 'Monitoreo NOC' },
        { label: 'UnMacMe', url: 'https://unmacme.telecom.com.ar/', note: 'Gestión MACs cablemodem' },
        { label: 'TAU', url: 'https://tau.telecom.com.ar/#/ppf/home', note: 'Portal TAU' },
        { label: 'ServAssure NXT', url: 'https://nxt.telecom.arg.telecom.com.ar:8443/nxt-ui/#/search', note: 'Monitoreo HFC' },
        { label: 'FMS', url: 'http://fmbrms-prod.corp.cablevision.com.ar:8180/fms/web/#/login', note: 'Alarmas FMS' },
        { label: 'BeFan', url: 'https://snap.telecom.com.ar/befan/index.html#/signin', note: 'Gestión BeFan' }
      ]
    },
    {
      title: 'Recursos técnicos',
      description: 'Accesos de referencia para FTTH y bases de consulta.',
      links: [
        { label: 'cei ftth', url: 'http://pwxdatadb3/welcome/', note: 'Portal PWX CEI' },
        { label: 'Gpon Status', url: 'http://10.9.44.132/gpon_status/v1_0/inventario/#', note: 'Inventario GPON' },
        { label: 'Buscador Sector Operativo/Central', url: 'https://cablevisionfibertel.sharepoint.com/sites/FIELDSERVICE/formdec/SitePages/Buscador-por-Sector-Operativo.aspx', note: 'SharePoint Field Service' }
      ]
    }
  ];

  const sectionsHtml = sections.map(section => `
    <section class="links-section">
      <div class="links-section-header">
        <h2>${section.title}</h2>
        ${section.description ? `<p>${section.description}</p>` : ''}
      </div>
      <div class="links-grid">
        ${section.links.map(link => `
          <a class="link-card" href="${link.url}" target="_blank" rel="noopener">
            <span class="link-title">${link.label}</span>
            ${link.note ? `<span class="link-note">${link.note}</span>` : ''}
          </a>
        `).join('')}
      </div>
    </section>
  `).join('');

  win.document.write(`<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <base href="${baseHref}">
  <title>🔗 Links útiles - Panel Fulfillment</title>
  <style>
    :root {
      color-scheme: light dark;
    }
    body {
      margin: 0;
      font-family: 'Segoe UI', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
      background: #0f172a;
      color: #e2e8f0;
      min-height: 100vh;
    }
    .links-wrapper {
      max-width: 1100px;
      margin: 0 auto;
      padding: 40px 24px 80px;
    }
    header {
      display: flex;
      flex-direction: column;
      gap: 8px;
      margin-bottom: 32px;
    }
    header h1 {
      margin: 0;
      font-size: 2.25rem;
      color: #38bdf8;
    }
    header p {
      margin: 0;
      color: #cbd5f5;
      font-size: 0.95rem;
      max-width: 560px;
    }
    .links-section {
      background: rgba(15, 23, 42, 0.75);
      border: 1px solid rgba(148, 163, 184, 0.25);
      border-radius: 18px;
      padding: 24px 24px 28px;
      margin-bottom: 24px;
      box-shadow: 0 18px 40px rgba(15, 23, 42, 0.55);
    }
    .links-section-header h2 {
      margin: 0;
      font-size: 1.4rem;
      color: #f8fafc;
    }
    .links-section-header p {
      margin: 6px 0 0 0;
      color: #cbd5f5;
      font-size: 0.95rem;
    }
    .links-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(240px, 1fr));
      gap: 16px;
      margin-top: 20px;
    }
    .link-card {
      display: flex;
      flex-direction: column;
      gap: 6px;
      background: rgba(30, 41, 59, 0.9);
      border: 1px solid rgba(148, 163, 184, 0.2);
      border-radius: 14px;
      padding: 16px;
      text-decoration: none;
      color: inherit;
      transition: transform 0.18s ease, border-color 0.18s ease;
    }
    .link-card:hover {
      transform: translateY(-3px);
      border-color: rgba(56, 189, 248, 0.6);
      box-shadow: 0 12px 25px rgba(56, 189, 248, 0.25);
    }
    .link-title {
      font-weight: 600;
      font-size: 1.05rem;
      color: #f8fafc;
    }
    .link-note {
      font-size: 0.85rem;
      color: #94a3b8;
    }
    footer {
      margin-top: 32px;
      font-size: 0.8rem;
      color: rgba(148, 163, 184, 0.7);
      text-align: center;
    }
    @media (max-width: 640px) {
      header h1 { font-size: 1.8rem; }
      .links-section { padding: 20px 18px 24px; }
      .links-grid { grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); }
    }
  </style>
</head>
<body>
  <div class="links-wrapper">
    <header>
      <h1>🔗 Links útiles</h1>
      <p>Accesos verificados en las últimas 24 horas para acelerar el soporte operativo y técnico del equipo.</p>
    </header>
    ${sectionsHtml}
    <footer>Actualizado ${new Date().toLocaleDateString('es-AR')} • Panel Fulfillment v5.0</footer>
  </div>
</body>
</html>`);
  win.document.close();
}

class DataProcessor {
  constructor() {
    this.consolidado1 = null;
    this.consolidado2 = null;
    this.nodosData = null;
    this.fmsData = null;
    this.nodosMap = new Map();
    this.fmsMap = new Map();
    this.allColumns = new Set();
    this.fallbackIdCounter = 1;
  }

  getOrderId(row) {
    if (!row) return '';

    const preferredKeys = [
      'Número de cita',
      'Numero de cita',
      'N° de cita',
      'Número de caso',
      'Numero de caso',
      'N° de caso',
      'Número de cuenta',
      'Numero de cuenta',
      'Cuenta'
    ];

    for (const key of preferredKeys) {
      const val = row[key];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        return String(val).trim();
      }
    }

    const dynamicKey = Object.keys(row).find(k => {
      const normalized = stripAccents(k).toLowerCase();
      return normalized.includes('numero') && (normalized.includes('cita') || normalized.includes('caso') || normalized.includes('cuenta'));
    });

    if (dynamicKey) {
      const val = row[dynamicKey];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        return String(val).trim();
      }
    }

    return '';
  }
  
  async loadExcel(file, tipo) {
    try {
      const data = new Uint8Array(await file.arrayBuffer());

      const wb = XLSX.read(data, {
        type: 'array',
        cellDates: false,
        cellText: false,
        cellFormula: false,
        raw: false
      });

      if (!wb || !wb.SheetNames || !wb.SheetNames.length) {
        return {success:false, error:'Libro sin hojas'};
      }
      const ws = wb.Sheets[wb.SheetNames[0]];
      if (!ws) {
        return {success:false, error:'Hoja vacía o no encontrada'};
      }
      
      // Si no tiene rango definido, intentar inferirlo
      if (!ws['!ref']) {
        // Buscar el rango real escaneando celdas
        let maxR = 0, maxC = 0;
        for (const key in ws) {
          if (key[0] !== '!') {
            const cell = XLSX.utils.decode_cell(key);
            if (cell.r > maxR) maxR = cell.r;
            if (cell.c > maxC) maxC = cell.c;
          }
        }
        if (maxR > 0 || maxC > 0) {
          ws['!ref'] = XLSX.utils.encode_range({s:{r:0,c:0}, e:{r:maxR,c:maxC}});
          console.log(`⚠️ Rango inferido: ${ws['!ref']}`);
        } else {
          return {success:false, error:'Hoja sin datos detectables'};
        }
      }
      
      const range = XLSX.utils.decode_range(ws['!ref']);
      
      let headerRow = 0;
      for (let R = 0; R <= 20; R++) {
        let hasContent = false;
        for (let C = 0; C <= 15; C++) {
          const cellAddr = XLSX.utils.encode_cell({r: R, c: C});
          const cell = ws[cellAddr];
          if (cell && cell.v) {
            const cellValue = String(cell.v).toLowerCase();
            if (cellValue.includes('zona') && cellValue.includes('tecnica')) {
              headerRow = R;
              console.log(`✅ Headers detectados en fila ${R + 1} (columna ${C}): "${cell.v}"`);
              hasContent = true;
              break;
            }
            if (cellValue.includes('numero') && (cellValue.includes('caso') || cellValue.includes('cuenta'))) {
              headerRow = R;
              console.log(`✅ Headers detectados en fila ${R + 1} (columna ${C}): "${cell.v}"`);
              hasContent = true;
              break;
            }
          }
        }
        if (hasContent) break;
      }
      
      console.log(`📋 Usando fila ${headerRow + 1} como headers`);
      
      range.s.r = headerRow;
      ws['!ref'] = XLSX.utils.encode_range(range);
      
      const jsonData = XLSX.utils.sheet_to_json(ws, { 
        defval: '', 
        raw: false, 
        dateNF: 'dd/mm/yyyy' 
      });
      
      console.log(`✅ Archivo tipo ${tipo}: ${jsonData.length} registros`);
      
      if (tipo === 1) this.consolidado1 = jsonData;
      else if (tipo === 2) this.consolidado2 = jsonData;
      else if (tipo === 3) {
        this.nodosData = jsonData;
        this.processNodos();
      } else if (tipo === 4) {
        this.fmsData = jsonData;
        this.processFMS();
      }
      
      if (jsonData.length > 0) {
        Object.keys(jsonData[0]).forEach(col => this.allColumns.add(col));
      }
      
      return {success: true, rows: jsonData.length};
    } catch (e) {
      console.error('Error loading Excel:', e);
      return {success: false, error: e.message};
    }
  }
  
  async loadCSV(file) {
    try {
      const text = await file.text();
      const lines = text.split('\n').filter(l => l.trim());
      
      if (lines.length < 2) {
        return {success: false, error: 'CSV vacío'};
      }
      
      const parseCSVLine = (line) => {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
          const char = line[i];
          
          if (char === '"') {
            inQuotes = !inQuotes;
          } else if (char === ',' && !inQuotes) {
            result.push(current.trim());
            current = '';
          } else {
            current += char;
          }
        }
        result.push(current.trim());
        return result;
      };
      
      const headers = parseCSVLine(lines[0]);
      const rows = [];
      
      for (let i = 1; i < lines.length; i++) {
        const values = parseCSVLine(lines[i]);
        const obj = {};
        headers.forEach((h, idx) => {
          obj[h] = values[idx] || '';
        });
        rows.push(obj);
      }
      
      console.log(`✅ CSV Alarmas: ${rows.length} alarmas leídas`);
      this.fmsData = rows;
      this.processFMS();
      
      return {success: true, rows: rows.length};
    } catch (e) {
      console.error('Error loading CSV:', e);
      return {success: false, error: e.message};
    }
  }
  
  processNodos() {
    if (!this.nodosData) return;
    
    this.nodosData.forEach(n => {
      const nodo = String(n['Nodo'] || '').trim();
      if (!nodo || nodo === '-' || nodo.includes('SIN_NODO')) return;
      
      const up = parseFloat(n['Up']) || 0;
      const down = parseFloat(n['Down']) || 0;
      const cmts = String(n['CMTS'] || '').trim();
      
      let estado = 'up';
      if (down > 50) estado = 'critical';
      else if (down > 0) estado = 'down';
      
      this.nodosMap.set(nodo, {
        cmts,
        up,
        down,
        estado,
        total: up + down
      });
    });
    
    console.log(`✅ Procesados ${this.nodosMap.size} nodos`);
  }
  
  processFMS() {
    if (!this.fmsData) return;
    
    console.log('Procesando alarmas FMS...');
    
    this.fmsData.forEach(f => {
      const zonaTecnica = String(f['networkElement.technicalZone'] || '').trim();
      if (!zonaTecnica) return;
      
      const alarma = {
        eventId: f['eventId'] || '',
        type: f['type'] || '',
        creationDate: f['creationDate'] || '',
        recoveryDate: f['recoveryDate'] || '',
        elementType: f['networkElement.type'] || '',
        elementCode: f['networkElement.code'] || '',
        damage: f['damage'] || '',
        incidentClassification: f['incidentClassification'] || '',
        damageClassification: f['damageClassification'] || '',
        claims: f['claims'] || '',
        childCount: f['networkElement.childCount'] || '',
        cmCount: f['networkElement.cmCount'] || '',
        taskName: f['task.name'] || f['taskName'] || f['workTaskName'] || f['task'] || '',
        taskType: f['task.type'] || f['taskType'] || '',
        description: f['description'] || f['alarmDescription'] || '',
        isActive: !f['recoveryDate'] || f['recoveryDate'] === ''
      };
      
      if (!this.fmsMap.has(zonaTecnica)) {
        this.fmsMap.set(zonaTecnica, []);
      }
      this.fmsMap.get(zonaTecnica).push(alarma);
    });
    
    console.log(`✅ Procesadas alarmas para ${this.fmsMap.size} zonas técnicas`);
  }
  
  merge() {
    if (!this.consolidado1 && !this.consolidado2) return [];
    if (!this.consolidado1) return this.consolidado2 || [];
    if (!this.consolidado2) return this.consolidado1 || [];

    const map = new Map();
    let missingIdCount = 0;

    const ensureId = (value, prefix, idx) => {
      const id = value && String(value).trim();
      if (id) return id;
      missingIdCount++;
      return `${prefix}${idx + 1}`;
    };

    this.consolidado1.forEach((r, idx) => {
      const cita = ensureId(this.getOrderId(r), 'c1-', idx);
      if (!map.has(cita)) {
        map.set(cita, { ...r, _merged: false });
      }
    });

    this.consolidado2.forEach((r, idx) => {
      const cita = ensureId(this.getOrderId(r), 'c2-', idx);
      if (map.has(cita)) {
        const existing = map.get(cita);
        Object.keys(r).forEach(key => {
          if (!existing[key] || existing[key] === '') {
            existing[key] = r[key];
          }
        });
        existing._merged = true;
      } else {
        map.set(cita, { ...r, _merged: false });
      }
    });

    const result = Array.from(map.values());
    const suffix = missingIdCount > 0 ? ` (asignados ${missingIdCount} IDs faltantes)` : '';
    console.log(`✅ Total órdenes únicas (deduplicadas): ${result.length}${suffix}`);

    return result;
  }
  
  processZones(rows) {
    const zoneGroups = new Map();
    
    rows.forEach(r => {
      const {zonaPrincipal, tipo} = this.getZonaPrincipal(r);
      if (!zonaPrincipal) return;
      
      if (!zoneGroups.has(zonaPrincipal)) {
        zoneGroups.set(zonaPrincipal, {
          zona: zonaPrincipal,
          tipo: tipo,
          zonaHFC: r['Zona Tecnica HFC'] || r['Zona Tecnica'] || '',
          zonaFTTH: r['Zona Tecnica FTTH'] || '',
          territorio: r['Territorio de servicio: Nombre'] || r['Territorio'] || '',
          ordenes: []
        });
      }
      
      zoneGroups.get(zonaPrincipal).ordenes.push(r);
    });
    
    return Array.from(zoneGroups.values());
  }
  
  getZonaPrincipal(orden) {
    const hfc = orden['Zona Tecnica HFC'] || orden['Zona Tecnica'] || '';
    const ftth = orden['Zona Tecnica FTTH'] || '';
    
    const esFTTH = /9\d{2}/.test(hfc);
    
    if (esFTTH && ftth) {
      return {zonaPrincipal: ftth, tipo: 'FTTH'};
    }
    
    return {zonaPrincipal: hfc || ftth, tipo: 'HFC'};
  }
  
  analyzeZones(zones, daysWindow) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    console.log(`✅ Analizando zonas con ventana de ${daysWindow} días`);

    return zones.map(z => {
      const seenOTperDay = new Map();
      const ordenesPerDay = new Map();
      let ordenesEnVentana = 0;

      const diagnosticos = new Set();

      z.ordenes.forEach(o => {
        const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || o['Fecha de inicio'];
        const dt = DateUtils.parse(fecha);
        if (!dt) return;

        const daysAgo = DateUtils.daysBetween(today, dt);

        if (daysAgo <= daysWindow) {
          ordenesEnVentana++;

          const dk = DateUtils.toDayKey(dt);
          if (!seenOTperDay.has(dk)) seenOTperDay.set(dk, new Set());
          if (!ordenesPerDay.has(dk)) ordenesPerDay.set(dk, []);

          const cita = o['Número de cita'];
          if (cita) {
            seenOTperDay.get(dk).add(cita);
            ordenesPerDay.get(dk).push(o);
          }

          const diag = o['Diagnostico Tecnico'] || o['Diagnóstico Técnico'] || '';
          if (diag) diagnosticos.add(diag);
        }
      });

      const todayKey = DateUtils.toDayKey(today);
      const yesterdayKey = DateUtils.toDayKey(new Date(today.getTime() - 86400000));

      const ingresoN  = seenOTperDay.get(todayKey)?.size || 0;
      const ingresoN1 = seenOTperDay.get(yesterdayKey)?.size || 0;
      const maxDia = Math.max(0, ...Array.from(seenOTperDay.values()).map(s => s.size));

      const score = 4 * ingresoN + 2 * ingresoN1 + 1.5 * maxDia;
      const criticidad = score >= 12 ? 'CRÍTICO' : score >= 7 ? 'ALTO' : 'MEDIO';

      const last7Days = [];
      const last7DaysCounts = [];
      for (let i = 6; i >= 0; i--) {
        const date = new Date(today);
        date.setDate(date.getDate() - i);
        const dk = DateUtils.toDayKey(date);
        last7Days.push(DateUtils.format(date).slice(0, 5));
        last7DaysCounts.push(seenOTperDay.get(dk)?.size || 0);
      }

      let nodoInfo = null;
      if (this.nodosMap.has(z.zona)) {
        nodoInfo = this.nodosMap.get(z.zona);
      } else if (z.zonaHFC && this.nodosMap.has(z.zonaHFC)) {
        nodoInfo = this.nodosMap.get(z.zonaHFC);
      }

      const alarmasFMS = this.fmsMap.get(z.zona) ||
                         this.fmsMap.get(z.zonaHFC) ||
                         this.fmsMap.get(z.zonaFTTH) || [];
      const alarmasActivas = alarmasFMS.filter(a => a.isActive);

      return {
        ...z,
        totalOTs: ordenesEnVentana,
        ingresoN,
        ingresoN1,
        maxDia,
        score,
        criticidad,
        cmts: nodoInfo?.cmts || '',
        nodoUp: nodoInfo?.up || 0,
        nodoDown: nodoInfo?.down || 0,
        nodoEstado: nodoInfo?.estado || 'unknown',
        last7Days,
        last7DaysCounts,
        diagnosticos: Array.from(diagnosticos),
        alarmas: alarmasFMS,
        alarmasActivas: alarmasActivas.length,
        tieneAlarma: alarmasActivas.length > 0,
        ordenesOriginales: z.ordenes
      };
    }).sort((a, b) => b.score - a.score);
  }
  
  analyzeCMTS(zones) {
    const cmtsGroups = new Map();
    
    zones.forEach(z => {
      if (!z.cmts) return;
      
      if (!cmtsGroups.has(z.cmts)) {
        cmtsGroups.set(z.cmts, {
          cmts: z.cmts,
          zonas: [],
          totalOTs: 0,
          zonasUp: 0,
          zonasDown: 0,
          zonasCriticas: 0
        });
      }
      
      const group = cmtsGroups.get(z.cmts);
      group.zonas.push(z);
      group.totalOTs += z.totalOTs;
      
      if (z.nodoEstado === 'up') group.zonasUp++;
      else if (z.nodoEstado === 'down' || z.nodoEstado === 'critical') group.zonasDown++;
      
      if (z.nodoEstado === 'critical') group.zonasCriticas++;
    });
    
    return Array.from(cmtsGroups.values()).sort((a, b) => b.totalOTs - a.totalOTs);
  }
  
  analyzeTerritorios(zones) {
    const territoriosMap = new Map();
    
    zones.forEach(z => {
      const territorioOriginal = z.territorio;
      if (!territorioOriginal) return;
      
      const territorioNormalizado = TerritorioUtils.normalizar(territorioOriginal);
      const numeroFlex = TerritorioUtils.extraerFlex(territorioOriginal);
      
      if (!territoriosMap.has(territorioNormalizado)) {
        territoriosMap.set(territorioNormalizado, {
          territorio: territorioNormalizado,
          territorioOriginal: territorioOriginal,
          zonas: [],
          flexSubdivisiones: new Map(),
          totalOTs: 0,
          zonasConAlarma: 0,
          zonasCriticas: 0,
          zonasUp: 0,
          zonasDown: 0,
          esCritico: false
        });
      }
      
      const grupo = territoriosMap.get(territorioNormalizado);
      grupo.zonas.push(z);
      grupo.totalOTs += z.totalOTs;
      
      if (z.tieneAlarma) grupo.zonasConAlarma++;
      if (z.criticidad === 'CRÍTICO') grupo.zonasCriticas++;
      if (z.nodoEstado === 'up') grupo.zonasUp++;
      if (z.nodoEstado === 'down' || z.nodoEstado === 'critical') grupo.zonasDown++;
      
      if (numeroFlex !== null) {
        const flexKey = `Flex ${numeroFlex}`;
        if (!grupo.flexSubdivisiones.has(flexKey)) {
          grupo.flexSubdivisiones.set(flexKey, {
            nombre: flexKey,
            zonas: [],
            totalOTs: 0,
            zonasCriticas: 0
          });
        }
        
        const flexGrupo = grupo.flexSubdivisiones.get(flexKey);
        flexGrupo.zonas.push(z);
        flexGrupo.totalOTs += z.totalOTs;
        if (z.criticidad === 'CRÍTICO') flexGrupo.zonasCriticas++;
      }
      
      if (z.criticidad === 'CRÍTICO') {
        grupo.esCritico = true;
      }
    });
    
    return Array.from(territoriosMap.values())
      .sort((a, b) => b.totalOTs - a.totalOTs);
  }
}


function escapeHtml(str = '') {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function truncateText(str = '', maxLength = 60) {
  const text = String(str || '').trim();
  if (!text) return '';
  if (text.length <= maxLength) return text;
  return `${text.slice(0, maxLength - 1)}…`;
}

const zoneFilterState = {
  allOptions: [],
  filteredOptions: [],
  selected: new Set(),
  search: '',
  open: false,
  lastOptionsSignature: '',
  lastFilteredSignature: ''
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
        const estRaw =
          r['Estado.1'] ||
          r['Estado']   ||
          r['Estado.2'] ||
          '';

        const est = normalizeEstado(estRaw);
        if (!est) return false;

        if (estadosOcultosNorm.includes(est)) return false;
        return estadosPermitidosNorm.includes(est);
      });
    }
    
    const catecActive = this.catec;
    const excludeCatecActive = !catecActive && this.excludeCatec;

    if (catecActive) {
      filtered = filtered.filter(r => {
        const tipo = r['Tipo de trabajo: Nombre de tipo de trabajo'] || '';
        const upperTipo = String(tipo).toUpperCase();
        return upperTipo.includes('CATEC');
      });
    }

    if (excludeCatecActive) {
      filtered = filtered.filter(r => {
        const tipo = r['Tipo de trabajo: Nombre de tipo de trabajo'] || '';
        const upperTipo = String(tipo).toUpperCase();
        return !upperTipo.includes('CATEC');
      });
    }
    
    if (this.ftth) {
      filtered = filtered.filter(r => {
        const hfc = r['Zona Tecnica HFC'] || '';
        return /9\d{2}/.test(hfc);
      });
    }
    
    if (this.excludeFTTH) {
      filtered = filtered.filter(r => {
        const hfc = r['Zona Tecnica HFC'] || '';
        return !/9\d{2}/.test(hfc);
      });
    }
    
    if (this.territorio) {
      filtered = filtered.filter(r => {
        const terr = r['Territorio de servicio: Nombre'] || '';
        return terr === this.territorio;
      });
    }
    
    if (this.sistema) {
      filtered = filtered.filter(r => {
        const numCaso = r['Número del caso'] || r['Caso Externo'] || '';
        const sistema = TextUtils.detectarSistema(numCaso);
        return sistema === this.sistema;
      });
    }
    
    if (this.quickSearch) {
      const queries = this.quickSearch.split(';').map(q => q.trim()).filter(Boolean);

      filtered = filtered.filter(r => {
        const searchable = [
          r['Zona Tecnica HFC'],
          r['Zona Tecnica FTTH'],
          r['Calle'],
          r['Caso Externo'],
          r['External Case Id'],
          r['Número del caso']
        ].join(' ');
        
        if (queries.length === 1) {
          return TextUtils.matches(searchable, queries[0]);
        } else {
          return TextUtils.matchesMultiple(searchable, queries);
        }
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

    if (this.criticidad === 'critico') {
      filtered = filtered.filter(z => z.criticidad === 'CRÍTICO');
    }
    
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

    if (this.selectedZonas && this.selectedZonas.length) {
      const zonesSet = new Set(this.selectedZonas);
      filtered = filtered.filter(z => zonesSet.has(z.zona));
    }

    if (this.ordenarPorIngreso === 'desc') {
      filtered.sort((a, b) => b.totalOTs - a.totalOTs);
    } else if (this.ordenarPorIngreso === 'asc') {
      filtered.sort((a, b) => a.totalOTs - b.totalOTs);
    }
    
    return filtered;
  }
};

const zoneDamageSummaryCache = new WeakMap();

function extractDamageSummaryFromZone(zone) {
  if (!zone || typeof zone !== 'object') return [];

  if (zoneDamageSummaryCache.has(zone)) {
    return zoneDamageSummaryCache.get(zone).map(item => ({ ...item }));
  }

  let summary = [];

  if (Array.isArray(zone.damageSummary) && zone.damageSummary.length) {
    summary = zone.damageSummary.map(item => ({
      diagnostico: item.diagnostico,
      count: item.count
    }));
  } else {
    const counts = new Map();
    const ordenes = Array.isArray(zone.ordenesOriginales)
      ? zone.ordenesOriginales
      : Array.isArray(zone.ordenes)
        ? zone.ordenes
        : [];

    ordenes.forEach(o => {
      const diag = o['Diagnostico Tecnico'] ||
                   o['Diagnóstico Técnico'] ||
                   o['Diagnostico tecnico'] ||
                   o['Diagnostico'] ||
                   'Sin diagnóstico';
      const key = (diag && String(diag).trim()) || 'Sin diagnóstico';
      if (!counts.has(key)) {
        counts.set(key, { diagnostico: key, count: 0 });
      }
      counts.get(key).count += 1;
    });

    summary = Array.from(counts.values()).sort((a, b) => b.count - a.count);
  }

  zoneDamageSummaryCache.set(zone, summary);
  return summary.map(item => ({ ...item }));
}

function buildZoneFilterOptions(zones = []) {
  const optionsMap = new Map();

  (zones || []).forEach(item => {
    if (!item) return;

    if (typeof item === 'string') {
      const name = item.trim();
      if (!name) return;
      if (!optionsMap.has(name)) {
        optionsMap.set(name, {
          name,
          totalOTs: 0,
          damageSummary: []
        });
      }
      return;
    }

    const name = (item.zona || item.name || '').trim();
    if (!name) return;

    const total = Number.isFinite(item.totalOTs)
      ? item.totalOTs
      : Array.isArray(item.ordenesOriginales)
        ? item.ordenesOriginales.length
        : Array.isArray(item.ordenes)
          ? item.ordenes.length
          : 0;

    optionsMap.set(name, {
      name,
      totalOTs: total,
      damageSummary: extractDamageSummaryFromZone(item),
      tipo: item.tipo || '',
      criticidad: item.criticidad || '',
      tieneAlarma: Boolean(item.tieneAlarma)
    });
  });

  const options = Array.from(optionsMap.values());
  options.sort((a, b) => a.name.localeCompare(b.name, 'es', { sensitivity: 'base' }));
  return options;
}

function createZoneOptionsSignature(options = []) {
  const list = Array.isArray(options) ? options : [];

  return list.map(option => {
    const diagnostics = (option.damageSummary || [])
      .slice(0, 3)
      .map(d => `${d.diagnostico}:${d.count}`)
      .join('|');

    return [
      option.name || '',
      Number.isFinite(option.totalOTs) ? option.totalOTs : 0,
      option.tipo || '',
      option.criticidad || '',
      option.tieneAlarma ? '1' : '0',
      diagnostics
    ].join('~');
  }).join('||');
}

function setZoneFilterOptions(options, { force = false } = {}) {
  const signature = createZoneOptionsSignature(options);
  const changed = force || signature !== zoneFilterState.lastOptionsSignature;

  zoneFilterState.allOptions = options;

  if (changed) {
    zoneFilterState.lastOptionsSignature = signature;
  }

  return changed;
}

function setZoneFilterFilteredOptions(options, { force = false } = {}) {
  const signature = createZoneOptionsSignature(options);

  if (!force && signature === zoneFilterState.lastFilteredSignature) {
    zoneFilterState.filteredOptions = options;
    return false;
  }

  zoneFilterState.filteredOptions = options;
  zoneFilterState.lastFilteredSignature = signature;
  return true;
}

function filterZoneOptionsBySearch(options, term) {
  const normalized = TextUtils.normalize(term || '');
  if (!normalized) return options.slice();

  return options.filter(option => {
    if (TextUtils.normalize(option.name).includes(normalized)) return true;
    return option.damageSummary.some(d => TextUtils.normalize(d.diagnostico).includes(normalized));
  });
}

function initializeZoneFilterOptions(zones, options = {}) {
  const {
    resetSelection = true,
    resetSearch = resetSelection,
    forceRender = false
  } = options;

  const previousSelection = new Set(zoneFilterState.selected);
  const builtOptions = buildZoneFilterOptions(zones);
  const optionsChanged = setZoneFilterOptions(builtOptions, { force: forceRender });

  let selectionChanged = false;

  if (resetSelection) {
    selectionChanged = previousSelection.size > 0;
    zoneFilterState.selected = new Set();
  } else if (optionsChanged) {
    const validNames = new Set(builtOptions.map(opt => opt.name));
    const nextSelection = new Set();

    previousSelection.forEach(zoneName => {
      if (validNames.has(zoneName)) {
        nextSelection.add(zoneName);
      }
    });

    selectionChanged = nextSelection.size !== previousSelection.size;
    zoneFilterState.selected = nextSelection;
  }

  if (resetSearch) {
    if (zoneFilterState.search) {
      const searchInput = document.getElementById('zoneFilterSearch');
      if (searchInput) searchInput.value = '';
    }
    zoneFilterState.search = '';
  }

  const filteredList = filterZoneOptionsBySearch(
    zoneFilterState.allOptions,
    zoneFilterState.search
  );

  const filteredChanged = setZoneFilterFilteredOptions(filteredList, {
    force: forceRender || optionsChanged || resetSearch
  });

  if (optionsChanged || filteredChanged || forceRender) {
    renderZoneFilterOptions();
  }

  updateZoneFilterSummary();
  Filters.selectedZonas = Array.from(zoneFilterState.selected);

  return selectionChanged;
}

function renderZoneFilterOptions() {
  const container = document.getElementById('zoneFilterOptions');
  if (!container) return;

  if (!zoneFilterState.allOptions.length) {
    container.innerHTML = '<div class="zone-multiselect__empty">Cargá los reportes para habilitar el filtro por zonas</div>';
    return;
  }

  if (!zoneFilterState.filteredOptions.length) {
    container.innerHTML = '<div class="zone-multiselect__empty">No hay coincidencias para la búsqueda actual</div>';
    return;
  }

  const optionsHtml = zoneFilterState.filteredOptions.map(option => {
    const checked = zoneFilterState.selected.has(option.name) ? 'checked' : '';
    const safeName = escapeHtml(option.name);
    const total = Number.isFinite(option.totalOTs) ? option.totalOTs : 0;
    const totalLabel = `${total} OT${total === 1 ? '' : 's'}`;

    const diagnostics = (option.damageSummary || []).slice(0, 3);
    const diagnosticsHtml = diagnostics.length
      ? `<ul class="zone-option__diagnostics">${diagnostics.map(d => {
          const label = escapeHtml(truncateText(d.diagnostico, 48));
          const tooltip = escapeHtml(d.diagnostico || '');
          return `<li class="zone-option__diagnostic"><span class="zone-option__diagnostic-name" title="${tooltip}">${label}</span><span class="zone-option__diagnostic-count">${d.count}</span></li>`;
        }).join('')}</ul>`
      : '<div class="zone-option__diagnostics zone-option__diagnostics--empty">Sin órdenes técnicas en ventana</div>';

    return `
      <label class="zone-option">
        <input type="checkbox" value="${safeName}" ${checked}>
        <div class="zone-option__content">
          <div class="zone-option__header">
            <span class="zone-option__name" title="${safeName}">${safeName}</span>
            <span class="zone-option__count">${totalLabel}</span>
          </div>
          ${diagnosticsHtml}
        </div>
      </label>
    `;
  }).join('');

  container.innerHTML = optionsHtml;
}

function updateZoneFilterSummary() {
  const summaryEl = document.getElementById('zoneFilterSummary');
  const trigger = document.getElementById('zoneFilterTrigger');
  if (!summaryEl) return;

  const hasOptions = zoneFilterState.allOptions.length > 0;
  if (trigger) trigger.disabled = !hasOptions;

  if (!hasOptions) {
    summaryEl.textContent = 'Sin datos (carga reportes)';
    summaryEl.title = 'Carga los reportes para habilitar el filtro por zonas';
    return;
  }

  const count = zoneFilterState.selected.size;

  if (count === 0) {
    summaryEl.textContent = 'Todas las zonas';
    summaryEl.title = 'Mostrar todas las zonas disponibles';
  } else if (count === 1) {
    const [only] = Array.from(zoneFilterState.selected);
    summaryEl.textContent = only;
    summaryEl.title = only;
  } else {
    summaryEl.textContent = `${count} zonas seleccionadas`;
    const listPreview = Array.from(zoneFilterState.selected).slice(0, 8).join(', ');
    summaryEl.title = listPreview;
  }
}

function toggleZoneFilter(event) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }

  const container = document.getElementById('zoneFilter');
  if (!container) return;

  const willOpen = !container.classList.contains('open');
  if (willOpen) {
    container.classList.add('open');
    zoneFilterState.open = true;
    renderZoneFilterOptions();

    setTimeout(() => {
      const searchInput = document.getElementById('zoneFilterSearch');
      if (searchInput) {
        searchInput.focus();
        if (zoneFilterState.search) {
          const pos = zoneFilterState.search.length;
          searchInput.setSelectionRange(pos, pos);
        }
      }
    }, 0);
  } else {
    closeZoneFilter();
  }
}

function closeZoneFilter() {
  const container = document.getElementById('zoneFilter');
  if (!container) return;
  container.classList.remove('open');
  zoneFilterState.open = false;
}

function handleZoneFilterOutsideClick(event) {
  const container = document.getElementById('zoneFilter');
  if (!container) return;
  if (!container.contains(event.target)) {
    closeZoneFilter();
  }
}

const updateZoneFilterSearch = debounce(value => {
  const nextValue = value || '';
  if (nextValue === zoneFilterState.search) return;

  zoneFilterState.search = nextValue;
  const filtered = filterZoneOptionsBySearch(zoneFilterState.allOptions, nextValue);
  setZoneFilterFilteredOptions(filtered, { force: true });
  renderZoneFilterOptions();
}, 150);

function onZoneFilterSearch(event) {
  updateZoneFilterSearch(event.target.value || '');
}

function onZoneOptionChange(event) {
  const target = event.target;
  if (!target || target.type !== 'checkbox') return;

  const zone = target.value;
  if (!zone) return;

  if (target.checked) {
    zoneFilterState.selected.add(zone);
  } else {
    zoneFilterState.selected.delete(zone);
  }

  Filters.selectedZonas = Array.from(zoneFilterState.selected);
  updateZoneFilterSummary();
  applyFilters();
}

function selectAllZones(event) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }

  if (!zoneFilterState.allOptions.length) return;

  const hasSearch = Boolean(zoneFilterState.search);
  const source = hasSearch ? zoneFilterState.filteredOptions : zoneFilterState.allOptions;
  if (!source.length) return;

  source.forEach(option => zoneFilterState.selected.add(option.name));

  Filters.selectedZonas = Array.from(zoneFilterState.selected);
  renderZoneFilterOptions();
  updateZoneFilterSummary();
  applyFilters();
}

function resetZoneFilterState(options = {}) {
  const { apply = false, keepSearch = false } = options;

  zoneFilterState.selected = new Set();

  if (!keepSearch) {
    zoneFilterState.search = '';
    setZoneFilterFilteredOptions(zoneFilterState.allOptions.slice(), { force: true });
    const searchInput = document.getElementById('zoneFilterSearch');
    if (searchInput) searchInput.value = '';
  } else {
    const filtered = filterZoneOptionsBySearch(
      zoneFilterState.allOptions,
      zoneFilterState.search
    );
    setZoneFilterFilteredOptions(filtered, { force: true });
  }

  renderZoneFilterOptions();
  updateZoneFilterSummary();
  Filters.selectedZonas = [];

  if (apply) {
    applyFilters();
  }
}

function clearZoneSelection(event) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }

  resetZoneFilterState({ apply: true, keepSearch: true });
}

const UIRenderer = {
  renderStats(data) {
    return `
      <div class="stat-card clickable" style="cursor:pointer" onclick="filterByStat('total')">
        <div class="stat-label">Total Órdenes</div>
        <div class="stat-value">${NumberUtils.format(data.total)}</div>
      </div>
      <div class="stat-card clickable" style="cursor:pointer" onclick="filterByStat('zonas')">
        <div class="stat-label">Zonas Analizadas</div>
        <div class="stat-value">${NumberUtils.format(data.zonas)}</div>
      </div>
      <div class="stat-card clickable" style="cursor:pointer" onclick="filterByStat('territoriosCriticos')">
        <div class="stat-label">Territorios Críticos</div>
        <div class="stat-value">${NumberUtils.format(data.territoriosCriticos)}</div>
      </div>
      <div class="stat-card clickable" style="cursor:pointer" onclick="filterByStat('ftth')">
        <div class="stat-label">Zonas FTTH</div>
        <div class="stat-value">${NumberUtils.format(data.ftth)}</div>
      </div>
      <div class="stat-card clickable" style="cursor:pointer" onclick="filterByStat('alarmas')">
        <div class="stat-label">Con Alarmas</div>
        <div class="stat-value">${NumberUtils.format(data.conAlarmas)}</div>
      </div>
      <div class="stat-card clickable" style="cursor:pointer" onclick="filterByStat('nodosCriticos')">
        <div class="stat-label">Nodos Críticos</div>
        <div class="stat-value">${NumberUtils.format(data.nodosCriticos)}</div>
      </div>
    `;
  },
  
  normalizeCounts(counts) {
    if (!Array.isArray(counts)) return [];

    return counts.map(c => {
      if (c === null || c === undefined) return 0;
      if (typeof c === 'number') {
        return Number.isFinite(c) ? c : 0;
      }

      const parsed = parseFloat(String(c).trim().replace(',', '.'));
      return Number.isFinite(parsed) ? parsed : 0;
    });
  },

  renderSparkline(counts, labels) {
    const safeCounts = this.normalizeCounts(counts);

    if (!safeCounts.length) {
      return '<div class="sparkline-placeholder">Sin datos</div>';
    }

    const width = 120;
    const height = 30;
    const max = Math.max(...safeCounts, 1);
    const barWidth = width / safeCounts.length;

    let svg = `<svg class="sparkline" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">`;

    safeCounts.forEach((count, i) => {
      const normalized = max > 0 ? (count / max) : 0;
      const barHeight = normalized * height;
      const x = i * barWidth;
      const y = height - barHeight;
      const color = count >= max * 0.7 ? '#D13438' : count >= max * 0.4 ? '#F7630C' : '#0078D4';

      svg += `<rect x="${x}" y="${y}" width="${Math.max(barWidth - 2, 1)}" height="${barHeight}" fill="${color}" rx="2"/>`;
    });

    svg += '</svg>';
    return svg;
  },
  
  renderZonas(zones) {
    if (!zones.length) return '<div class="loading-message"><p>No hay zonas para mostrar</p></div>';
    
    let html = '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
    html += '<th>Zona</th><th>Tipo</th><th>Red</th><th>CMTS</th><th>Nodo</th>';
    html += '<th>Alarma</th><th>Gráfico 7 Días</th>';
    html += '<th class="number">Total</th><th class="number">N</th><th class="number">N-1</th>';
    html += '<th>Acción</th>';
    html += '</tr></thead><tbody>';
    
    zones.slice(0, 200).forEach((z, idx) => {
      const badgeTipo = z.tipo === 'FTTH'
        ? '<span class="badge badge-ftth">FTTH</span>'
        : '<span class="badge badge-hfc">HFC</span>';
      const red = z.tipo === 'FTTH' ? z.zonaHFC : '-';
      
      let badgeNodo = '<span class="badge">Sin datos</span>';
      if (z.nodoEstado === 'up') {
        badgeNodo = `<span class="badge badge-up">✓ UP (${z.nodoUp})</span>`;
      } else if (z.nodoEstado === 'critical') {
        badgeNodo = `<span class="badge badge-critical">⚠ CRÍTICO (↓${z.nodoDown})</span>`;
      } else if (z.nodoEstado === 'down') {
        badgeNodo = `<span class="badge badge-down">↓ DOWN (${z.nodoDown})</span>`;
      }
      
      let badgeAlarma = '<span class="badge">Sin alarma</span>';
      if (z.tieneAlarma) {
        badgeAlarma = `<span class="badge badge-alarma badge-alarma-activa" onclick="showAlarmaInfo(${idx})">🚨 ${z.alarmasActivas} Activa(s)</span>`;
      }
      
      const cmtsShort = z.cmts ? z.cmts.substring(0, 15) + (z.cmts.length > 15 ? '...' : '') : '-';
      const sparkline = this.renderSparkline(z.last7DaysCounts, z.last7Days);
      
      html += `<tr>
        <td><strong>${z.zona}</strong></td>
        <td>${badgeTipo}</td>
        <td>${red}</td>
        <td title="${z.cmts}">${cmtsShort}</td>
        <td>${badgeNodo}</td>
        <td>${badgeAlarma}</td>
        <td>${sparkline}</td>
        <td class="number">${z.totalOTs}</td>
        <td class="number">${z.ingresoN}</td>
        <td class="number">${z.ingresoN1}</td>
        <td>
          <div style="display:flex; gap:4px;">
            <button class="btn btn-primary" style="padding: 6px 12px; font-size: 0.75rem;" onclick="openModal(${idx})">👁️ Ver</button>
            <button class="btn btn-secondary" style="padding: 6px 12px; font-size: 0.75rem;" onclick="event.stopPropagation(); exportZonaExcel(${idx})">📥 Excel</button>
          </div>
        </td>
      </tr>`;
    });
    
    html += '</tbody></table></div></div>';
    return html;
  },

  renderCMTS(cmtsData) {
    if (!cmtsData.length) return '<div class="loading-message"><p>No hay datos de CMTS para mostrar</p></div>';
    
    let html = '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
    html += '<th>CMTS</th><th class="number">Zonas</th><th class="number">Total OTs</th>';
    html += '<th class="number">Zonas UP</th><th class="number">Zonas DOWN</th><th class="number">Zonas Críticas</th>';
    html += '<th>Acción</th>';
    html += '</tr></thead><tbody>';
    
    cmtsData.forEach((c, idx) => {
      const cmtsShort = c.cmts.substring(0, 30) + (c.cmts.length > 30 ? '...' : '');
      
      html += `<tr>
        <td title="${c.cmts}"><strong>${cmtsShort}</strong></td>
        <td class="number">${c.zonas.length}</td>
        <td class="number">${c.totalOTs}</td>
        <td class="number">${c.zonasUp}</td>
        <td class="number">${c.zonasDown}</td>
        <td class="number">${c.zonasCriticas}</td>
        <td>
          <div style="display:flex; gap:4px;">
            <button class="btn btn-primary" style="padding: 6px 12px; font-size: 0.75rem;" onclick="openCMTSDetail('${c.cmts.replace(/'/g, "\\'")}')">👁️ Ver Zonas</button>
            <button class="btn btn-secondary" style="padding: 6px 12px; font-size: 0.75rem;" onclick="event.stopPropagation(); exportCMTSExcel('${c.cmts.replace(/'/g, "\\'")}')">📥 Excel</button>
          </div>
        </td>
      </tr>`;
    });
    
    html += '</tbody></table></div></div>';
    return html;
  },
  
  renderEdificios(ordenes) {
    const fmsMap = dataProcessor ? dataProcessor.fmsMap : new Map();
    return EdificiosMejorado.render(ordenes, fmsMap);
  },


  
  renderEquipos(ordenes) {
    const grupos = new Map();
    const territorios = new Set();

    ordenes.forEach((o) => {
      const zona = o['Zona Tecnica HFC'] || o['Zona Tecnica FTTH'] || '';
      const territorio = o['Territorio de servicio: Nombre'] || '';
      const numCaso = o['Número del caso'] || o['Caso Externo'] || '';
      const sistema = TextUtils.detectarSistema(numCaso);

      const colInfo = findDispositivosColumn(o);
      const infoDispositivos = colInfo ? o[colInfo] : '';

      const dispositivos = TextUtils.parseDispositivosJSON(infoDispositivos);
      dispositivos.forEach(d => {
        const item = {
          zona,
          territorio,
          numCaso,
          sistema,
          serialNumber: d.serialNumber,
          macAddress: d.macAddress,
          tipo: d.description,
          marca: d.category,
          modelo: d.model
        };
        if (!grupos.has(zona)) grupos.set(zona, []);
        grupos.get(zona).push(item);
        if (territorio) territorios.add(territorio);
      });
    });

    const modelos = new Set();
    const marcas = new Set();
    grupos.forEach(arr => {
      arr.forEach(e => {
        if (e.modelo) modelos.add(e.modelo);
        if (e.marca) marcas.add(e.marca);
      });
    });

    const zonas = Array.from(grupos.keys()).sort((a, b) => a.localeCompare(b));
    if (!zonas.length) {
      return '<div class="loading-message"><p>⚠️ No se encontraron equipos en las órdenes.</p></div>';
    }

    if (!window.equiposOpen) window.equiposOpen = new Set();

    let html = '<div class="equipos-filters">';
    
    html += '<div class="filter-group">';
    html += '<div class="filter-label">Filtrar por Modelo</div>';
    html += '<select id="filterEquipoModelo" class="input-select" multiple size="3" onchange="applyEquiposFilters()">';
    Array.from(modelos).sort().forEach(m => {
      const selected = Filters.equipoModelo.includes(m) ? 'selected' : '';
      html += `<option value="${m}" ${selected}>${m || '(Sin modelo)'}</option>`;
    });
    html += '</select>';
    html += '</div>';

    html += '<div class="filter-group">';
    html += '<div class="filter-label">Filtrar por Marca</div>';
    html += '<select id="filterEquipoMarca" class="input-select" onchange="applyEquiposFilters()">';
    html += '<option value="">Todas las marcas</option>';
    Array.from(marcas).sort().forEach(m => {
      const selected = Filters.equipoMarca === m ? 'selected' : '';
      html += `<option value="${m}" ${selected}>${m || '(Sin marca)'}</option>`;
    });
    html += '</select>';
    html += '</div>';

    html += '<div class="filter-group">';
    html += '<div class="filter-label">Filtrar por Territorio</div>';
    html += '<select id="filterEquipoTerritorio" class="input-select" onchange="applyEquiposFilters()">';
    html += '<option value="">Todos los territorios</option>';
    Array.from(territorios).sort().forEach(t => {
      const selected = Filters.equipoTerritorio === t ? 'selected' : '';
      html += `<option value="${t}" ${selected}>${t}</option>`;
    });
    html += '</select>';
    html += '</div>';

    html += '<div class="filter-group">';
    html += '<button class="btn btn-secondary" onclick="clearEquiposFilters()" style="margin-top: 18px;">🔄 Limpiar filtros</button>';
    html += '</div>';
    html += '</div>';

    let gruposFiltrados = new Map();
    grupos.forEach((arr, zona) => {
      let filtrado = arr;
      
      if (Filters.equipoModelo.length > 0) {
        filtrado = filtrado.filter(e => Filters.equipoModelo.includes(e.modelo));
      }
      
      if (Filters.equipoMarca) {
        filtrado = filtrado.filter(e => e.marca === Filters.equipoMarca);
      }

      if (Filters.equipoTerritorio) {
        filtrado = filtrado.filter(e => e.territorio === Filters.equipoTerritorio);
      }
      
      if (filtrado.length > 0) {
        gruposFiltrados.set(zona, filtrado);
      }
    });

    const zonasFiltradas = Array.from(gruposFiltrados.keys()).sort((a, b) => a.localeCompare(b));
    
    if (zonasFiltradas.length === 0) {
      html += '<div class="loading-message"><p>No hay equipos que coincidan con los filtros seleccionados</p></div>';
      return html;
    }

    html += '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
    html += '<th>Zona</th><th class="number">Equipos</th><th>Acción</th>';
    html += '</tr></thead><tbody>';

    zonasFiltradas.forEach((z) => {
      const arr = gruposFiltrados.get(z) || [];
      const open = window.equiposOpen.has(z);
      html += `<tr class="clickable" onclick="toggleEquiposGrupo('${z.replace(/'/g, "\\'")}')">
        <td><strong>${z || '-'}</strong></td>
        <td class="number">${arr.length}</td>
        <td><button class="btn btn-secondary" style="padding:6px 12px; font-size:0.75rem;" onclick="event.stopPropagation(); exportEquiposGrupoExcel('${z.replace(/'/g, "\\'")}', true)">📥 Exportar</button></td>
      </tr>`;

      if (open) {
        html += `<tr><td colspan="3">`;
        html += '<div class="table-container"><div class="table-wrapper"><table class="detail-table"><thead><tr>';
        html += '<th>Caso</th><th>Sistema</th><th>Serial</th><th>MAC</th><th>Tipo</th><th>Marca</th><th>Modelo</th>';
        html += '</tr></thead><tbody>';
        arr.slice(0, 1000).forEach(e => {
          const badge = e.sistema === 'OPEN' ? '<span class="badge badge-open">OPEN</span>'
                       : e.sistema === 'FAN' ? '<span class="badge badge-fan">FAN</span>' : '';
          html += `<tr>
            <td style="font-family:monospace">${e.numCaso || ''}</td>
            <td>${badge}</td>
            <td style="font-family:monospace">${e.serialNumber || ''}</td>
            <td style="font-family:monospace">${e.macAddress || ''}</td>
            <td>${e.tipo || ''}</td>
            <td>${e.marca || ''}</td>
            <td>${e.modelo || ''}</td>
          </tr>`;
        });
        html += '</tbody></table></div></div>';
        html += `</td></tr>`;
      }
    });

    html += '</tbody></table></div></div>';
    html += `<p style="text-align:center;margin-top:12px;color:var(--text-secondary)">Total: ${zonasFiltradas.length} zonas • Click para expandir/colapsar</p>`;
    
    window.equiposPorZona = gruposFiltrados;
    window.equiposPorZonaCompleto = grupos;
    
    return html;
  }
};

const EdificiosMejorado = {
  /**
   * Renderiza edificios con CATEC, FMS y panel de ingreso por día
   * @param {Array} ordenes - Órdenes técnicas
   * @param {Map} fmsMap - Mapa de alarmas FMS
   * @returns {string} - HTML
   */
  render(ordenes, fmsMap) {
    const edificios = new Map();
    
    ordenes.forEach(o => {
      const dir = TextUtils.normalize(o['Calle']);
      if (!dir || dir.length < 5) return;
      
      if (!edificios.has(dir)) {
        edificios.set(dir, {
          direccion: o['Calle'],
          zona: o['Zona Tecnica HFC'] || o['Zona Tecnica FTTH'] || '',
          territorio: o['Territorio de servicio: Nombre'] || '',
          casos: [],
          casosCATEC: [],
          casosNoCATEC: [],
          alarmas: [],
          ingresosPorDia: new Map()
        });
      }
      
      const edificio = edificios.get(dir);
      edificio.casos.push(o);
      
      // Clasificar CATEC / No CATEC
      const tipoTrabajo = String(o['Tipo de trabajo: Nombre de tipo de trabajo'] || '').toUpperCase();
      if (tipoTrabajo.includes('CATEC')) {
        edificio.casosCATEC.push(o);
      } else {
        edificio.casosNoCATEC.push(o);
      }
      
      // Relacionar con alarmas FMS
      const zona = edificio.zona;
      if (zona && fmsMap && fmsMap.has(zona)) {
        const alarmasZona = fmsMap.get(zona) || [];
        edificio.alarmas = alarmasZona;
        edificio.alarmasActivas = alarmasZona.filter(a => a.isActive);
      } else {
        edificio.alarmas = [];
        edificio.alarmasActivas = [];
      }
      
      // Panel de ingreso por día
      const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || '';
      if (fecha) {
        const fechaStr = String(fecha).split(' ')[0]; // Extraer solo la fecha
        if (!edificio.ingresosPorDia.has(fechaStr)) {
          edificio.ingresosPorDia.set(fechaStr, []);
        }
        edificio.ingresosPorDia.get(fechaStr).push(o);
      }
    });
    
    const sorted = Array.from(edificios.values())
      .filter(e => e.casos.length >= 2)
      .sort((a, b) => b.casos.length - a.casos.length)
      .slice(0, 50);
    
    if (!sorted.length) {
      return '<div class="loading-message"><p>No hay edificios con 2+ incidencias</p></div>';
    }
    
    // Guardar en window para exportación
    window.edificiosData = sorted;
    
    let html = '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
    html += '<th>Dirección</th><th>Zona</th><th>Territorio</th>';
    html += '<th class="number">Total OTs</th>';
    html += '<th class="number">CATEC</th>';
    html += '<th class="number">No CATEC</th>';
    html += '<th class="number">🚨 Alarmas</th>';
    html += '<th>Acción</th>';
    html += '</tr></thead><tbody>';
    
    sorted.forEach((e, idx) => {
      const activas = e.alarmasActivas?.length || 0;
      const totalAlarmas = e.alarmas?.length || 0;
      const hasAlarmas = activas > 0 || totalAlarmas > 0;
      const alarmasClass = activas > 0 ? 'badge-alarma-activa' : '';
      
      html += `<tr>
        <td><strong>${e.direccion}</strong></td>
        <td>${e.zona}</td>
        <td>${e.territorio}</td>
        <td class="number">${e.casos.length}</td>
        <td class="number"><span class="badge badge-success">${e.casosCATEC.length}</span></td>
        <td class="number"><span class="badge badge-warning">${e.casosNoCATEC.length}</span></td>
        <td class="number">
          ${hasAlarmas ? `<span class="badge badge-alarma ${alarmasClass}" title="${activas} activas / ${totalAlarmas} totales">${activas || totalAlarmas}</span>` : '-'}
        </td>
        <td>
          <button class="btn btn-primary" style="padding: 6px 12px; font-size: 0.75rem;"
                  onclick="showEdificioDetail(${idx})">
            👁️ Ver Detalle
          </button>
        </td>
      </tr>`;
    });
    
    html += '</tbody></table></div></div>';
    return html;
  }
};


// C) NUEVA PESTAÑA FMS CON FILTRADO COMPLETO
// ═══════════════════════════════════════════════════════════════════

/* ===================== FMS HELPERS ===================== */
function renderFMSMessage(text, icon='ℹ️'){
  return `<div class="alert alert-info d-flex align-items-center mb-0">
    <span style="font-size:18px;margin-right:8px">${icon}</span> ${text}
  </div>`;
}

function normalizeText(s){
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase().trim();
}

function formatFMSTypeLabel(t){
  const v = String(t || '').toUpperCase();
  if (v.includes('CMTS')) return 'CMTS';
  if (v.includes('PUERTO')) return 'Puerto';
  if (v.includes('NODO')) return 'Nodo';
  if (v.includes('OPTICO') || v.includes('ÓPTICO')) return 'Óptico';
  if (!v) return 'Elemento';
  return t;
}

function formatAlarmaDamage(a){
  // Soporta estructuras variadas
  const c = a && (a.causal || a.Causal || a.damage || a.Daño || a.Danio || a['Tipo de daño'] || a.damageClassification || a.incidentClassification);
  const s = (typeof c === 'string' ? c : (c && c.tipo)) || '';
  return s || 'Sin dato';
}

/* ================ FMS PANEL (robusto y sin cuelgues) ================ */
const FMSPanel = {
  /**
   * Renderiza la pestaña de FMS con filtros.
   * @param {Array} _ordenes - compatible (no usado aquí)
   * @param {Map} fmsMap - Mapa de alarmas FMS (por zona->alarmas[] o por id->obj)
   */
  render(_ordenes, fmsMap) {
    if (!fmsMap || typeof fmsMap.size !== 'number' || fmsMap.size === 0) {
      return renderFMSMessage('No hay datos de FMS/Alarmas cargados', '⚠️');
    }

    const fmsGroups = this.groupByFMS(fmsMap);
    const damageStats = this.getDamageStats(fmsGroups);
    const filterOptions = this.buildFilterOptions(fmsGroups, damageStats);

    // Top defensivo para no bloquear UI
    const topGroups = fmsGroups.slice(0, 100);

    let html = `
      <div class="fms-panel">
        <div class="card text-white bg-warning shadow-sm mb-3">
          <div class="card-body">
            <h2 class="card-title h4 mb-2">🚨 Panel FMS - Alarmas y Daños</h2>
            <p class="card-text mb-0 small">
              Elementos monitoreados: ${fmsGroups.length} • Total registros en mapa: ${fmsMap.size}
            </p>
          </div>
        </div>

        <div class="card shadow-sm mb-4">
          <div class="card-body">
            <div class="form-row">
              <div class="form-group col-12 col-md-6 col-lg-4">
                <label class="font-weight-bold text-dark mb-1">🔍 Filtrar por Tipo de Elemento FMS</label>
                <select id="fmsTipoElemento" class="form-control" onchange="filtrarFMS()">
                  ${filterOptions.tipoOptions}
                </select>
              </div>
              <div class="form-group col-12 col-md-6 col-lg-4">
                <label class="font-weight-bold text-dark mb-1">⚠️ Filtrar por Daño / Causal</label>
                <select id="fmsDamage" class="form-control" onchange="filtrarFMS()">
                  ${filterOptions.damageOptions}
                </select>
              </div>
            </div>
          </div>
        </div>

        <div id="fmsListContainer">
          ${this.renderFMSList(topGroups, { limitNotice: true })}
        </div>
      </div>
    `;

    // Exponer en window para filtrado y detalle
    window.fmsGroupsData = fmsGroups;

    return html;
  },

  /** Consolida estadísticas globales de daños (para combo) */
  getDamageStats(fmsGroups) {
    const damageCount = new Map();
    fmsGroups.forEach(g => {
      g.damageSummary.forEach((count, damage) => {
        damageCount.set(damage, (damageCount.get(damage) || 0) + count);
      });
    });
    return Array.from(damageCount.entries()).sort((a, b) => b[1] - a[1]);
  },

  /**
   * Agrupa alarmas por elemento FMS y normaliza estructura
   * Soporta:
   *  - Map<zona, Alarma[]> donde Alarma = { elementCode, elementType, isActive, ... }
   *  - Map<id, { nombre, zonasAfectadas:[], causales:[{tipo}], ... }>
   */
  groupByFMS(fmsMap) {
    const groups = new Map();

    fmsMap.forEach((val, key) => {
      // Caso A: Map<zona, alarmas[]>
      if (Array.isArray(val)) {
        const zona = key;
        val.forEach(alarma => {
          const elementCode = alarma.elementCode || alarma.code || alarma.id || 'Sin código';
          const elementType = alarma.elementType || alarma.type || 'Elemento';
          const isActive = !!(alarma.isActive || alarma.activa || alarma.estado === 'ACTIVA');

          const gkey = `${elementCode}__${elementType}`;
          if (!groups.has(gkey)) {
            groups.set(gkey, {
              id: gkey,
              elementCode,
              elementType,
              alarmas: [],
              zonasAfectadas: new Set(),
              damageSummary: new Map()
            });
          }
          const g = groups.get(gkey);
          g.alarmas.push(alarma);
          g.zonasAfectadas.add(zona);
          const dKey = formatAlarmaDamage(alarma);
          g.damageSummary.set(dKey, (g.damageSummary.get(dKey) || 0) + 1);
          if (isActive) g.__active = (g.__active || 0) + 1;
        });
      }
      // Caso B: Map<id, objeto consolidado>
      else if (val && typeof val === 'object') {
        const id = val.id || key;
        const elementCode = val.elementCode || val.nombre || id || 'Sin código';
        const elementType = val.elementType || val.tipo || 'Elemento';
        const zonas = Array.isArray(val.zonasAfectadas) ? val.zonasAfectadas : [];
        const causales = Array.isArray(val.causales) ? val.causales : [];

        const gkey = `${elementCode}__${elementType}`;
        if (!groups.has(gkey)) {
          groups.set(gkey, {
            id: gkey,
            elementCode,
            elementType,
            alarmas: [],
            zonasAfectadas: new Set(),
            damageSummary: new Map()
          });
        }
        const g = groups.get(gkey);
        zonas.forEach(z => g.zonasAfectadas.add(z));
        causales.forEach(c => {
          const dKey = formatAlarmaDamage(c);
          g.damageSummary.set(dKey, (g.damageSummary.get(dKey) || 0) + 1);
        });
        // No sabemos cuántas están activas: estimación 0
        g.__active = g.__active || 0;
      }
    });

    const enriched = Array.from(groups.values()).map(g => {
      const activeCount = g.__active || 0;
      const totalAlarmas =
        g.alarmas?.length ??
        Array.from(g.damageSummary.values()).reduce((acc, n) => acc + (n || 0), 0);
      return {
        ...g,
        activeCount,
        totalAlarmas
      };
    }).sort((a, b) => {
      if (b.activeCount !== a.activeCount) return b.activeCount - a.activeCount;
      if (b.totalAlarmas !== a.totalAlarmas) return b.totalAlarmas - a.totalAlarmas;
      return String(a.elementCode).localeCompare(String(b.elementCode));
    });

    enriched.forEach((g, idx) => g.rank = idx + 1);
    return enriched;
  },

  /** Lista de elementos FMS (tarjetas) */
  renderFMSList(fmsGroups, options = {}) {
    const { limitNotice = false } = options;

    if (!fmsGroups.length) {
      return '<div class="alert alert-info mb-0">Sin resultados para los filtros aplicados</div>';
    }

    let html = '';
    if (limitNotice) {
      html += '<div class="alert alert-light border mb-3">Vista resumida. Ajustá la búsqueda para ver otros elementos.</div>';
    }

    html += '<div class="fms-list">';

    fmsGroups.forEach((group, idx) => {
      const tipoLabel = formatFMSTypeLabel(group.elementType);
      const alarmasActivas = group.activeCount || 0;
      const totalAlarmas = group.totalAlarmas || 0;

      const damageBadges = Array.from(group.damageSummary.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 3)
        .map(([damage, count]) => `<span class="badge badge-light text-muted mr-1 mb-1">${damage} (${count})</span>`)
        .join(' ');
      const damageContent = damageBadges || '<span class="badge badge-secondary">Sin daños reportados</span>';

      const zonas = Array.from(group.zonasAfectadas || []).sort();
      const zonasBadges = zonas.map(z => `<span class="badge badge-primary mr-1 mb-1">${z}</span>`).join(' ');

      const targetId = String(group.id || idx).replace(/'/g, "\\'");
      const borderColorClass = alarmasActivas > 0 ? 'border-warning' : 'border-secondary';

      html += `
        <div class="card mb-3 shadow-sm border-left ${borderColorClass} fms-item-card">
          <div class="card-body">
            <div class="d-flex flex-column flex-md-row justify-content-between align-items-start">
              <div class="mb-3 mb-md-0">
                <h3 class="h5 mb-2 text-dark">${tipoLabel}: ${group.elementCode}</h3>
                <div class="d-flex flex-wrap small text-muted">
                  <div class="mr-3"><strong>Zonas relacionadas:</strong> ${zonas.length}</div>
                  <div class="mr-3"><strong>Alarmas activas:</strong> ${alarmasActivas}</div>
                  <div><strong>Total alarmas:</strong> ${totalAlarmas}</div>
                </div>
              </div>
              <button class="btn btn-primary btn-sm" onclick="verDetalleFMS('${targetId}')">👁️ Ver detalle</button>
            </div>
            <div class="mt-3">
              <div class="small text-muted font-weight-bold mb-1">Daños detectados</div>
              <div class="d-flex flex-wrap">${damageContent}</div>
            </div>
            <div class="mt-3">
              <div class="small text-muted font-weight-bold mb-1">Zonas asociadas</div>
              <div class="d-flex flex-wrap">${zonasBadges}</div>
            </div>
          </div>
        </div>
      `;
    });

    html += '</div>';
    return html;
  },

  /** Construye opciones para selects de filtro */
  buildFilterOptions(fmsGroups, damageStats) {
    const tipos = new Set();
    fmsGroups.forEach(g => g.elementType && tipos.add(g.elementType));

    const tipoOptions = [
      '<option value="">Todos los tipos</option>',
      ...Array.from(tipos).sort().map(tipo => `<option value="${tipo}">${formatFMSTypeLabel(tipo)}</option>`)
    ].join('');

    const damageOptions = [
      '<option value="">Todos los daños</option>',
      ...damageStats.map(([damage, count]) => `<option value="${damage}">${damage} (${count})</option>`)
    ].join('');

    return { tipoOptions, damageOptions };
  }
};

/* ===================== FILTRO & DETALLE ===================== */
function filtrarFMS() {
  const tipoElementoSelect = document.getElementById('fmsTipoElemento');
  const damageSelect = document.getElementById('fmsDamage');
  const container = document.getElementById('fmsListContainer');

  if (!window.fmsGroupsData || !container || !tipoElementoSelect || !damageSelect) return;

  const tipoElemento = tipoElementoSelect.value || '';
  const damageFilter = damageSelect.value || '';

  let resultados = window.fmsGroupsData;

  if (tipoElemento) {
    const key = normalizeText(tipoElemento);
    resultados = resultados.filter(g => normalizeText(g.elementType) === key);
  }
  if (damageFilter) {
    const dmg = normalizeText(damageFilter);
    resultados = resultados.filter(g =>
      Array.from(g.damageSummary.keys()).some(k => normalizeText(k) === dmg)
    );
  }

  // Top defensivo de la lista filtrada
  container.innerHTML = FMSPanel.renderFMSList(resultados.slice(0, 200), { limitNotice: !tipoElemento && !damageFilter });
}

function verDetalleFMS(idxOrId) {
  const grupos = window.fmsGroupsData || [];
  let fmsItem = null;

  if (typeof idxOrId === 'number') {
    fmsItem = grupos[idxOrId];
  } else {
    const id = String(idxOrId);
    fmsItem = grupos.find(g => String(g.id) === id);
  }
  if (!fmsItem) return;

  const tipoLabel = formatFMSTypeLabel(fmsItem.elementType);
  const alarmasActivas = fmsItem.activeCount || 0;
  const totalAlarmas = fmsItem.totalAlarmas || 0;

  let html = `
    <div class="fms-detalle">
      <div class="fms-header" style="background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%); padding: 20px; border-radius: 12px; color: white; margin-bottom: 20px;">
        <h2 style="margin: 0 0 10px 0; font-size: 24px;">🚨 ${tipoLabel}: ${fmsItem.elementCode}</h2>
        <div style="display: flex; gap: 20px; flex-wrap: wrap; margin-top: 15px; font-size: 14px;">
          <div><strong>Zonas relacionadas:</strong> ${fmsItem.zonasAfectadas.size}</div>
          <div><strong>Alarmas activas:</strong> ${alarmasActivas}</div>
          <div><strong>Total alarmas:</strong> ${totalAlarmas}</div>
        </div>
      </div>
  `;

  if (fmsItem.damageSummary.size) {
    html += `
      <div class="panel-causales" style="margin-bottom: 20px;">
        <h3 style="margin: 0 0 12px 0; color: #333; font-size: 18px;">📊 Daños reportados</h3>
        <div class="table-container">
          <table style="width: 100%; font-size: 13px;">
            <thead>
              <tr><th>Daño</th><th class="number">Cantidad</th></tr>
            </thead>
            <tbody>
              ${Array.from(fmsItem.damageSummary.entries()).sort((a,b)=>b[1]-a[1]).map(([d,c])=>(
                `<tr><td>${d}</td><td class="number"><strong>${c}</strong></td></tr>`
              )).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  if (fmsItem.alarmas && fmsItem.alarmas.length) {
    html += `
      <div class="panel-alarmas">
        <h3 style="margin: 0 0 12px 0; color: #333; font-size: 18px;">📋 Alarmas (muestra)</h3>
        <div class="table-container">
          <table style="width: 100%; font-size: 13px;">
            <thead>
              <tr><th>Código</th><th>Tipo</th><th>Estado</th><th>Causal</th></tr>
            </thead>
            <tbody>
              ${fmsItem.alarmas.slice(0,50).map(a=>{
                const code = a.elementCode || a.code || a.id || '';
                const typ  = a.elementType || a.type || '';
                const est  = (a.isActive || a.activa || a.estado==='ACTIVA') ? 'Activa' : 'Histórica';
                const dmg  = formatAlarmaDamage(a);
                return `<tr><td>${code}</td><td>${typ}</td><td>${est}</td><td>${dmg}</td></tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  // Zonas asociadas
  const zonas = Array.from(fmsItem.zonasAfectadas || []).sort();
  html += `
    <div class="panel-zonas" style="margin-top: 16px;">
      <h3 style="margin: 0 0 12px 0; color: #333; font-size: 18px;">📍 Zonas asociadas</h3>
      <div class="d-flex flex-wrap">${zonas.map(z => `<span class="badge badge-primary mr-1 mb-1">${z}</span>`).join(' ')}</div>
    </div>
  </div>`;

  // Si existe el modal de alarmas, úsalo
  const body = document.getElementById('alarmaModalBody');
  const backdrop = document.getElementById('alarmaBackdrop');
  if (body && backdrop) {
    body.innerHTML = html;
    backdrop.style.display = 'block';
  } else {
    // fallback
    console.log('[FMS detalle]', fmsItem);
    alert('Detalle FMS listo en consola.');
  }
}


  html += `
      <div class="panel-zonas" style="margin-bottom: 25px;">
        <h3 style="margin: 0 0 15px 0; color: #333; font-size: 18px;">🌐 Zonas asociadas</h3>
        <div style="display: flex; gap: 8px; flex-wrap: wrap;">
  `;

  Array.from(fmsItem.zonasAfectadas).sort().forEach(zona => {
    html += `<span class="badge badge-primary">${zona}</span>`;
  });

  html += `
        </div>
      </div>

      <div class="panel-alarmas" style="margin-bottom: 25px;">
        <h3 style="margin: 0 0 15px 0; color: #333; font-size: 18px;">⚠️ Alarmas (${totalAlarmas})</h3>
        <div style="display: grid; gap: 12px;">
  `;

  const alarmasOrdenadas = Array.from(fmsItem.alarmas).sort((a, b) => Number(b.isActive) - Number(a.isActive));
  alarmasOrdenadas.forEach((alarma, idx) => {
    const estadoBadge = alarma.isActive ? '<span class="badge badge-alarma-activa">ACTIVA</span>' : '<span class="badge">CERRADA</span>';
    const damageLabel = formatAlarmaDamage(alarma);
    const incident = alarma.incidentClassification || alarma.damageClassification || '-';
    const taskName = alarma.taskName ? alarma.taskName : 'Sin tarea informada';
    const tipoElemento = formatFMSTypeLabel(alarma.elementType);

    html += `
      <div style="background: var(--bg-tertiary); border-radius: 10px; padding: 16px; border-left: 4px solid ${alarma.isActive ? '#e67e22' : 'var(--border-muted)'};">
        <div style="display:flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
          <div style="font-weight: 600; color: var(--text-primary);">Alarma ${idx + 1}</div>
          ${estadoBadge}
        </div>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; font-size: 13px;">
          <div><strong>Elemento</strong><br>${tipoElemento}</div>
          <div><strong>Código</strong><br>${alarma.elementCode || '-'}</div>
          <div><strong>Daño</strong><br>${damageLabel}</div>
          <div><strong>Clasificación</strong><br>${incident}</div>
          <div><strong>Tarea</strong><br>${taskName}</div>
          <div><strong>Tipo de alarma</strong><br>${alarma.type || '-'}</div>
          <div><strong>Creación</strong><br>${alarma.creationDate || '-'}</div>
          <div><strong>Recuperación</strong><br>${alarma.recoveryDate || 'ABIERTO'}</div>
          <div><strong>Reclamos</strong><br>${alarma.claims || '-'}</div>
          <div><strong>CM afectados</strong><br>${alarma.cmCount || '-'}</div>
        </div>
      </div>
    `;
  });

  html += `
        </div>
      </div>
    </div>
  `;

  const backdrop = document.getElementById('modalBackdrop');
  const modalTitle = document.getElementById('modalTitle');
  const modalBody = document.getElementById('modalBody');
  const modalFilters = document.getElementById('modalFilters');
  const modalFooter = document.getElementById('modalFooter');

  if (backdrop && modalTitle && modalBody) {
    modalTitle.textContent = `🚨 FMS: ${tipoLabel} - ${fmsItem.elementCode}`;
    modalBody.innerHTML = html;
    if (modalFilters) modalFilters.style.display = 'none';
    if (modalFooter) {
      modalFooter.innerHTML = `
        <div></div>
        <div style="display:flex; gap:8px;">
          <button class="btn btn-secondary" onclick="closeModal()">Cerrar</button>
        </div>`;
    }
    backdrop.classList.add('show');
    document.body.classList.add('modal-open');
  }
}
function applyEquiposFilters() {
  const select = document.getElementById('filterEquipoModelo');
  Filters.equipoModelo = Array.from(select.selectedOptions).map(opt => opt.value);
  Filters.equipoMarca = document.getElementById('filterEquipoMarca').value;
  Filters.equipoTerritorio = document.getElementById('filterEquipoTerritorio').value;
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
}

function clearEquiposFilters() {
  Filters.equipoModelo = [];
  Filters.equipoMarca = '';
  Filters.equipoTerritorio = '';
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
}

function toggleEquiposGrupo(zona){
  if (!window.equiposOpen) window.equiposOpen = new Set();
  if (window.equiposOpen.has(zona)) window.equiposOpen.delete(zona);
  else window.equiposOpen.add(zona);
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
}

/* =========================================
   HEADERS DE EXPORT
   ========================================= */
const ZONE_EXPORT_HEADERS = [
  'Fecha',
  'ZonaHFC',
  'ZonaFTTH',
  'Territorio',
  'Ubicacion',
  'Caso',
  'NumeroOrden',
  'NumeroOTuca',
  'Diagnostico',
  'Tipo',
  'TipoTrabajo',
  'Estado1',
  'Estado2',
  'Estado3',
  'MAC',
  'Modelo',
  'TipoEquipo'
];

/* =========================================
   MAPEO DE CAMPOS DE ORDEN (aliases)
   ========================================= */
const ORDER_FIELD_KEYS = {
  fecha: [
    'Fecha de creación',
    'Fecha/Hora de apertura',
    'Fecha de inicio',
    'Fecha'
  ],
  zonaHFC: [
    'Zona Tecnica HFC',
    'Zona Tecnica',
    'Zona HFC',
    'Zona'
  ],
  zonaFTTH: [
    'Zona Tecnica FTTH',
    'Zona FTTH'
  ],
  territorio: [
    'Territorio de servicio: Nombre',
    'Territorio',
    'Territorio servicio'
  ],
  ubicacionCalle: [
    'Calle',
    'Dirección',
    'Direccion',
    'Domicilio',
    'Dirección del servicio',
    'Direccion del servicio',
    'Dirección Servicio'
  ],
  ubicacionAltura: [
    'Altura',
    'Altura (N°)',
    'Altura domicilio',
    'Número',
    'Numero',
    'Nro',
    'Altura Domicilio'
  ],
  ubicacionLocalidad: [
    'Localidad',
    'Localidad Instalación',
    'Localidad Instalacion',
    'Ciudad',
    'Partido',
    'Distrito',
    'Provincia'
  ],
  caso: [
    'Número del caso',
    'Numero del caso',
    'Caso Externo',
    'Caso externo',
    'External Case Id',
    'Nro Caso',
    'Número de Caso',
    'Numero Caso'
  ],
  numeroOrden: [
    'Número de cita',
    'Numero de cita',
    'Número de Orden',
    'Numero de Orden',
    'Número Orden',
    'Numero Orden',
    'Orden',
    'Orden de trabajo',
    'N° Orden',
    'Nro Orden',
    'Número de orden de trabajo',
    'Numero de orden de trabajo',
    'N° OT'
  ],
  numeroOTuca: [
    'Número OT UCA',
    'Numero OT UCA',
    'Número de OT UCA',
    'Numero de OT UCA',
    'NumeroOTUca',
    'OT UCA',
    'OT_UCA',
    'Nro OT UCA'
  ],
  // Agrego más aliases frecuentes y una variante con paréntesis
  diagnostico: [
    'Diagnostico Tecnico',
    'Diagnóstico Técnico',
    'Diagnostico',
    'Diagnóstico',
    'Diagnostico Cliente',
    'Diagnostico tecnico',
    'Diagnóstico tecnico',
    'Diagnóstico Técnico (Detalle)',
    'Diagnostico Técnico (Detalle)'
  ],
  tipo: [
    'Tipo',
    'Tipo de caso',
    'Tipo caso',
    'Tipo OPEN',
    'Tipo FAN',
    'Tipo de OT',
    'Tipo Servicio'
  ],
  tipoTrabajo: [
    'Tipo de trabajo: Nombre de tipo de trabajo',
    'Tipo de trabajo',
    'Tipo Trabajo',
    'TipoTrabajo'
  ],
  estado1: [
    'Estado.1',
    'Estado',
    'Estado inicial'
  ],
  estado2: [
    'Estado.2',
    'Estado final',
    'Estado actual'
  ],
  estado3: [
    'Estado.3',
    'Estado final',
    'Estado detalle',
    'Estado gestion',
    'Estado gestión',
    'Estado Gestión'
  ],
  mac: [
    'MAC',
    'Mac',
    'Mac Address',
    'Dirección MAC',
    'Direccion MAC',
    'MAC Address'
  ]
};

/* =========================================
   NORMALIZACIÓN Y PICK ROBUSTO
   ========================================= */
function normalizeLabel(s){
  if (!s) return '';
  return String(s)
    .normalize('NFD')                // separa acentos
    .replace(/[\u0300-\u036f]/g, '') // quita diacríticos
    .toLowerCase()
    .replace(/\s+/g, ' ')            // colapsa espacios
    .replace(/[^\w\s:./-]/g, '')     // limpia simbolitos raros pero deja dos puntos/guiones comunes
    .trim();
}

// Construye (y cachea) un índice de headers normalizados del row
function ensureNormalizedHeaderIndex(obj){
  if (!obj || typeof obj !== 'object') return {};
  if (obj.__normIndex) return obj.__normIndex;

  const idx = {};
  for (const k of Object.keys(obj)){
    const nk = normalizeLabel(k);
    if (!idx[nk]) idx[nk] = k; // guarda key original para recuperar el valor exacto
  }
  Object.defineProperty(obj, '__normIndex', {
    value: idx,
    enumerable: false,
    configurable: false
  });
  return idx;
}

function pickFirstValue(obj, keys){
  if (!obj || !keys) return '';
  // 1) Intento exacto primero
  for (const key of keys){
    if (!key) continue;
    const value = obj[key];
    if (value !== null && value !== undefined){
      const str = String(value).trim();
      if (str) return str;
    }
  }
  // 2) Intento por equivalencia normalizada (ignora acentos/case/espacios)
  const normIdx = ensureNormalizedHeaderIndex(obj);
  for (const key of keys){
    if (!key) continue;
    const nk = normalizeLabel(key);
    const original = normIdx[nk];
    if (original !== undefined){
      const value = obj[original];
      if (value !== null && value !== undefined){
        const str = String(value).trim();
        if (str) return str;
      }
    }
  }
  // 3) Fallback específico para "Diagnostico" por si viene con sufijos tipo "(Descripción)"
  //    Busca la PRIMERA columna cuyo normalizado empiece con 'diagnostico'
  const keysNorm = Object.keys(normIdx);
  if (keys === ORDER_FIELD_KEYS.diagnostico){
    const kFound = keysNorm.find(k => k.startsWith('diagnostico'));
    if (kFound){
      const original = normIdx[kFound];
      const v = obj[original];
      if (v !== null && v !== undefined){
        const str = String(v).trim();
        if (str) return str;
      }
    }
  }
  return '';
}

/* =========================================
   UBICACIÓN Y ROW BUILDER
   ========================================= */
function buildUbicacion(order){
  const calle = pickFirstValue(order, ORDER_FIELD_KEYS.ubicacionCalle);
  const altura = pickFirstValue(order, ORDER_FIELD_KEYS.ubicacionAltura);
  const localidad = pickFirstValue(order, ORDER_FIELD_KEYS.ubicacionLocalidad);

  const parts = [];
  if (calle && altura){
    parts.push(`${calle} ${altura}`.trim());
  } else if (calle){
    parts.push(calle);
  } else if (altura){
    parts.push(altura);
  }

  if (localidad){
    parts.push(localidad);
  }

  return parts.join(', ');
}

function buildOrderExportRow(order, zoneInfo){
  if (!order) return null;
  const meta = order.__meta || {};

  let fecha = pickFirstValue(order, ORDER_FIELD_KEYS.fecha);
  if (!fecha && meta.fecha && typeof DateUtils !== 'undefined' && DateUtils.format){
    try { fecha = DateUtils.format(meta.fecha); } catch(e){ /* noop */ }
  }

  const zonaHFC = meta.zonaHFC || zoneInfo?.zonaHFC || pickFirstValue(order, ORDER_FIELD_KEYS.zonaHFC) || zoneInfo?.zona || '';
  const zonaFTTH = meta.zonaFTTH || zoneInfo?.zonaFTTH || pickFirstValue(order, ORDER_FIELD_KEYS.zonaFTTH) || '';
  const territorio = meta.territorio || pickFirstValue(order, ORDER_FIELD_KEYS.territorio) || '';
  const ubicacion = buildUbicacion(order);
  const caso = meta.numeroCaso || pickFirstValue(order, ORDER_FIELD_KEYS.caso) || '';
  const numeroOrden = pickFirstValue(order, ORDER_FIELD_KEYS.numeroOrden) || '';
  const numeroOTuca = pickFirstValue(order, ORDER_FIELD_KEYS.numeroOTuca) || '';

  // Clave: robustecer este campo
  const diagnostico = pickFirstValue(order, ORDER_FIELD_KEYS.diagnostico) || '';

  const tipo = pickFirstValue(order, ORDER_FIELD_KEYS.tipo) || zoneInfo?.tipo || '';
  const tipoTrabajo = pickFirstValue(order, ORDER_FIELD_KEYS.tipoTrabajo) || '';

  const estado1 = pickFirstValue(order, ORDER_FIELD_KEYS.estado1) || '';
  let estado2 = pickFirstValue(order, ORDER_FIELD_KEYS.estado2) || '';
  let estado3 = pickFirstValue(order, ORDER_FIELD_KEYS.estado3) || '';

  if (estado3 && estado2 && estado3 === estado2){
    const alternativas = ['Estado final', 'Estado gestión', 'Estado Gestion', 'Estado detalle'];
    const altern = pickFirstValue(order, alternativas);
    if (altern && altern !== estado2){
      estado3 = altern;
    }
  }

  let mac = pickFirstValue(order, ORDER_FIELD_KEYS.mac) || '';
  let modelo = '';
  let tipoEquipo = '';

  // Intentar extraer dispositivos si no están en meta
  if (!Array.isArray(meta.dispositivos) || !meta.dispositivos.length) {
    if (typeof findDispositivosColumn === 'function'){
      const colInfo = findDispositivosColumn(order);
      if (colInfo && order[colInfo] && typeof TextUtils !== 'undefined' && TextUtils.parseDispositivosJSON){
        try {
          const dispositivos = TextUtils.parseDispositivosJSON(order[colInfo]);
          if (dispositivos && dispositivos.length) meta.dispositivos = dispositivos;
        } catch(e){ /* noop */ }
      }
    }
  }

  // Extraer MAC, Modelo y Tipo de equipo del primer dispositivo
  if (Array.isArray(meta.dispositivos) && meta.dispositivos.length){
    const device = meta.dispositivos[0] || {};
    if (!mac) mac = String(device.macAddress || device.mac || '').trim();
    modelo = String(device.model || device.modelo || '').trim();
    tipoEquipo = String(device.type || device.tipo || '').trim();
  }

  return {
    Fecha: fecha || '',
    ZonaHFC: zonaHFC || '',
    ZonaFTTH: zonaFTTH || '',
    Territorio: territorio || '',
    Ubicacion: ubicacion || '',
    Caso: caso || '',
    NumeroOrden: numeroOrden || '',
    NumeroOTuca: numeroOTuca || '',
    Diagnostico: diagnostico || '',
    Tipo: tipo || '',
    TipoTrabajo: tipoTrabajo || '',
    Estado1: estado1 || '',
    Estado2: estado2 || '',
    Estado3: estado3 || '',
    MAC: mac || '',
    Modelo: modelo || '',
    TipoEquipo: tipoEquipo || ''
  };
}


/* =================== EXPORT: helpers y builders =================== */
function buildZoneExportRows(zoneData){
  if (!zoneData) return [];
  const source = zoneData.ordenesOriginales || zoneData.ordenes || [];
  return source
    .map(order => buildOrderExportRow(order, zoneData)) // se asume definida globalmente
    // Evita dependencia de ZONE_EXPORT_HEADERS para el filtro: valida que haya al menos un valor no vacío
    .filter(row => {
      if (!row || typeof row !== 'object') return false;
      return Object.values(row).some(v => String(v ?? '').trim().length > 0);
    });
}

function createWorksheetFromRows(rows, headers){
  if (!rows || !rows.length) return null;
  // Si no hay headers definidos, los infiere del primer row
  const finalHeaders = (Array.isArray(headers) && headers.length)
    ? headers
    : Object.keys(rows[0]);
  const data = rows.map(row => finalHeaders.map(h => (row[h] ?? '')));
  return XLSX.utils.aoa_to_sheet([finalHeaders, ...data]);
}

function sanitizeSheetName(name){
  const fallback = 'Hoja';
  if (!name) return fallback;
  // Remueve caracteres inválidos para nombres de hoja
  const invalidChars = /[\/\?\*\:\[\]]/g; // escapadas correctas dentro del charclass
  const cleaned = name.toString()
    .replace(invalidChars, ' ')
    .replace(/[\u0000-\u001f]/g, ' ')
    .trim();
  const truncated = cleaned.substring(0, 31);
  return truncated || fallback;
}

function appendSheet(workbook, worksheet, desiredName, usedNames){
  if (!worksheet || !workbook) return false;
  const base = sanitizeSheetName(desiredName);
  let name = base || 'Hoja';
  let counter = 1;
  while (usedNames.has(name)){
    counter += 1;
    const suffix = `_${counter}`;
    const baseTrim = (base || 'Hoja').substring(0, Math.max(31 - suffix.length, 1));
    name = `${baseTrim}${suffix}`;
  }
  usedNames.add(name);
  XLSX.utils.book_append_sheet(workbook, worksheet, name);
  return true;
}

/* =================== EXPORT: acciones =================== */
function exportEquiposGrupoExcel(zona, useFiltered = false){
  const source = (useFiltered && window.equiposPorZona)
    ? window.equiposPorZona
    : window.equiposPorZonaCompleto;

  if (!source) return toast && toast('No hay datos de equipos');

  const arr = source.get ? (source.get(zona) || []) : [];
  if (!arr.length) return toast && toast('No hay equipos en esa zona');

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(arr);
  XLSX.utils.book_append_sheet(wb, ws, `Equipos_${(zona || 'NA').toString().slice(0, 25)}`);

  const FiltersSafe = (typeof Filters === 'object' && Filters) ? Filters : {};
  const filterInfo = (useFiltered && (
      (Array.isArray(FiltersSafe.equipoModelo) && FiltersSafe.equipoModelo.length > 0) ||
      FiltersSafe.equipoMarca ||
      FiltersSafe.equipoTerritorio
    )) ? `_filtrado` : '';

  XLSX.writeFile(
    wb,
    `Equipos_${zona || 'NA'}${filterInfo}_${new Date().toISOString().slice(0, 10)}.xlsx`
  );
  toast && toast(`✅ Exportados ${arr.length} equipos de ${zona}`);
}

function exportZonaExcel(zoneIdx) {
  const zonaData = window.currentAnalyzedZones?.[zoneIdx];
  if (!zonaData) return toast && toast('No hay datos de la zona');

  const rows = buildZoneExportRows(zonaData);
  if (!rows.length) {
    return toast && toast('No hay órdenes para exportar en esta zona');
  }

  const wb = XLSX.utils.book_new();
  const usedNames = new Set();

  // Usa ZONE_EXPORT_HEADERS si existe; si no, createWorksheetFromRows infiere headers
  const headers = (typeof ZONE_EXPORT_HEADERS !== 'undefined' && Array.isArray(ZONE_EXPORT_HEADERS) && ZONE_EXPORT_HEADERS.length)
    ? ZONE_EXPORT_HEADERS
    : undefined;

  const sheet = createWorksheetFromRows(rows, headers);
  appendSheet(wb, sheet, `Zona_${zonaData.zona}`, usedNames);

  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Zona_${zonaData.zona}_${fecha}.xlsx`);
  toast && toast(`✅ Exportada zona ${zonaData.zona} (${rows.length} órdenes)`);
}

function exportCMTSExcel(cmts) {
  const cmtsData = (window.currentCMTSData || []).find(c => c.cmts === cmts);
  if (!cmtsData) return toast && toast('No hay datos del CMTS');

  const wb = XLSX.utils.book_new();
  const zonasFlat = (cmtsData.zonas || []).map(z => ({
    Zona: z.zona,
    Tipo: z.tipo,
    Total_OTs: z.totalOTs,
    Ingreso_N: z.ingresoN,
    Ingreso_N1: z.ingresoN1,
    Estado_Nodo: z.nodoEstado
  }));

  const ws = XLSX.utils.json_to_sheet(zonasFlat);
  XLSX.utils.book_append_sheet(wb, ws, `CMTS_${String(cmts).slice(0, 20)}`);

  XLSX.writeFile(wb, `CMTS_${String(cmts).slice(0, 20)}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  toast && toast(`✅ Exportado CMTS ${cmts}`);
}

/* =================== STATE GLOBAL =================== */
let currentData = null;
let allZones = [];
let allCMTS = [];
let currentZone = null;
let selectedOrders = new Set();
window.equiposOpen = new Set();

// (¡Corregido!) Definición ÚNICA de fmsRenderState
const fmsRenderState = {
  needsRender: true,
  isRendering: false,
  lastOrders: [],
  lastFmsMap: new Map()
};

/* =================== INIT & LISTENERS =================== */
document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
});

function setupEventListeners() {
  const el1 = document.getElementById('fileConsolidado1');
  const el2 = document.getElementById('fileConsolidado2');
  const el3 = document.getElementById('fileNodos');
  const el4 = document.getElementById('fileFMS');

  el1 && el1.addEventListener('change', e => loadFile(e, 1));
  el2 && el2.addEventListener('change', e => loadFile(e, 2));
  el3 && el3.addEventListener('change', e => loadFile(e, 3));
  el4 && el4.addEventListener('change', e => loadFile(e, 4));

  const filterCatec = document.getElementById('filterCATEC');
  const filterExcludeCatec = document.getElementById('filterExcludeCATEC');

  if (filterCatec && filterExcludeCatec) {
    filterCatec.addEventListener('change', e => {
      if (e.target.checked) filterExcludeCatec.checked = false;
      applyFilters && applyFilters();
    });

    filterExcludeCatec.addEventListener('change', e => {
      if (e.target.checked) filterCatec.checked = false;
      applyFilters && applyFilters();
    });
  }

  const showAllStates = document.getElementById('showAllStates');
  showAllStates && showAllStates.addEventListener('change', () => applyFilters && applyFilters());

  const filterFTTH = document.getElementById('filterFTTH');
  const filterExcludeFTTH = document.getElementById('filterExcludeFTTH');
  filterFTTH && filterFTTH.addEventListener('change', e => {
    if (e.target.checked && filterExcludeFTTH) filterExcludeFTTH.checked = false;
    applyFilters && applyFilters();
  });
  filterExcludeFTTH && filterExcludeFTTH.addEventListener('change', e => {
    if (e.target.checked && filterFTTH) filterFTTH.checked = false;
    applyFilters && applyFilters();
  });

  const binding = [
    'filterNodoEstado','filterCMTS','daysWindow','filterTerritorio',
    'filterSistema','filterAlarma','ordenarPorIngreso'
  ];
  binding.forEach(id => {
    const el = document.getElementById(id);
    el && el.addEventListener('change', () => applyFilters && applyFilters());
  });

  const quickSearch = document.getElementById('quickSearch');
  if (quickSearch) {
    const deb = (typeof debounce === 'function') ? debounce : (fn => fn);
    quickSearch.addEventListener('input', deb(() => applyFilters && applyFilters(), 300));
  }

  // Multiselect de zonas
  const zoneFilterSearch = document.getElementById('zoneFilterSearch');
  zoneFilterSearch && zoneFilterSearch.addEventListener('input', onZoneFilterSearch);

  const zoneFilterOptions = document.getElementById('zoneFilterOptions');
  zoneFilterOptions && zoneFilterOptions.addEventListener('change', onZoneOptionChange);

  document.addEventListener('click', handleZoneFilterOutsideClick);
}

/* =================== LOAD FILES =================== */
async function loadFile(e, tipo) {
  const file = e.target.files && e.target.files[0];
  if (!file) return;

  const statusEl = document.getElementById(`status${tipo}`);
  if (statusEl) {
    statusEl.textContent = 'Cargando...';
    statusEl.classList.remove('loaded');
  }

  let result = { success: false, rows: 0, error: 'Sin procesar' };
  try {
    if (tipo === 4 && /\.csv$/i.test(file.name)) {
      result = await dataProcessor.loadCSV(file);
    } else {
      result = await dataProcessor.loadExcel(file, tipo);
    }
  } catch (err) {
    result = { success: false, rows: 0, error: (err && err.message) || String(err) };
  }

  if (result.success) {
    statusEl && (statusEl.textContent = `✓ ${result.rows} filas cargadas`, statusEl.classList.add('loaded'));
    const nombres = ['Consolidado 1', 'Consolidado 2', 'Nodos UP/DOWN', 'Alarmas FMS'];
    toast && toast(`${nombres[tipo - 1]} cargado: ${result.rows} registros`);

    // Si al menos uno de los consolidado está, procesamos para habilitar vistas
    if (dataProcessor.consolidado1 || dataProcessor.consolidado2) {
      const ms = document.getElementById('mergeStatus');
      ms && (ms.style.display = 'flex', (document.getElementById('mergeStatusText') || {}).textContent = 'Procesando...');
      // processData debe existir globalmente
      typeof processData === 'function' && processData();
    }
  } else {
    statusEl && (statusEl.textContent = `✗ Error`);
    toast && toast(`Error al cargar archivo: ${result.error}`);
  }
}

function processData() {
  const merged = dataProcessor.merge();
  if (!merged.length) return;
  
  const daysWindow = parseInt(document.getElementById('daysWindow').value);
  
  const zones = dataProcessor.processZones(merged);
  allZones = dataProcessor.analyzeZones(zones, daysWindow);
  allCMTS = dataProcessor.analyzeCMTS(allZones);
  
  const territoriosAnalisis = dataProcessor.analyzeTerritorios(allZones);
  window.territoriosData = territoriosAnalisis;
  
  currentData = {
    ordenes: merged,
    zonas: allZones
  };
  
  populateFilters();
  
  const stats = {
    total: allZones.reduce((sum, z) => sum + z.totalOTs, 0),
    zonas: allZones.length,
    territoriosCriticos: territoriosAnalisis.filter(t => t.esCritico).length,
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

function populateFilters() {
  const territorios = [...new Set(
    currentData.ordenes
      .map(o => o['Territorio de servicio: Nombre'])
      .filter(Boolean)
  )];
  const terrSelect = document.getElementById('filterTerritorio');
  terrSelect.innerHTML = '<option value="">Todos</option>' + territorios.sort().map(t => `<option value="${t}">${t}</option>`).join('');

  const cmtsList = [...new Set(
    allZones
      .map(z => z.cmts)
      .filter(Boolean)
  )];
  const cmtsSelect = document.getElementById('filterCMTS');
  cmtsSelect.innerHTML = '<option value="">Todos</option>' + cmtsList.sort().map(c => `<option value="${c}">${c}</option>`).join('');

  initializeZoneFilterOptions(currentData.zonas || [], {
    resetSelection: true,
    resetSearch: true
  });
}

function collectFiltersFromUI() {
  const getCheckbox = id => {
    const el = document.getElementById(id);
    return el ? Boolean(el.checked) : false;
  };

  const getValue = id => {
    const el = document.getElementById(id);
    if (!el) return '';
    return el.value ?? '';
  };

  Filters.catec = getCheckbox('filterCATEC');
  Filters.excludeCatec = !Filters.catec && getCheckbox('filterExcludeCATEC');
  if (Filters.catec) {
    const excludeCatecEl = document.getElementById('filterExcludeCATEC');
    if (excludeCatecEl && excludeCatecEl.checked) excludeCatecEl.checked = false;
  } else if (Filters.excludeCatec) {
    const catecEl = document.getElementById('filterCATEC');
    if (catecEl && catecEl.checked) catecEl.checked = false;
  }
  Filters.showAllStates = getCheckbox('showAllStates');
  Filters.ftth = getCheckbox('filterFTTH');
  Filters.excludeFTTH = !Filters.ftth && getCheckbox('filterExcludeFTTH');
  if (Filters.ftth) {
    const excludeFtthEl = document.getElementById('filterExcludeFTTH');
    if (excludeFtthEl && excludeFtthEl.checked) excludeFtthEl.checked = false;
  } else if (Filters.excludeFTTH) {
    const ftthEl = document.getElementById('filterFTTH');
    if (ftthEl && ftthEl.checked) ftthEl.checked = false;
  }
  Filters.nodoEstado = getValue('filterNodoEstado');
  Filters.cmts = getValue('filterCMTS');
  Filters.territorio = getValue('filterTerritorio');
  Filters.sistema = getValue('filterSistema');
  Filters.alarma = getValue('filterAlarma');
  Filters.quickSearch = (getValue('quickSearch') || '').trim();
  Filters.ordenarPorIngreso = getValue('ordenarPorIngreso') || 'desc';

  const daysInput = document.getElementById('daysWindow');
  const parsedDays = parseInt(daysInput?.value, 10);
  if (Number.isFinite(parsedDays) && parsedDays > 0) {
    Filters.days = parsedDays;
  } else {
    Filters.days = CONFIG.defaultDays;
    if (daysInput) daysInput.value = CONFIG.defaultDays;
  }

  Filters.selectedZonas = Array.from(zoneFilterState.selected);
}

function runFiltersPipeline() {
  const sourceOrders = Array.isArray(currentData?.ordenes) ? currentData.ordenes : [];
  const filtered = Filters.apply(sourceOrders);

  const zones = dataProcessor.processZones(filtered);
  let analyzed = dataProcessor.analyzeZones(zones, Filters.days);
  analyzed = Filters.applyToZones(analyzed);

  const cmtsFiltered = dataProcessor.analyzeCMTS(analyzed);
  const territoriosAnalisis = dataProcessor.analyzeTerritorios(analyzed);

  return { filtered, analyzed, cmtsFiltered, territoriosAnalisis };
}

function renderFMSMessage(message, icon = 'ℹ️') {
  return `<div class="loading-message"><p>${icon} ${message}</p></div>`;
}

function updateFMSPendingData(orders, fmsMap) {
  fmsRenderState.lastOrders = Array.isArray(orders) ? orders : [];
  fmsRenderState.lastFmsMap = fmsMap instanceof Map ? fmsMap : new Map();
  fmsRenderState.needsRender = true;

  const fmsPanelEl = document.getElementById('fmsPanel');
  if (fmsPanelEl && !fmsPanelEl.classList.contains('active')) {
    fmsPanelEl.innerHTML = renderFMSMessage('El panel FMS se generará al abrir la pestaña.');
  } else {
    renderFMSPanelIfNeeded();
  }
}

function renderFMSPanelIfNeeded(options = {}) {
  const fmsPanelEl = document.getElementById('fmsPanel');
  if (!fmsPanelEl || fmsRenderState.isRendering) return;

  const isActive = fmsPanelEl.classList.contains('active');
  const shouldRender = options.force || (isActive && fmsRenderState.needsRender);

  if (!shouldRender) return;

  const fmsMap = fmsRenderState.lastFmsMap || new Map();
  if (!fmsMap.size) {
    fmsPanelEl.innerHTML = renderFMSMessage('No hay datos de FMS/Alarmas cargados', '⚠️');
    fmsRenderState.needsRender = false;
    return;
  }

  fmsRenderState.isRendering = true;
  fmsPanelEl.innerHTML = renderFMSMessage('Procesando panel FMS...', '⏳');
  requestAnimationFrame(() => {
    try {
      fmsPanelEl.innerHTML = FMSPanel.render(fmsRenderState.lastOrders || [], fmsMap);
    } catch (err) {
      console.error('Error rendering FMS panel', err);
      fmsPanelEl.innerHTML = renderFMSMessage('No se pudo renderizar el panel FMS.', '❌');
    } finally {
      fmsRenderState.needsRender = false;
      fmsRenderState.isRendering = false;
    }
  });
}

function applyFilters() {
  if (!currentData) return;

  collectFiltersFromUI();

  let pipelineResult = null;
  let trimmed = false;
  let attempts = 0;

  do {
    Filters.selectedZonas = Array.from(zoneFilterState.selected);
    pipelineResult = runFiltersPipeline();

    trimmed = initializeZoneFilterOptions(pipelineResult.analyzed, {
      resetSelection: false,
      resetSearch: false,
      forceRender: attempts === 0
    });

    attempts += 1;
  } while (trimmed && attempts < 5);

  if (trimmed) {
    console.warn('Zone filter selection required repeated trimming; latest results may exclude inconsistent zonas.');
  }

  if (!pipelineResult) return;

  Filters.selectedZonas = Array.from(zoneFilterState.selected);

  const { filtered, analyzed, cmtsFiltered, territoriosAnalisis } = pipelineResult;
  window.territoriosData = territoriosAnalisis;

  const stats = {
    total: analyzed.reduce((sum, z) => sum + z.totalOTs, 0),
    zonas: analyzed.length,
    territoriosCriticos: territoriosAnalisis.filter(t => t.esCritico).length,
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

  document.getElementById('btnExportExcel').disabled = analyzed.length === 0 && filtered.length === 0;
  document.getElementById('btnExportExcelZonas').disabled = analyzed.length === 0;

  document.getElementById('zonasPanel').innerHTML = UIRenderer.renderZonas(analyzed);
  document.getElementById('cmtsPanel').innerHTML = UIRenderer.renderCMTS(cmtsFiltered);
  document.getElementById('edificiosPanel').innerHTML = UIRenderer.renderEdificios(filtered);
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(filtered);
  const fmsMap = dataProcessor ? dataProcessor.fmsMap : new Map();
  updateFMSPendingData(filtered, fmsMap);
}

function resetFiltersState() {
  const catecEl = document.getElementById('filterCATEC');
  const excludeCatecEl = document.getElementById('filterExcludeCATEC');

  if (catecEl) catecEl.checked = false;
  if (excludeCatecEl) excludeCatecEl.checked = false;
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
  Filters.criticidad = '';
  resetZoneFilterState();
  closeZoneFilter();
}

function clearFilters() {
  resetFiltersState();
  applyFilters();
}

function filterByStat(statType) {
  resetFiltersState();

  switch (statType) {
    case 'total':
    case 'zonas':
      toast('🔄 Mostrando todas las zonas (filtros reseteados)');
      break;
    
    case 'territoriosCriticos':
      if (!window.territoriosData) {
        toast('❌ No hay datos de territorios disponibles');
        break;
      }
      
      const territoriosCriticos = window.territoriosData.filter(t => t.esCritico);
      if (territoriosCriticos.length === 0) {
        toast('✅ No hay territorios críticos en este momento');
        break;
      }
      
      showTerritoriosCriticosModal(territoriosCriticos);
      return;
    
    case 'ftth':
      document.getElementById('filterFTTH').checked = true;
      document.getElementById('filterExcludeFTTH').checked = false;
      toast('🔌 Mostrando solo zonas FTTH');
      break;
    
    case 'alarmas':
      document.getElementById('filterAlarma').value = 'con-alarma';
      toast('🚨 Mostrando zonas con alarmas activas');
      break;
    
    case 'nodosCriticos':
      document.getElementById('filterNodoEstado').value = 'critical';
      toast('🟥 Mostrando zonas con NODOS CRÍTICOS');
      break;
  }

  applyFilters();
}

function showTerritoriosCriticosModal(territoriosCriticos) {
  let html = '<div class="territorios-criticos-container">';
  html += '<h3 style="color: #D13438; margin-bottom: 20px;">🚨 Territorios Críticos Detectados</h3>';
  
  territoriosCriticos.forEach((t, idx) => {
    html += `<div class="territorio-critico-card" style="
      background: white;
      border: 2px solid #D13438;
      border-radius: 8px;
      padding: 16px;
      margin-bottom: 16px;
      box-shadow: 0 4px 8px rgba(209, 52, 56, 0.15);
    ">`;
    
    html += `<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">`;
    html += `<h4 style="color: #0078D4; margin: 0; font-size: 1.1rem;">${t.territorio}</h4>`;
    html += `<span class="badge badge-critical" style="font-size: 0.85rem;">CRÍTICO</span>`;
    html += `</div>`;
    
    html += `<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin-bottom: 12px;">`;
    html += `<div><strong>Total OTs:</strong> ${t.totalOTs}</div>`;
    html += `<div><strong>Zonas Críticas:</strong> ${t.zonasCriticas}</div>`;
    html += `<div><strong>Con Alarmas:</strong> ${t.zonasConAlarma}</div>`;
    html += `<div><strong>Total Zonas:</strong> ${t.zonas.length}</div>`;
    html += `</div>`;
    
    if (t.flexSubdivisiones.size > 0) {
      html += `<div style="margin-top: 12px; padding-top: 12px; border-top: 1px solid #E1DFDD;">`;
      html += `<strong style="color: #605E5C; font-size: 0.9rem;">📍 Subdivisiones:</strong>`;
      html += `<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)); gap: 8px; margin-top: 8px;">`;
      
      Array.from(t.flexSubdivisiones.values()).forEach(flex => {
        const badgeColor = flex.zonasCriticas > 0 ? '#D13438' : '#107C10';
        html += `<div style="
          background: ${badgeColor}15;
          border: 1px solid ${badgeColor};
          border-radius: 6px;
          padding: 8px;
          text-align: center;
        ">`;
        html += `<div style="font-weight: 600; color: ${badgeColor};">${flex.nombre}</div>`;
        html += `<div style="font-size: 0.75rem; color: #605E5C;">${flex.zonas.length} zonas</div>`;
        html += `<div style="font-size: 0.75rem; color: #605E5C;">${flex.totalOTs} OTs</div>`;
        if (flex.zonasCriticas > 0) {
          html += `<div style="font-size: 0.75rem; color: ${badgeColor};">⚠️ ${flex.zonasCriticas} críticas</div>`;
        }
        html += `</div>`;
      });
      
      html += `</div></div>`;
    }
    
    html += `<div style="margin-top: 12px;">`;
    html += `<button class="btn btn-primary" style="font-size: 0.85rem; padding: 8px 16px;" 
              onclick="filtrarPorTerritorio('${t.territorio.replace(/'/g, "\\'")}')">
      🔍 Ver Zonas de ${t.territorio}
    </button>`;
    html += `</div>`;
    
    html += `</div>`;
  });
  
  html += `<div style="text-align: center; margin-top: 20px;">`;
  html += `<button class="btn btn-secondary" onclick="closeTerritoriosCriticosModal()">Cerrar</button>`;
  html += `</div>`;
  
  html += '</div>';
  
  document.getElementById('modalTitle').textContent = `📊 Territorios Críticos (${territoriosCriticos.length})`;
  document.getElementById('modalBody').innerHTML = html;
  document.getElementById('modalFilters').style.display = 'none';
  document.getElementById('modalFooter').innerHTML = '';
  
  document.getElementById('modalBackdrop').classList.add('show');
  document.body.classList.add('modal-open');
}

function closeTerritoriosCriticosModal() {
  closeModal();
}

function filtrarPorTerritorio(territorioNormalizado) {
  closeModal();
  
  const territoriosOriginales = [...new Set(
    currentData.ordenes
      .map(o => o['Territorio de servicio: Nombre'])
      .filter(t => t && TerritorioUtils.normalizar(t) === territorioNormalizado)
  )];
  
  if (territoriosOriginales.length === 0) {
    toast('❌ No se encontraron zonas para este territorio');
    return;
  }
  
  document.getElementById('filterTerritorio').value = territoriosOriginales[0];
  toast(`🔍 Filtrando por territorio: ${territorioNormalizado}`);
  applyFilters();
  
  switchTab('zonas');
}

function switchTab(tabName) {
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(panel => panel.classList.remove('active'));

  document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
  document.getElementById(`${tabName}Panel`).classList.add('active');

  if (tabName === 'fms') {
    renderFMSPanelIfNeeded({ force: true });
  }
}

function showAlarmaInfo(zoneIdx) {
  const zonaData = window.currentAnalyzedZones[zoneIdx];
  if (!zonaData || !zonaData.alarmas.length) return;

  let html = '<div class="alarma-info-box">';
  html += `<h4>🚨 Alarmas en Zona: ${zonaData.zona}</h4>`;

  zonaData.alarmas.forEach((a, idx) => {
    const tipoLabel = formatFMSTypeLabel(a.elementType);
    const damageLabel = formatAlarmaDamage(a);
    const incident = a.incidentClassification || a.damageClassification || '-';
    const taskName = a.taskName ? a.taskName : 'Sin tarea informada';

    html += `<div style="margin-bottom: 20px; padding-bottom: 20px; border-bottom: 1px solid var(--border-subtle);">`;
    html += `<h5 style="color: var(--win-blue); margin-bottom: 10px;">Alarma ${idx + 1} ${a.isActive ? '<span class="badge badge-alarma-activa">ACTIVO</span>' : '<span class="badge">CERRADO</span>'}</h5>`;
    html += `<div class="alarma-info-grid">`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Tipo Alarma</span><span class="alarma-info-value">${a.type || '-'}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Elemento</span><span class="alarma-info-value">${tipoLabel}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Código Elemento</span><span class="alarma-info-value">${a.elementCode || '-'}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Fecha Creación</span><span class="alarma-info-value">${a.creationDate || '-'}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Fecha Cierre</span><span class="alarma-info-value">${a.recoveryDate || 'ABIERTO'}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Daño</span><span class="alarma-info-value">${damageLabel}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Clasificación</span><span class="alarma-info-value">${incident}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Tarea</span><span class="alarma-info-value">${taskName}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">Reclamos</span><span class="alarma-info-value">${a.claims || '-'}</span></div>`;
    html += `<div class="alarma-info-item"><span class="alarma-info-label">CM afectados</span><span class="alarma-info-value">${a.cmCount || '-'}</span></div>`;
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

  const alarmasActivas = edificio.alarmasActivas?.length || 0;
  const totalAlarmas = edificio.alarmas?.length || 0;

  let html = `<div style="margin-bottom: 16px; padding: 12px; background: var(--bg-tertiary); border-radius: 8px;">
    <strong>Zona:</strong> ${edificio.zona}<br>
    <strong>Territorio:</strong> ${edificio.territorio}<br>
    <strong>Total OTs:</strong> ${edificio.casos.length}<br>
    <strong>Alarmas activas:</strong> ${alarmasActivas}${totalAlarmas && totalAlarmas !== alarmasActivas ? ` (de ${totalAlarmas})` : ''}
  </div>`;

  if (totalAlarmas > 0) {
    const alarmasOrdenadas = Array.from(edificio.alarmas).sort((a, b) => Number(b.isActive) - Number(a.isActive));
    html += `<div class="panel-alarmas" style="margin-bottom: 20px;">
      <h3 style="margin: 0 0 12px 0; color: var(--text-primary); font-size: 16px;">🚨 Alarmas vinculadas</h3>
      <div style="display: grid; gap: 12px;">
    `;

    alarmasOrdenadas.forEach((alarma, idxAlarma) => {
      const estadoBadge = alarma.isActive
        ? '<span class="badge badge-alarma-activa">ACTIVA</span>'
        : '<span class="badge">CERRADA</span>';
      const damageLabel = formatAlarmaDamage(alarma);
      const taskName = alarma.taskName ? alarma.taskName : 'Sin tarea informada';
      const tipoLabel = formatFMSTypeLabel(alarma.elementType);

      html += `<div style="background: var(--bg-tertiary); border-radius: 10px; padding: 14px; border-left: 4px solid ${alarma.isActive ? 'var(--accent-warning)' : 'var(--border-muted)'};">
        <div style="display:flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
          <div style="font-weight: 600; color: var(--text-primary);">Alarma ${idxAlarma + 1}</div>
          ${estadoBadge}
        </div>
        <div class="alarma-info-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 8px; font-size: 13px;">
          <div><span class="alarma-info-label">Elemento</span><br><span class="alarma-info-value">${tipoLabel}</span></div>
          <div><span class="alarma-info-label">Código</span><br><span class="alarma-info-value">${alarma.elementCode || '-'}</span></div>
          <div><span class="alarma-info-label">Daño</span><br><span class="alarma-info-value">${damageLabel}</span></div>
          <div><span class="alarma-info-label">Tarea</span><br><span class="alarma-info-value">${taskName}</span></div>
          <div><span class="alarma-info-label">Tipo</span><br><span class="alarma-info-value">${alarma.type || '-'}</span></div>
          <div><span class="alarma-info-label">Creación</span><br><span class="alarma-info-value">${alarma.creationDate || '-'}</span></div>
          <div><span class="alarma-info-label">Recuperación</span><br><span class="alarma-info-value">${alarma.recoveryDate || 'ABIERTO'}</span></div>
        </div>
      </div>`;
    });

    html += '</div></div>';
  }

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

function exportEdificioDetalle() {
  const titleEl = document.getElementById('edificioModalTitle');
  const title = titleEl ? titleEl.textContent || '' : '';
  if (!title || !title.startsWith('🏢 Edificio: ')) {
    toast('Seleccioná un edificio para exportar');
    return;
  }

  const direccion = title.replace('🏢 Edificio: ', '').trim();
  const edificios = Array.isArray(window.edificiosData) ? window.edificiosData : [];
  const edificio = edificios.find(e => (e.direccion || '').trim() === direccion);

  if (!edificio) {
    toast('No hay datos del edificio');
    return;
  }

  const casos = Array.isArray(edificio.casos) ? edificio.casos : [];
  if (!casos.length) {
    toast('El edificio no tiene órdenes para exportar');
    return;
  }

  const wb = XLSX.utils.book_new();
  const usedNames = new Set();
  appendSheet(wb, XLSX.utils.json_to_sheet(casos), 'Edificio_OTs', usedNames);

  const safeSlug = (direccion.normalize ? direccion.normalize('NFD') : direccion)
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 40) || 'Edificio';

  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Edificio_${safeSlug}_${fecha}.xlsx`);
  toast('✓ Edificio exportado');
}

function openModal(zoneIdx) {
  const zonaData = window.currentAnalyzedZones[zoneIdx];
  if (!zonaData) return;
  
  currentZone = zonaData;
  selectedOrders.clear();
  
  document.getElementById('modalTitle').textContent = `Detalle: ${currentZone.zona}`;
  
  const ordenesParaModal = currentZone.ordenesOriginales || currentZone.ordenes;
  currentZone.ordenes = ordenesParaModal;
  
  const dias = [...new Set(
    ordenesParaModal
      .map(o => {
        const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || o['Fecha de inicio'];
        const dt = DateUtils.parse(fecha);
        return dt ? DateUtils.format(dt) : null;
      })
      .filter(Boolean)
  )].sort();
  
  const diaSelect = document.getElementById('filterDia');
  diaSelect.innerHTML = '<option value="">Todos los días</option>' + dias.map(d => `<option value="${d}">${d}</option>`).join('');
  
  renderModalContent();
  
  document.getElementById('modalBackdrop').classList.add('show');
  document.body.classList.add('modal-open');
}

function openCMTSDetail(cmts) {
  const cmtsDataSource = window.currentCMTSData || allCMTS;
  
  if (!cmtsDataSource || !cmtsDataSource.length) {
    toast('❌ No hay datos de CMTS disponibles');
    return;
  }
  
  const cmtsData = cmtsDataSource.find(c => c.cmts === cmts);
  
  if (!cmtsData) {
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
  
  document.getElementById('modalFilters').style.display = 'flex';
  document.getElementById('modalFooter').innerHTML = `
    <div class="selection-info" id="selectionInfo">0 órdenes seleccionadas</div>
    <div style="display: flex; gap: 8px;">
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

function renderModalContent() {
  const horario = document.getElementById('filterHorario').value;
  const dia = document.getElementById('filterDia').value;
  
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
  
  const chartCounts = UIRenderer.normalizeCounts(currentZone.last7DaysCounts);
  const chartLabels = Array.isArray(currentZone.last7Days) ? currentZone.last7Days : [];
  const maxChartValue = Math.max(...chartCounts, 1);

  let html = '<div class="chart-container">';
  html += '<div class="chart-title">📊 Distribución de Ingresos (7 días)</div>';

  if (!chartCounts.length) {
    html += '<div class="sparkline-placeholder" style="height: 150px; display: flex; align-items: center; justify-content: center;">Sin datos</div>';
  } else {
    html += '<div style="display: flex; justify-content: space-around; align-items: flex-end; height: 150px; padding: 10px;">';

    chartCounts.forEach((count, i) => {
      const barHeight = maxChartValue > 0 ? (count / maxChartValue) * 120 : 0;
      const label = chartLabels[i] || '';
      html += `<div style="text-align: center;">
        <div style="width: 60px; height: ${barHeight}px; background: linear-gradient(180deg, #0078D4 0%, #005A9E 100%); border-radius: 4px 4px 0 0; margin: 0 auto;"></div>
        <div style="font-size: 0.75rem; font-weight: 700; margin-top: 4px;">${count}</div>
        <div style="font-size: 0.6875rem; color: var(--text-secondary);">${label}</div>
      </div>`;
    });

    html += '</div>';
  }

  html += '</div>';
  
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
    const cita = o['Número de cita'] || `row_${idx}`;
    const checked = selectedOrders.has(cita) ? 'checked' : '';
    html += `<tr>`;
    html += `<td><input type="checkbox" class="order-checkbox" data-cita="${cita}" ${checked} onchange="updateSelection()"></td>`;
    
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

function exportExcelVista(){
  const hasZones = Array.isArray(window.currentAnalyzedZones) && window.currentAnalyzedZones.length;
  const hasOrders = Array.isArray(window.lastFilteredOrders) && window.lastFilteredOrders.length;

  if (!hasZones && !hasOrders) {
    toast('No hay datos para exportar');
    return;
  }

  const wb = XLSX.utils.book_new();
  const usedNames = new Set();

  if (hasZones){
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
    const zonasSheet = XLSX.utils.json_to_sheet(zonasData);
    appendSheet(wb, zonasSheet, 'Zonas', usedNames);
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
    const cmtsSheet = XLSX.utils.json_to_sheet(cmtsData);
    appendSheet(wb, cmtsSheet, 'CMTS', usedNames);
  }

  if (window.edificiosData && window.edificiosData.length){
    const edi = window.edificiosData.map(e => ({
      Direccion: e.direccion,
      Zona: e.zona,
      Territorio: e.territorio,
      Total_OTs: e.casos.length
    }));
    const ediSheet = XLSX.utils.json_to_sheet(edi);
    appendSheet(wb, ediSheet, 'Edificios', usedNames);
  }

  if (window.equiposPorZona && window.equiposPorZona.size){
    const todos = [];
    window.equiposPorZona.forEach((arr, zona) => {
      arr.forEach(it => todos.push({ ...it, zona }));
    });
    const equiposSheet = XLSX.utils.json_to_sheet(todos);
    appendSheet(wb, equiposSheet, 'Equipos', usedNames);
  }

  if (hasZones){
    const detalleRows = [];
    window.currentAnalyzedZones.forEach(z => {
      const rows = buildZoneExportRows(z);
      if (rows && rows.length) {
        detalleRows.push(...rows);
      }
    });
    const detalleSheet = createWorksheetFromRows(detalleRows, ZONE_EXPORT_HEADERS);
    appendSheet(wb, detalleSheet, 'Detalle_Zonas', usedNames);
  }

  if (!wb.SheetNames || wb.SheetNames.length === 0){
    toast('No hay datos para exportar');
    return;
  }

  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Vista_Filtrada_${fecha}.xlsx`);
  toast('✓ Vista detallada exportada');
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
  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Zonas_Crudo_${fecha}.xlsx`);
  toast('✓ Zonas exportadas');
}

function exportModalDetalleExcel(){
  const title = document.getElementById('modalTitle').textContent || '';
  const wb = XLSX.utils.book_new();

  if (title.startsWith('Detalle: ') && window.currentZone){
    const ordenes = window.currentZone.ordenes || [];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ordenes), `Zona_${window.currentZone.zona || 'NA'}`);
  } else if (title.startsWith('Zonas del CMTS: ')){
    const cmts = title.replace('Zonas del CMTS: ', '').trim();
    const data = (window.currentCMTSData || []).find(c => c.cmts === cmts);
    if (!data) return toast('Sin datos de CMTS');
    const zonasFlat = data.zonas.map(z => ({
      Zona: z.zona, Tipo: z.tipo, Total_OTs: z.totalOTs, Ingreso_N: z.ingresoN, Ingreso_N1: z.ingresoN1, Estado_Nodo: z.nodoEstado
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(zonasFlat), `CMTS_${cmts.slice(0, 20)}`);
  } else {
    const edTitle = document.getElementById('edificioModalTitle')?.textContent || '';
    if (edTitle.startsWith('🏢 Edificio: ') && window.edificiosData){
      const dir = edTitle.replace('🏢 Edificio: ', '').trim();
      const e = window.edificiosData.find(x => x.direccion === dir);
      if (!e) return toast('Sin datos del edificio');
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(e.casos), 'Edificio_OTs');
    } else {
      return toast('No hay detalle para exportar');
    }
  }

  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Detalle_${fecha}.xlsx`);
  toast('✓ Detalle exportado');
}

console.log('✅ Panel v5.0 COMPLETO inicializado');
console.log('🎯 FUNCIONALIDADES:');
console.log('   • Territorios Críticos con agrupación Flex');
console.log('   • Filtros múltiples de equipos (modelo/marca/territorio)');
console.log('   • Stats clickeables con modales interactivos');
console.log('   • Exportaciones completas (Zonas/CMTS/Equipos/BEFAN)');
console.log('   • Planillas funcionales en nueva ventana');
console.log('📊 Estados activos:', CONFIG.estadosPermitidos);
