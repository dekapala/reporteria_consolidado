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

function normalizeDevice(device){
  if (!device || typeof device !== 'object') return null;
  const mapValue = (value) => (value === null || value === undefined) ? '' : String(value).trim();

  const macRaw = device.mac || device.macAddress || device.mac_address || device.MAC || device.macAddressRaw || '';
  const macFormatted = (typeof TextUtils !== 'undefined' && TextUtils?.formatMac)
    ? TextUtils.formatMac(macRaw)
    : (normalizeMac(macRaw) || mapValue(macRaw));
  const macValue = mapValue(macFormatted);
  const serialNumber = mapValue(device.serialNumber || device.serial || device.serie || device.sn || device.numeroSerie);
  const model = mapValue(device.model || device.modelo);
  const category = mapValue(device.category || device.brand || device.marca);
  const brand = mapValue(device.brand || device.category || device.marca || category);
  const technology = mapValue(device.technology || device.description || device.detalle);
  const type = mapValue(device.type || device.tipo);

  if (!(macValue || serialNumber || model || category || technology || type || brand)) {
    return null;
  }

  return {
    macAddress: macValue,
    serialNumber,
    model,
    category,
    technology,
    type,
    brand,
    mac: macValue
  };
}

function pickPreferredDevice(devs=[]){
  if (!Array.isArray(devs) || !devs.length) return null;

  const normalizedDevices = devs
    .map(normalizeDevice)
    .filter(Boolean);

  if (!normalizedDevices.length) return null;

  const withMac = normalizedDevices.find(d => d.mac);
  if (withMac) return withMac;

  const withModel = normalizedDevices.find(d => d.model);
  if (withModel) return withModel;

  return normalizedDevices[0];
}

function ensureOrderDeviceMeta(order){
  if (!order || typeof order !== 'object') return null;

  const meta = order.__meta = order.__meta || {};

  let dispositivos = [];

  if (Array.isArray(meta.dispositivos) && meta.dispositivos.length) {
    dispositivos = meta.dispositivos;
  } else {
    const colInfo = findDispositivosColumn(order);
    if (colInfo && order[colInfo]) {
      dispositivos = TextUtils.parseDispositivosJSON(order[colInfo]);
    }
  }

  const normalizedDevices = Array.isArray(dispositivos)
    ? dispositivos.map(normalizeDevice).filter(Boolean)
    : [];

  meta.dispositivos = normalizedDevices;

  const device = pickPreferredDevice(normalizedDevices);
  meta.device = device || null;

  return meta.device;
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

const TextUtils = window.TextUtils || {
  normalize(text) {
    return String(text || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
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
    if (!jsonStr || typeof jsonStr !== 'string') return [];

    let s = jsonStr.trim()
      .replace(/&quot;/g, '"')
      .replace(/\r?\n/g, ' ')
      .replace(/\s{2,}/g, ' ')
      .replace(/;+\s*$/g, '');

    const first = s.indexOf('{');
    const last = s.lastIndexOf('}');
    if (first !== -1 && last > first) {
      s = s.slice(first, last + 1);
    }

    if (/openResp/i.test(s) && !/\}\s*\]\s*\}/.test(s)) {
      if (/\}\s*$/.test(s)) s = s.replace(/\}\s*$/, '}]}');
      if (/\]\s*$/.test(s)) s = s.replace(/\]\s*$/, ']}');
    }

    const normalizeValue = (value) => {
      if (value === null || value === undefined) return '';
      return String(value).trim();
    };

    const tryParse = () => {
      try {
        const parsed = JSON.parse(s);
        let arr = null;

        if (Array.isArray(parsed)) {
          arr = parsed;
        } else if (parsed && typeof parsed === 'object') {
          const candidates = parsed.openResp || parsed.open_response || parsed.data;
          if (Array.isArray(candidates)) {
            arr = candidates;
          } else if (candidates && typeof candidates === 'object') {
            arr = [candidates];
          }
        }

        if (!arr) return undefined;

        return arr.map(d => ({
          category: normalizeValue(d?.category),
          description: normalizeValue(d?.description),
          model: normalizeValue(d?.model),
          serialNumber: normalizeValue(d?.serialNumber || d?.serial),
          macAddress: normalizeValue(d?.macAddress || d?.mac),
          type: normalizeValue(d?.type)
        }));
      } catch {
        return null;
      }
    };

    const parsedDevices = tryParse();
    if (Array.isArray(parsedDevices)) {
      return parsedDevices;
    }

    const pick = (key) => {
      const regex = new RegExp(`"${key}"\\s*:\\s*"([^"\\]+)"`, 'i');
      const match = s.match(regex);
      return match ? match[1].trim() : '';
    };

    const fallback = {
      category: pick('category'),
      description: pick('description'),
      model: pick('model'),
      serialNumber: pick('serialNumber') || pick('serial'),
      macAddress: pick('macAddress') || pick('mac'),
      type: pick('type')
    };

    if (Object.values(fallback).some(Boolean)) {
      return [fallback];
    }

    return [];
  },

  formatMac(mac) {
    if (!mac) return '';
    const hex = String(mac).replace(/[^0-9A-Fa-f]/g, '').toUpperCase();
    if (hex.length !== 12) return String(mac).trim();
    return hex.match(/.{1,2}/g).join(':');
  },

  detectarSistema(numCaso) {
    const numStr = String(numCaso || '').trim();
    if (numStr.startsWith('8')) return 'OPEN';
    if (numStr.startsWith('3')) return 'FAN';
    return '';
  }
};
window.TextUtils = TextUtils;

function buildEquiposFromSheet(rows) {
  const equipos = [];
  if (!Array.isArray(rows)) return equipos;

  rows.forEach(row => {
    if (!row || typeof row !== 'object') return;

    const col = findDispositivosColumn(row);
    const raw = (col && row[col])
      || row['Informacion Dispositivos']
      || row['Información Dispositivos']
      || row['Informacion_Dispositivos']
      || row['info_dispositivos']
      || '';

    if (!raw) return;

    const dispositivos = TextUtils.parseDispositivosJSON(raw);
    if (!Array.isArray(dispositivos) || !dispositivos.length) return;

    const fecha = row['Fecha']
      || row['Fecha de creación']
      || row['Fecha/Hora de apertura']
      || row['Fecha de inicio']
      || '';

    const zona = row['Zona/CAC']
      || row['Zona Tecnica HFC']
      || row['Zona Tecnica FTTH']
      || row['Zona Tecnica']
      || row['Zona']
      || '';

    const territorio = row['Territorio de servicio: Nombre'] || row['Territorio'] || '';

    const caso = row['Número del caso']
      || row['Numero del caso']
      || row['Caso Externo']
      || row['Caso']
      || row['External Case Id']
      || '';

    const numeroOrden = row['Numero Orden']
      || row['Número Orden']
      || row['Número de cita']
      || row['Numero de cita']
      || row['Orden']
      || row['Orden de trabajo']
      || '';

    const sistema = TextUtils.detectarSistema(caso);

    const clean = (value) => {
      if (value === null || value === undefined) return '';
      return String(value).trim();
    };

    dispositivos.forEach(d => {
      const mac = TextUtils.formatMac(d.macAddress || d.serialNumber);
      const tipoValue = clean(d.type) || clean(d.description);
      equipos.push({
        Zona: zona || '',
        Territorio: territorio || '',
        Caso: caso || '',
        Sistema: sistema || '',
        SerialNumber: clean(d.serialNumber || d.serial),
        MAC: mac || '',
        Modelo: clean(d.model),
        Tipo: tipoValue,
        Marca: clean(d.category),
        Categoria: clean(d.category),
        Descripcion: clean(d.description),
        Fecha: fecha || '',
        NumeroOrden: numeroOrden || ''
      });
    });
  });

  return equipos;
}

function enrichRowWithDeviceInfo(row) {
  if (!row || typeof row !== 'object') return row;

  const needModelo = !row['Modelo'];
  const needMac = !row['MAC'];
  const needTipo = !row['Tipo'];

  if (!(needModelo || needMac || needTipo)) return row;

  const meta = row.__meta = row.__meta || {};
  let device = ensureOrderDeviceMeta(row);

  if (!device) {
    const fallback = Array.isArray(meta.dispositivos) && meta.dispositivos.length
      ? meta.dispositivos[0]
      : null;
    if (fallback) {
      device = fallback;
    }
  }

  if (!device) {
    const col = findDispositivosColumn(row);
    const raw = (col && row[col])
      || row['Informacion Dispositivos']
      || row['Información Dispositivos']
      || row['Informacion_Dispositivos']
      || row['info_dispositivos']
      || '';

    const parsed = TextUtils.parseDispositivosJSON(raw);
    if (parsed.length) {
      const normalizedList = parsed.map(normalizeDevice).filter(Boolean);
      if (normalizedList.length) {
        meta.dispositivos = normalizedList;
        meta.device = normalizedList[0];
        device = meta.device;
      } else {
        device = parsed[0];
      }
    }
  }

  if (!device) return row;

  const modelo = device.model || device.Modelo || device.modelo || '';
  const tipo = device.type || device.Tipo || device.technology || device.description || '';
  const macCandidate = device.mac || device.MAC || device.macAddress || device.serialNumber || device.SerialNumber || '';
  const formattedMac = TextUtils.formatMac(macCandidate);

  if (needModelo && modelo) row['Modelo'] = modelo;
  if (needTipo && tipo) row['Tipo'] = tipo;
  if (needMac && formattedMac) row['MAC'] = formattedMac;

  return row;
}

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

function escapeHtml(str){
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function toggleUtilities() {
  const menu = document.getElementById('utilitiesMenu');
  if (!menu) return;
  menu.classList.toggle('show');
}

document.addEventListener('click', (e) => {
  const dropdown = document.querySelector('.utilities-dropdown');
  const menu = document.getElementById('utilitiesMenu');
  if (dropdown && !dropdown.contains(e.target) && menu) {
    menu.classList.remove('show');
  }
});

function openPlanillasNewTab() {
  document.getElementById('utilitiesMenu').classList.remove('show');

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
  document.getElementById('utilitiesMenu').classList.remove('show');

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
  }
  
  async loadExcel(file, tipo) {
    try {
      const data = await readFileAsUint8Array(file);
      
      const wb = XLSX.read(data, {
        type: 'array',
        cellDates: false,
        cellText: false,
        cellFormula: false,
        raw: false
      });

      if (!wb.SheetNames || wb.SheetNames.length === 0) {
        throw new Error('El archivo no contiene hojas');
      }

      const firstSheetName = wb.SheetNames[0];
      const ws = wb.Sheets[firstSheetName];
      if (!ws) {
        throw new Error('No se pudo abrir la primera hoja');
      }

      let ref = ws['!ref'];
      if (!ref) {
        const keys = Object.keys(ws).filter(k => k[0] !== '!');
        if (keys.length === 0) {
          throw new Error('La hoja está vacía');
        }
        let minR = Infinity, minC = Infinity, maxR = 0, maxC = 0;
        for (const addr of keys) {
          const { r, c } = XLSX.utils.decode_cell(addr);
          if (r < minR) minR = r;
          if (c < minC) minC = c;
          if (r > maxR) maxR = r;
          if (c > maxC) maxC = c;
        }
        ref = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
        ws['!ref'] = ref;
      }

      let range;
      try {
        range = XLSX.utils.decode_range(ws['!ref']);
      } catch (err) {
        range = { s: { r: 0, c: 0 }, e: { r: 999, c: 50 } };
        ws['!ref'] = XLSX.utils.encode_range(range);
      }
      
      let headerRow = 0;
      let headerDetected = false;
      for (let R = 0; R <= 20; R++) {
        let hasContent = false;
        for (let C = 0; C <= 15; C++) {
          const cellAddr = XLSX.utils.encode_cell({r: R, c: C});
          const cell = ws[cellAddr];
          if (cell && cell.v) {
            const cellValue = String(cell.v).toLowerCase();
            if (cellValue.includes('zona') && cellValue.includes('tecnica')) {
              headerRow = R;
              headerDetected = true;
              console.log(`✅ Headers detectados en fila ${R + 1} (columna ${C}): "${cell.v}"`);
              hasContent = true;
              break;
            }
            if (cellValue.includes('numero') && (cellValue.includes('caso') || cellValue.includes('cuenta'))) {
              headerRow = R;
              headerDetected = true;
              console.log(`✅ Headers detectados en fila ${R + 1} (columna ${C}): "${cell.v}"`);
              hasContent = true;
              break;
            }
          }
        }
        if (hasContent) break;
      }

      if (!headerDetected) {
        headerRow = 0;
        console.warn('⚠️ No se detectaron headers de zona/caso. Se usará la primera fila como encabezado.');
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
      const buffer = await readFileAsUint8Array(file);
      const decoder = new TextDecoder('utf-8');
      const text = decoder.decode(buffer || new Uint8Array());

      if (!text.trim()) {
        return {success: false, error: 'CSV vacío'};
      }

      const normalizedText = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      const rows = [];
      let currentField = '';
      let currentRow = [];
      let inQuotes = false;

      for (let i = 0; i < normalizedText.length; i++) {
        const char = normalizedText[i];

        if (char === '"') {
          if (inQuotes && normalizedText[i + 1] === '"') {
            currentField += '"';
            i++;
          } else {
            inQuotes = !inQuotes;
          }
        } else if (char === ',' && !inQuotes) {
          currentRow.push(currentField.trim());
          currentField = '';
        } else if (char === '\n' && !inQuotes) {
          currentRow.push(currentField.trim());
          rows.push(currentRow);
          currentRow = [];
          currentField = '';
        } else {
          currentField += char;
        }
      }

      if (currentField.length || currentRow.length) {
        currentRow.push(currentField.trim());
        rows.push(currentRow);
      }

      if (rows.length < 2) {
        return {success: false, error: 'CSV vacío'};
      }

      const headers = rows[0];
      const dataRows = rows.slice(1).filter(r => r.some(cell => cell !== ''));
      const mappedRows = dataRows.map(values => {
        const obj = {};
        headers.forEach((h, idx) => {
          obj[h] = values[idx] || '';
        });
        return obj;
      });

      console.log(`✅ CSV Alarmas: ${mappedRows.length} alarmas leídas`);
      this.fmsData = mappedRows;
      this.processFMS();

      return {success: true, rows: mappedRows.length};
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

    if (!(this.fmsMap instanceof Map)) {
      this.fmsMap = new Map();
    } else {
      this.fmsMap.clear();
    }

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
    
    this.consolidado1.forEach(r => {
      const cita = r['Número de cita'];
      if (cita) {
        if (!map.has(cita)) {
          map.set(cita, { ...r, _merged: false });
        }
      }
    });
    
    this.consolidado2.forEach(r => {
      const cita = r['Número de cita'];
      if (cita) {
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
      }
    });
    
    const result = Array.from(map.values());
    console.log(`✅ Total órdenes únicas (deduplicadas): ${result.length}`);
    
    return result;
  }
  
  processZones(rows) {
    const zoneGroups = new Map();

    rows.forEach(r => {
      const order = enrichRowWithDeviceInfo(r);
      ensureOrderDeviceMeta(order);
      const {zonaPrincipal, tipo} = this.getZonaPrincipal(order);
      if (!zonaPrincipal) return;

      if (!zoneGroups.has(zonaPrincipal)) {
        zoneGroups.set(zonaPrincipal, {
          zona: zonaPrincipal,
          tipo: tipo,
          zonaHFC: order['Zona Tecnica HFC'] || order['Zona Tecnica'] || '',
          zonaFTTH: order['Zona Tecnica FTTH'] || '',
          territorio: order['Territorio de servicio: Nombre'] || order['Territorio'] || '',
          ordenes: []
        });
      }

      zoneGroups.get(zonaPrincipal).ordenes.push(order);
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

          const diag = extractDiagnosticoTecnico(o);
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

// ====== Persistencia de filtros (localStorage) ======
const FilterPersistence = {
  key: 'panel-v5-filters',
  load() {
    try { return JSON.parse(localStorage.getItem(this.key) || '{}'); }
    catch { return {}; }
  },
  save(data) {
    try { localStorage.setItem(this.key, JSON.stringify(data || {})); }
    catch (e) { console.warn('No se pudo guardar filtros', e); }
  }
};

function persistFilters() {
  const snapshot = {
    // toggles / selects principales
    catec: document.getElementById('filterCATEC')?.checked || false,
    excludeCatec: document.getElementById('filterExcludeCATEC')?.checked || false,
    showAllStates: document.getElementById('showAllStates')?.checked || false,
    ftth: document.getElementById('filterFTTH')?.checked || false,
    excludeFTTH: document.getElementById('filterExcludeFTTH')?.checked || false,
    nodoEstado: document.getElementById('filterNodoEstado')?.value || '',
    cmts: document.getElementById('filterCMTS')?.value || '',
    days: parseInt(document.getElementById('daysWindow')?.value || CONFIG.defaultDays, 10),
    territorio: document.getElementById('filterTerritorio')?.value || '',
    sistema: document.getElementById('filterSistema')?.value || '',
    alarma: document.getElementById('filterAlarma')?.value || '',
    quickSearch: document.getElementById('quickSearch')?.value || '',
    ordenarPorIngreso: document.getElementById('ordenarPorIngreso')?.value || 'desc',

    // multiselect de zonas
    selectedZonas: Array.from(zoneFilterState?.selected || []),

    // filtros del panel de equipos
    equipoModelo: Array.isArray(Filters?.equipoModelo) ? Filters.equipoModelo : [],
    equipoMarca: Filters?.equipoMarca || '',
    equipoTerritorio: Filters?.equipoTerritorio || ''
  };
  FilterPersistence.save(snapshot);
}

function restoreFiltersToControls(saved = {}) {
  const set = (id, prop, value) => {
    const el = document.getElementById(id);
    if (el != null) el[prop] = value;
  };

  set('filterCATEC','checked', !!saved.catec);
  set('filterExcludeCATEC','checked', !!saved.excludeCatec);
  set('showAllStates','checked', !!saved.showAllStates);
  set('filterFTTH','checked', !!saved.ftth);
  set('filterExcludeFTTH','checked', !!saved.excludeFTTH);
  set('filterNodoEstado','value', saved.nodoEstado || '');
  set('filterCMTS','value', saved.cmts || '');
  set('daysWindow','value', String(saved.days || CONFIG.defaultDays));
  set('filterTerritorio','value', saved.territorio || '');
  set('filterSistema','value', saved.sistema || '');
  set('filterAlarma','value', saved.alarma || '');
  set('quickSearch','value', saved.quickSearch || '');
  set('ordenarPorIngreso','value', saved.ordenarPorIngreso || 'desc');

  // filtros de equipos
  Filters.equipoModelo = Array.isArray(saved.equipoModelo) ? saved.equipoModelo : [];
  Filters.equipoMarca = saved.equipoMarca || '';
  Filters.equipoTerritorio = saved.equipoTerritorio || '';

  // zonas seleccionadas (espera a que existan opciones)
  if (Array.isArray(saved.selectedZonas) && saved.selectedZonas.length) {
    Filters.selectedZonas = saved.selectedZonas;
    zoneFilterState.selected = new Set(saved.selectedZonas);
  }
}

const dataProcessor = new DataProcessor();

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

  if (!zoneFilterState.allOptions.length) {
    container.innerHTML = '<div class="zone-multiselect__empty">Cargá los reportes para habilitar el filtro por zonas</div>';
    return;
  }

  if (!zoneFilterState.filteredOptions.length) {
    container.innerHTML = '<div class="zone-multiselect__empty">No hay coincidencias para la búsqueda actual</div>';
    return;
  }

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

function onZoneFilterSearch(event) {
  const value = event.target.value || '';
  zoneFilterState.search = value;

  if (!zoneFilterState.allOptions.length) {
    renderZoneFilterOptions();
    return;
  }

  const normalizedTerm = TextUtils.normalize(value);
  if (!normalizedTerm) {
    zoneFilterState.filteredOptions = zoneFilterState.allOptions.slice();
  } else {
    zoneFilterState.filteredOptions = zoneFilterState.allOptions.filter(z => TextUtils.normalize(z).includes(normalizedTerm));
  }

  renderZoneFilterOptions();
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

  source.forEach(z => zoneFilterState.selected.add(z));

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
    zoneFilterState.filteredOptions = zoneFilterState.allOptions.slice();
    const searchInput = document.getElementById('zoneFilterSearch');
    if (searchInput) searchInput.value = '';
  } else {
    const normalizedTerm = TextUtils.normalize(zoneFilterState.search);
    if (!normalizedTerm) {
      zoneFilterState.filteredOptions = zoneFilterState.allOptions.slice();
    } else {
      zoneFilterState.filteredOptions = zoneFilterState.allOptions.filter(z => TextUtils.normalize(z).includes(normalizedTerm));
    }
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
    html += '<th>Diagnóstico técnico</th>';
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
      
      const diagnosticos = Array.isArray(z.diagnosticos) ? z.diagnosticos.filter(Boolean) : [];
      let diagnosticoHtml = '<span class="diagnostico-tag diagnostico-tag--empty">Sin datos</span>';
      if (diagnosticos.length) {
        const chips = diagnosticos.slice(0, 3).map(d => {
          const raw = String(d);
          const title = escapeHtml(raw);
          const truncated = raw.length > 80 ? `${raw.slice(0, 77)}…` : raw;
          const label = escapeHtml(truncated);
          return `<span class="diagnostico-tag" title="${title}">${label}</span>`;
        }).join('');
        const extra = diagnosticos.length > 3
          ? `<span class="diagnostico-tag diagnostico-tag--more">+${diagnosticos.length - 3}</span>`
          : '';
        diagnosticoHtml = `<div class="diagnostico-tags">${chips}${extra}</div>`;
      }

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
        <td>${diagnosticoHtml}</td>
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
    const edificios = new Map();
    
    ordenes.forEach(o => {
      const dir = TextUtils.normalize(o['Calle']);
      if (!dir || dir.length < 5) return;
      
      if (!edificios.has(dir)) {
        edificios.set(dir, {
          direccion: o['Calle'],
          zona: o['Zona Tecnica HFC'] || o['Zona Tecnica FTTH'],
          territorio: o['Territorio de servicio: Nombre'] || '',
          casos: []
        });
      }
      edificios.get(dir).casos.push(o);
    });
    
    const sorted = Array.from(edificios.values())
      .filter(e => e.casos.length >= 2)
      .sort((a, b) => b.casos.length - a.casos.length)
      .slice(0, 50);
    
    if (!sorted.length) return '<div class="loading-message"><p>No hay edificios con 2+ incidencias</p></div>';
    
    let html = '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
    html += '<th>Dirección</th><th>Zona</th><th>Territorio</th><th class="number">Total OTs</th><th>Acción</th>';
    html += '</tr></thead><tbody>';

    sorted.forEach((e, idx) => {
      const direccion = escapeHtml(e.direccion || '');
      const zona = escapeHtml(e.zona || '');
      const territorio = escapeHtml(e.territorio || '');
      html += `<tr>
        <td><strong>${direccion}</strong></td>
        <td>${zona}</td>
        <td>${territorio}</td>
        <td class="number">${e.casos.length}</td>
        <td><button type="button" class="btn btn-primary btn-ver-edificio" data-index="${idx}" data-direccion="${direccion}">👁️ Ver</button></td>
      </tr>`;
    });
    
    html += '</tbody></table></div></div>';
    
    window.edificiosData = sorted;
    
    return html;
  },
  
  renderEquipos(ordenes) {
    const equipos = buildEquiposFromSheet(ordenes);
    if (!equipos.length) {
      return '<div class="loading-message"><p>⚠️ No se encontraron equipos en las órdenes.</p></div>';
    }

    const grupos = new Map();
    const territorios = new Set();
    const modelos = new Set();
    const marcas = new Set();

    equipos.forEach(item => {
      const zona = item.Zona || '';
      const entry = {
        zona,
        territorio: item.Territorio || '',
        numCaso: item.Caso || '',
        sistema: item.Sistema || '',
        serialNumber: item.SerialNumber || '',
        macAddress: item.MAC || '',
        tipo: item.Tipo || '',
        marca: item.Marca || item.Categoria || '',
        modelo: item.Modelo || '',
        descripcion: item.Descripcion || '',
        fecha: item.Fecha || '',
        numeroOrden: item.NumeroOrden || ''
      };

      if (!grupos.has(zona)) grupos.set(zona, []);
      grupos.get(zona).push(entry);

      if (entry.territorio) territorios.add(entry.territorio);
      if (entry.modelo) modelos.add(entry.modelo);
      if (entry.marca) marcas.add(entry.marca);
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

    const gruposFiltrados = new Map();
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

    if (!zonasFiltradas.length) {
      html += '<div class="loading-message"><p>No hay equipos que coincidan con los filtros seleccionados</p></div>';
      return html;
    }

    html += '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
    html += '<th>Zona</th><th class="number">Equipos</th><th>Acción</th>';
    html += '</tr></thead><tbody>';

    zonasFiltradas.forEach((z) => {
      const arr = gruposFiltrados.get(z) || [];
      const open = window.equiposOpen.has(z);
      html += `<tr class="clickable" onclick="toggleEquiposGrupo('${z.replace(/'/g, "\'")}')">`;
      html += `<td><strong>${z || '-'}</strong></td>`;
      html += `<td class="number">${arr.length}</td>`;
      html += `<td><button class="btn btn-secondary" style="padding:6px 12px; font-size:0.75rem;" onclick="event.stopPropagation(); exportEquiposGrupoExcel('${z.replace(/'/g, "\'")}', true)">📥 Exportar</button></td>`;
      html += '</tr>';

      if (open) {
        html += `<tr><td colspan="3">`;
        html += '<div class="table-container"><div class="table-wrapper"><table class="detail-table"><thead><tr>';
        html += '<th>Caso</th><th>Sistema</th><th>Serial</th><th>MAC</th><th>Tipo</th><th>Marca</th><th>Modelo</th>';
        html += '</tr></thead><tbody>';
        arr.slice(0, 1000).forEach(e => {
          const badge = e.sistema === 'OPEN' ? '<span class="badge badge-open">OPEN</span>'
                       : e.sistema === 'FAN' ? '<span class="badge badge-fan">FAN</span>' : '';
          html += `<tr>`;
          html += `<td style="font-family:monospace">${e.numCaso || ''}</td>`;
          html += `<td>${badge}</td>`;
          html += `<td style="font-family:monospace">${e.serialNumber || ''}</td>`;
          html += `<td style="font-family:monospace">${e.macAddress || ''}</td>`;
          html += `<td>${e.tipo || ''}</td>`;
          html += `<td>${e.marca || ''}</td>`;
          html += `<td>${e.modelo || ''}</td>`;
          html += '</tr>';
        });
        html += '</tbody></table></div></div>';
        html += `</td></tr>`;
      }
    });

    html += '</tbody></table></div></div>';
    html += `<p style="text-align:center;margin-top:12px;color:var(--text-secondary)">Total: ${zonasFiltradas.length} zonas • Click para expandir/colapsar</p>`;

    window.equiposPorZona = gruposFiltrados;
    window.equiposPorZonaCompleto = grupos;
    window.equiposListado = equipos;

    return html;
  }
};

const FMSPanel = {
  render(ordenes, fmsMap) {
    if (!(fmsMap instanceof Map) || fmsMap.size === 0) {
      return '<div class="loading-message"><p>⚠️ Cargá el archivo de alarmas FMS para ver resultados.</p></div>';
    }

    if (!Array.isArray(ordenes) || ordenes.length === 0) {
      return '<div class="loading-message"><p>Sin resultados</p></div>';
    }

    const normalizedFmsMap = new Map();
    fmsMap.forEach((alarms, zoneKey) => {
      const baseKey = String(zoneKey || '').trim();
      if (!baseKey) return;
      if (!normalizedFmsMap.has(baseKey)) normalizedFmsMap.set(baseKey, alarms);
      const upperKey = baseKey.toUpperCase();
      if (!normalizedFmsMap.has(upperKey)) normalizedFmsMap.set(upperKey, alarms);
    });

    const aggregated = new Map();

    ordenes.forEach(order => {
      const candidates = new Set();
      const meta = order.__meta || {};
      if (meta.zona) candidates.add(meta.zona);
      if (meta.zonaPrincipal) candidates.add(meta.zonaPrincipal);
      if (meta.zonaHFC) candidates.add(meta.zonaHFC);
      if (meta.zonaFTTH) candidates.add(meta.zonaFTTH);
      const zonaHFC = pickFirstValue(order, ORDER_FIELD_KEYS.zonaHFC);
      if (zonaHFC) candidates.add(zonaHFC);
      const zonaFTTH = pickFirstValue(order, ORDER_FIELD_KEYS.zonaFTTH);
      if (zonaFTTH) candidates.add(zonaFTTH);

      candidates.forEach(candidate => {
        const base = String(candidate || '').trim();
        if (!base) return;
        const alarms = normalizedFmsMap.get(base) || normalizedFmsMap.get(base.toUpperCase());
        if (!alarms || !alarms.length) return;
        const storageKey = base.toUpperCase();
        if (!aggregated.has(storageKey)) {
          aggregated.set(storageKey, { zona: base, display: candidate, orders: [], alarms });
        }
        const entry = aggregated.get(storageKey);
        if (!entry.display) entry.display = candidate;
        entry.orders.push(order);
        entry.alarms = alarms;
      });
    });

    if (!aggregated.size) {
      return '<div class="loading-message"><p>Sin resultados</p></div>';
    }

    const rows = Array.from(aggregated.values()).map(entry => {
      const alarms = entry.alarms || [];
      const activas = alarms.filter(a => a?.isActive).length;
      const total = alarms.length;
      const causal = getCausalForZona(entry.zona, entry.orders, fmsMap);
      const latest = alarms.reduce((memo, alarm) => {
        if (!alarm) return memo;
        if (!memo) return alarm;
        const currentDate = new Date(alarm.recoveryDate || alarm.creationDate || 0);
        const memoDate = new Date(memo.recoveryDate || memo.creationDate || 0);
        return currentDate > memoDate ? alarm : memo;
      }, null);
      const ultima = latest ? (latest.recoveryDate || latest.creationDate || '') : '';
      return {
        zona: entry.display || entry.zona,
        activas,
        total,
        causal,
        ultima: ultima || '—'
      };
    }).sort((a, b) => {
      if (b.activas !== a.activas) return b.activas - a.activas;
      if (b.total !== a.total) return b.total - a.total;
      return (a.zona || '').localeCompare(b.zona || '');
    });

    let html = '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
    html += '<th>Zona</th><th class="number">Alarmas activas</th><th class="number">Total alarmas</th><th>Causal principal</th><th>Última actualización</th>';
    html += '</tr></thead><tbody>';

    rows.forEach(row => {
      html += `<tr>
        <td>${escapeHtml(row.zona || '')}</td>
        <td class="number">${row.activas}</td>
        <td class="number">${row.total}</td>
        <td>${escapeHtml(row.causal || 'N/D')}</td>
        <td>${escapeHtml(row.ultima || '—')}</td>
      </tr>`;
    });

    html += '</tbody></table></div></div>';
    return html;
  }
};

function applyEquiposFilters() {
  const select = document.getElementById('filterEquipoModelo');
  Filters.equipoModelo = Array.from(select.selectedOptions).map(opt => opt.value);
  Filters.equipoMarca = document.getElementById('filterEquipoMarca').value;
  Filters.equipoTerritorio = document.getElementById('filterEquipoTerritorio').value;
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
  persistFilters();
}

function clearEquiposFilters() {
  Filters.equipoModelo = [];
  Filters.equipoMarca = '';
  Filters.equipoTerritorio = '';
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
  persistFilters();
}

function toggleEquiposGrupo(zona){
  if (!window.equiposOpen) window.equiposOpen = new Set();
  if (window.equiposOpen.has(zona)) window.equiposOpen.delete(zona);
  else window.equiposOpen.add(zona);
  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(window.lastFilteredOrders || []);
}

const ZONE_EXPORT_HEADERS = [
  'Fecha',
  'Zona/CAC',
  'Caso',
  'Numero Orden',
  'Diagnostico Tecnico',
  'Tipo Trabajo',
  'MAC',
  'Modelo',
  'Tipo',
  'tipo_daño'
];

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
  diagnostico: [
    'Diagnostico Tecnico',
    'Diagnóstico Técnico',
    'Diagnostico',
    'Diagnóstico',
    'Diagnostico Cliente',
    'Diagnostico tecnico'
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

const ZONE_CAUSAL_KEYS = [
  'Tipo de daño',
  'Tipo de daño/causal',
  'Tipo de daño / Causal',
  'Tipo Daño',
  'Tipo_Daño',
  'Tipo dano',
  'Causal',
  'Causal reportada',
  'Causal reclamo',
  'Tipo de Causal',
  'Daño',
  'Damage',
  'damageClassification',
  'Damage Classification',
  'incidentClassification',
  'Incident Classification'
];

function pickFirstValue(obj, keys){
  if (!obj || !keys) return '';
  for (const key of keys){
    if (!key) continue;
    const value = obj[key];
    if (value !== null && value !== undefined){
      const str = String(value).trim();
      if (str) return str;
    }
  }
  return '';
}

function getCausalSeverityScore(label){
  if (!label) return 0;
  const normalized = TextUtils.normalize(label);
  if (normalized.includes('red') || normalized.includes('roj')) return 4;
  if (normalized.includes('senal') || normalized.includes('signal')) return 3;
  if (normalized.includes('client')) return 2;
  if (normalized.includes('otro')) return 1;
  return 0;
}

function getCausalForZona(zona, ots, fmsMap){
  const counts = new Map();
  const addValue = (value, weight = 1) => {
    if (!value) return;
    const label = String(value).trim();
    if (!label) return;
    const key = TextUtils.normalize(label) || label;
    const current = counts.get(key) || { count: 0, severity: getCausalSeverityScore(label), label };
    current.count += weight;
    current.severity = Math.max(current.severity, getCausalSeverityScore(label));
    if (!current.label) current.label = label;
    counts.set(key, current);
  };

  if (Array.isArray(ots)) {
    ots.forEach(order => {
      const valor = pickFirstValue(order, ZONE_CAUSAL_KEYS);
      if (valor) {
        addValue(valor, 2);
      }
    });
  }

  const candidateZones = new Set();
  if (zona) candidateZones.add(zona);

  if (Array.isArray(ots)) {
    ots.forEach(order => {
      const meta = order.__meta || {};
      if (meta.zona) candidateZones.add(meta.zona);
      if (meta.zonaPrincipal) candidateZones.add(meta.zonaPrincipal);
      if (meta.zonaHFC) candidateZones.add(meta.zonaHFC);
      if (meta.zonaFTTH) candidateZones.add(meta.zonaFTTH);
      const zonaHFC = pickFirstValue(order, ORDER_FIELD_KEYS.zonaHFC);
      if (zonaHFC) candidateZones.add(zonaHFC);
      const zonaFTTH = pickFirstValue(order, ORDER_FIELD_KEYS.zonaFTTH);
      if (zonaFTTH) candidateZones.add(zonaFTTH);
    });
  }

  const normalizedFmsMap = new Map();
  if (fmsMap instanceof Map) {
    fmsMap.forEach((alarms, zoneKey) => {
      const baseKey = String(zoneKey || '').trim();
      if (!baseKey) return;
      if (!normalizedFmsMap.has(baseKey)) normalizedFmsMap.set(baseKey, alarms);
      const upperKey = baseKey.toUpperCase();
      if (!normalizedFmsMap.has(upperKey)) normalizedFmsMap.set(upperKey, alarms);
    });
  }

  candidateZones.forEach(candidate => {
    const base = String(candidate || '').trim();
    if (!base) return;
    const alarms = normalizedFmsMap.get(base) || normalizedFmsMap.get(base.toUpperCase());
    if (!alarms || !alarms.length) return;
    alarms.forEach(alarm => {
      const causal = alarm?.damageClassification || alarm?.damage || alarm?.incidentClassification || alarm?.type || '';
      addValue(causal, alarm?.isActive ? 2 : 1);
    });
  });

  if (!counts.size) return 'N/D';

  let selected = { label: 'N/D', count: 0, severity: -1 };
  counts.forEach(info => {
    if (
      info.count > selected.count ||
      (info.count === selected.count && info.severity > selected.severity)
    ) {
      selected = info;
    }
  });

  return selected.label || 'N/D';
}

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

  ensureOrderDeviceMeta(order);
  const meta = order.__meta || {};
  const device = meta.device || {};

  let fecha = pickFirstValue(order, ORDER_FIELD_KEYS.fecha);
  if (!fecha && meta.fecha){
    fecha = DateUtils.format(meta.fecha);
  }

  const zonaCAC = meta.zona ||
                   zoneInfo?.zona ||
                   meta.zonaHFC ||
                   meta.zonaFTTH ||
                   zoneInfo?.zonaHFC ||
                   zoneInfo?.zonaFTTH ||
                   pickFirstValue(order, ORDER_FIELD_KEYS.zonaHFC) ||
                   pickFirstValue(order, ORDER_FIELD_KEYS.zonaFTTH);

  const caso = meta.numeroCaso || pickFirstValue(order, ORDER_FIELD_KEYS.caso);
  const numeroOrden = pickFirstValue(order, ORDER_FIELD_KEYS.numeroOrden);
  const diagnosticoTecnico = extractDiagnosticoTecnico(order) || pickFirstValue(order, ORDER_FIELD_KEYS.diagnostico);
  const tipoTrabajo = pickFirstValue(order, ORDER_FIELD_KEYS.tipoTrabajo);

  let mac = pickFirstValue(order, ORDER_FIELD_KEYS.mac);
  if (!mac && Array.isArray(order.__meta?.dispositivos) && order.__meta.dispositivos.length) {
    const primaryDevice = order.__meta.dispositivos[0];
    mac = primaryDevice.macAddress || primaryDevice.mac || '';
  }
  const formattedMac = TextUtils.formatMac(mac);
  mac = device.mac || formattedMac || mac || '';
  const modelo = device.model || '';
  const tipoEquipo = device.type || device.technology || pickFirstValue(order, ORDER_FIELD_KEYS.tipo);
  const tipoDano = zoneInfo?.causalPrincipal || 'N/D';

  return {
    'Fecha': fecha || '',
    'Zona/CAC': zonaCAC || '',
    'Caso': caso || '',
    'Numero Orden': numeroOrden || '',
    'Diagnostico Tecnico': diagnosticoTecnico || '',
    'Tipo Trabajo': tipoTrabajo || '',
    'MAC': mac || '',
    'Modelo': modelo || '',
    'Tipo': tipoEquipo || '',
    'tipo_daño': tipoDano || 'N/D'
  };
}

function buildZoneExportRows(zoneData){
  if (!zoneData) return [];
  const source = zoneData.ordenesOriginales || zoneData.ordenes || [];
  const causalPrincipal = zoneData.causalPrincipal || getCausalForZona(zoneData.zona, source, dataProcessor?.fmsMap);
  zoneData.causalPrincipal = causalPrincipal;
  return source
    .map(order => buildOrderExportRow(order, zoneData))
    .filter(row => row && ZONE_EXPORT_HEADERS.some(header => (row[header] || '').toString().trim().length));
}

function createWorksheetFromRows(rows, headers){
  if (!rows || !rows.length) return null;
  const data = rows.map(row => headers.map(header => row[header] || ''));
  return XLSX.utils.aoa_to_sheet([headers, ...data]);
}

function sanitizeSheetName(name){
  const fallback = 'Hoja';
  if (!name) return fallback;
  const invalidChars = /[\/?*:[\]]/g;
  const cleaned = name.toString().replace(invalidChars, ' ').replace(/[\u0000-\u001f]/g, ' ').trim();
  const truncated = cleaned.substring(0, 31);
  return truncated || fallback;
}

function appendSheet(workbook, worksheet, desiredName, usedNames){
  if (!worksheet) return false;
  const base = sanitizeSheetName(desiredName);
  let name = base;
  let counter = 1;
  while (usedNames.has(name)){
    counter += 1;
    const suffix = `_${counter}`;
    const baseTrim = base.substring(0, Math.max(31 - suffix.length, 1));
    name = `${baseTrim}${suffix}`;
  }
  usedNames.add(name);
  XLSX.utils.book_append_sheet(workbook, worksheet, name);
  return true;
}

function exportEquiposGrupoExcel(zona, useFiltered = false){
  const source = useFiltered && window.equiposPorZona ? window.equiposPorZona : window.equiposPorZonaCompleto;
  
  if (!source) return toast('No hay datos de equipos');
  
  const arr = source.get(zona) || [];
  if (!arr.length) return toast('No hay equipos en esa zona');

  const wb = XLSX.utils.book_new();
  const exportRows = arr.map(e => ({
    Zona: e.zona || zona || '',
    Territorio: e.territorio || '',
    Caso: e.numCaso || '',
    Sistema: e.sistema || '',
    Serial: e.serialNumber || '',
    MAC: e.macAddress || '',
    Tipo: e.tipo || '',
    Marca: e.marca || '',
    Modelo: e.modelo || '',
    Descripcion: e.descripcion || '',
    Fecha: e.fecha || '',
    NumeroOrden: e.numeroOrden || ''
  }));
  const ws = XLSX.utils.json_to_sheet(exportRows);
  XLSX.utils.book_append_sheet(wb, ws, `Equipos_${zona || 'NA'}`);

  const filterInfo = useFiltered && (Filters.equipoModelo.length > 0 || Filters.equipoMarca || Filters.equipoTerritorio)
    ? `_filtrado`
    : '';
  
  XLSX.writeFile(wb, `Equipos_${zona || 'NA'}${filterInfo}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  toast(`✅ Exportados ${arr.length} equipos de ${zona}`);
}

function exportZonaExcel(zoneIdx) {
  const zonaData = window.currentAnalyzedZones?.[zoneIdx];
  if (!zonaData) return toast('No hay datos de la zona');

  const rows = buildZoneExportRows(zonaData);
  if (!rows.length) {
    toast('No hay órdenes para exportar en esta zona');
    return;
  }

  const wb = XLSX.utils.book_new();
  const usedNames = new Set();
  const sheet = createWorksheetFromRows(rows, ZONE_EXPORT_HEADERS);
  appendSheet(wb, sheet, `Zona_${zonaData.zona}`, usedNames);

  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Zona_${zonaData.zona}_${fecha}.xlsx`);
  toast(`✅ Exportada zona ${zonaData.zona} (${rows.length} órdenes)`);
}

function exportCMTSExcel(cmts) {
  const cmtsData = (window.currentCMTSData || []).find(c => c.cmts === cmts);
  if (!cmtsData) return toast('No hay datos del CMTS');
  
  const wb = XLSX.utils.book_new();
  const zonasFlat = cmtsData.zonas.map(z => ({
    Zona: z.zona,
    Tipo: z.tipo,
    Total_OTs: z.totalOTs,
    Ingreso_N: z.ingresoN,
    Ingreso_N1: z.ingresoN1,
    Estado_Nodo: z.nodoEstado
  }));
  
  const ws = XLSX.utils.json_to_sheet(zonasFlat);
  XLSX.utils.book_append_sheet(wb, ws, `CMTS_${cmts.slice(0, 20)}`);
  
  XLSX.writeFile(wb, `CMTS_${cmts.slice(0, 20)}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  toast(`✅ Exportado CMTS ${cmts}`);
}

let currentData = null;
let allZones = [];
let allCMTS = [];
let currentZone = null;
let selectedOrders = new Set();
let savedFiltersSnapshot = {};
let edificiosPanelListenerBound = false;

document.addEventListener('DOMContentLoaded', () => {
  savedFiltersSnapshot = FilterPersistence.load();
  setupEventListeners();
  restoreFiltersToControls(savedFiltersSnapshot);
});

function setupEventListeners() {
  on('fileConsolidado1', 'change', e => loadFile(e, 1));
  on('fileConsolidado2', 'change', e => loadFile(e, 2));
  on('fileNodos', 'change', e => loadFile(e, 3));
  on('fileFMS', 'change', e => loadFile(e, 4));

  const filterCatec = document.getElementById('filterCATEC');
  const filterExcludeCatec = document.getElementById('filterExcludeCATEC');
  on('filterCATEC', 'change', e => {
    if (e.target.checked && filterExcludeCatec) {
      filterExcludeCatec.checked = false;
    }
    applyFilters();
  });
  on('filterExcludeCATEC', 'change', e => {
    if (e.target.checked && filterCatec) {
      filterCatec.checked = false;
    }
    applyFilters();
  });

  on('showAllStates', 'change', applyFilters);

  const filterFTTH = document.getElementById('filterFTTH');
  const filterExcludeFTTH = document.getElementById('filterExcludeFTTH');
  on('filterFTTH', 'change', e => {
    if (e.target.checked && filterExcludeFTTH) {
      filterExcludeFTTH.checked = false;
    }
    applyFilters();
  });
  on('filterExcludeFTTH', 'change', e => {
    if (e.target.checked && filterFTTH) {
      filterFTTH.checked = false;
    }
    applyFilters();
  });
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

  const edificiosPanel = document.getElementById('edificiosPanel');
  if (edificiosPanel && !edificiosPanelListenerBound) {
    edificiosPanel.addEventListener('click', handleEdificiosPanelClick, false);
    edificiosPanelListenerBound = true;
  }

  document.addEventListener('click', handleZoneFilterOutsideClick);
}

async function loadFile(e, tipo) {
  const input = e.target;
  const file = input.files?.[0];
  if (!file) {
    return;
  }

  console.log(`📂 Archivo recibido (tipo ${tipo}): ${file.name || 'sin nombre'} • ${file.size || 0} bytes`);

  const statusEl = document.getElementById(`status${tipo}`);
  if (!statusEl) {
    console.warn(`No se encontró el indicador de estado para el archivo tipo ${tipo}`);
  } else {
    statusEl.textContent = 'Cargando...';
    statusEl.classList.remove('loaded');
  }

  try {
    const fileName = (file.name || '').toLowerCase();
    let result;

    const isCSV = fileName.endsWith('.csv');
    if (isCSV) {
      result = await dataProcessor.loadCSV(file);
    } else {
      result = await dataProcessor.loadExcel(file, tipo);
    }

    if (result.success) {
      console.log(`✅ Archivo procesado (tipo ${tipo}): ${file.name || 'sin nombre'} • ${result.rows} filas`);
      if (statusEl) {
        statusEl.textContent = `✓ ${result.rows} filas cargadas`;
        statusEl.classList.add('loaded');
      }

      const nombres = ['Consolidado 1', 'Consolidado 2', 'Nodos UP/DOWN', 'Alarmas FMS'];
      toast(`${nombres[tipo - 1]} cargado: ${result.rows} registros`);

      if (dataProcessor.consolidado1 || dataProcessor.consolidado2) {
        document.getElementById('mergeStatus').style.display = 'flex';
        processData();
      }
    } else {
      if (statusEl) {
        statusEl.textContent = '✗ Error';
      }
      toast(`Error al cargar archivo: ${result.error}`);
    }
  } catch (error) {
    console.error('Falla al procesar archivo:', error);
    if (statusEl) {
      statusEl.textContent = '✗ Error';
    }
    toast(error?.message || 'Error al procesar archivo');
  } finally {
    if (input) {
      input.value = '';
    }
  }
}

function processData() {
  const merged = dataProcessor.merge();
  if (!merged.length) return;

  merged.forEach(order => {
    enrichRowWithDeviceInfo(order);
    ensureOrderDeviceMeta(order);
  });

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

  const zoneNames = [...new Set(
    (currentData.zonas || [])
      .map(z => z.zona)
      .filter(Boolean)
  )];
  initializeZoneFilterOptions(zoneNames);
}

function applyFilters() {
  if (!currentData) return;
  
  const catecEl = document.getElementById('filterCATEC');
  const excludeCatecEl = document.getElementById('filterExcludeCATEC');

  Filters.catec = catecEl ? catecEl.checked : false;
  Filters.excludeCatec = excludeCatecEl ? excludeCatecEl.checked : false;
  Filters.showAllStates = document.getElementById('showAllStates').checked;
  Filters.ftth = document.getElementById('filterFTTH').checked;
  Filters.excludeFTTH = document.getElementById('filterExcludeFTTH').checked;
  Filters.nodoEstado = document.getElementById('filterNodoEstado').value;
  Filters.cmts = document.getElementById('filterCMTS').value;

  const daysValue = parseInt(document.getElementById('daysWindow').value, 10);
  Filters.days = Number.isFinite(daysValue) ? daysValue : CONFIG.defaultDays;

  Filters.territorio = document.getElementById('filterTerritorio').value;
  Filters.sistema = document.getElementById('filterSistema').value;
  Filters.alarma = document.getElementById('filterAlarma').value;
  Filters.quickSearch = document.getElementById('quickSearch').value;
  Filters.ordenarPorIngreso = document.getElementById('ordenarPorIngreso').value;
  Filters.selectedZonas = Array.from(zoneFilterState.selected);

  const filtered = Filters.apply(currentData.ordenes);
  
  const zones = dataProcessor.processZones(filtered);
  let analyzed = dataProcessor.analyzeZones(zones, Filters.days);
  
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

  const edificiosPanelEl = document.getElementById('edificiosPanel');
  if (edificiosPanelEl) {
    edificiosPanelEl.innerHTML = UIRenderer.renderEdificios(filtered);
    const botonesVer = edificiosPanelEl.querySelectorAll('.btn-ver-edificio');
    console.log('🔍 Botones "Ver" en edificios:', botonesVer.length); // verificación rápida del renderizado
  }

  const fmsPanelEl = document.getElementById('fmsPanel');
  if (fmsPanelEl) {
    const fmsMap = dataProcessor?.fmsMap instanceof Map ? dataProcessor.fmsMap : new Map();
    console.assert(fmsMap instanceof Map && fmsMap.size > 0, 'fmsMap vacío');
    fmsPanelEl.innerHTML = FMSPanel.render(filtered, fmsMap);
  }

  document.getElementById('equiposPanel').innerHTML = UIRenderer.renderEquipos(filtered);
  persistFilters();
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
  const quickSearchInput = document.getElementById('quickSearch');

  if (statType === 'territoriosCriticos') {
    if (!window.territoriosData) {
      toast('❌ No hay datos de territorios disponibles');
      return;
    }

    const territoriosCriticos = window.territoriosData.filter(t => t.esCritico);
    if (territoriosCriticos.length === 0) {
      toast('✅ No hay territorios críticos en este momento');
      return;
    }

    showTerritoriosCriticosModal(territoriosCriticos);
    return;
  }

  if (quickSearchInput) {
    quickSearchInput.value = '';
  }
  Filters.quickSearch = '';

  resetZoneFilterState({ apply: false });

  switch (statType) {
    case 'total':
    case 'zonas':
      toast('🔄 Mostrando todas las zonas (filtros avanzados conservados)');
      break;

    case 'ftth': {
      const ftthEl = document.getElementById('filterFTTH');
      const excludeEl = document.getElementById('filterExcludeFTTH');
      if (ftthEl) ftthEl.checked = true;
      if (excludeEl) excludeEl.checked = false;
      toast('🔌 Mostrando solo zonas FTTH (filtros avanzados conservados)');
      break;
    }

    case 'alarmas': {
      const alarmaEl = document.getElementById('filterAlarma');
      if (alarmaEl) alarmaEl.value = 'con-alarma';
      toast('🚨 Mostrando zonas con alarmas activas (filtros avanzados conservados)');
      break;
    }

    case 'nodosCriticos': {
      const nodoEl = document.getElementById('filterNodoEstado');
      if (nodoEl) nodoEl.value = 'critical';
      toast('🟥 Mostrando zonas con NODOS CRÍTICOS (filtros avanzados conservados)');
      break;
    }

    default:
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
  document.querySelectorAll('.tab-panel').forEach(panel => {
    panel.classList.remove('active');
    panel.style.display = 'none';
  });

  const targetBtn = document.querySelector(`[data-tab="${tabName}"]`);
  if (targetBtn) {
    targetBtn.classList.add('active');
  }

  const targetPanel = document.getElementById(`${tabName}Panel`);
  if (targetPanel) {
    targetPanel.classList.add('active');
    targetPanel.style.display = '';
  }
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

function handleEdificiosPanelClick(event) {
  const trigger = event.target.closest('.btn-ver-edificio');
  if (!trigger) return;
  event.preventDefault();

  const indexAttr = trigger.getAttribute('data-index');
  const direccionAttr = trigger.getAttribute('data-direccion');
  if (indexAttr == null && !direccionAttr) {
    console.warn('No se pudo determinar el identificador del edificio', trigger);
    toast('❌ No se pudo abrir el edificio seleccionado');
    return;
  }

  showEdificioModal(indexAttr ?? direccionAttr);
}

function showEdificioModal(identifier) {
  const data = Array.isArray(window.edificiosData) ? window.edificiosData : [];
  if (!data.length) {
    console.warn('No hay datos de edificios disponibles para abrir el detalle');
    toast('❌ No hay información de edificios disponible');
    return;
  }

  let edificio = null;
  if (identifier !== undefined && identifier !== null) {
    const idx = Number(identifier);
    if (!Number.isNaN(idx) && data[idx]) {
      edificio = data[idx];
    }
  }

  if (!edificio && typeof identifier === 'string') {
    const normalizedTarget = TextUtils.normalize(identifier);
    edificio = data.find(item => TextUtils.normalize(item.direccion) === normalizedTarget) || null;
  }

  if (!edificio) {
    console.warn('No se encontró información del edificio solicitado', identifier);
    toast('❌ No se encontraron datos del edificio seleccionado');
    return;
  }

  console.log(`open edificio ${identifier}`); // traceo de apertura solicitado

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

function showEdificioDetail(idx) {
  showEdificioModal(idx);
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

  let resumenDispositivo = null;
  for (const orden of ordenes) {
    const device = ensureOrderDeviceMeta(orden);
    if (device && (device.brand || device.technology || device.model || device.mac || device.type)) {
      resumenDispositivo = device;
      break;
    }
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

  if (resumenDispositivo) {
    const valores = {
      'Marca': resumenDispositivo.brand,
      'Tecnología': resumenDispositivo.technology,
      'Modelo': resumenDispositivo.model,
      'MAC': resumenDispositivo.mac,
      'Tipo': resumenDispositivo.type || resumenDispositivo.technology
    };
    const infoHtml = Object.entries(valores).map(([label, value]) => {
      const safeValue = value ? escapeHtml(value) : '—';
      return `<div style="min-width: 140px;"><strong>${label}:</strong> ${safeValue}</div>`;
    }).join('');

    html += `<div style="margin-bottom: 16px; padding: 12px; background: var(--bg-tertiary); border-radius: 8px;">
      <div style="font-weight: 600; margin-bottom: 8px;">Resumen del equipo</div>
      <div style="display: flex; flex-wrap: wrap; gap: 12px; font-size: 0.85rem;">${infoHtml}</div>
    </div>`;
  } else {
    html += `<div style="margin-bottom: 16px; padding: 12px; background: var(--bg-tertiary); border-radius: 8px; font-size: 0.85rem; color: var(--text-secondary);">
      Sin información de dispositivos en las órdenes visibles.
    </div>`;
  }

  const diagnosticosTecnicos = Array.isArray(currentZone.diagnosticos) ? currentZone.diagnosticos.filter(Boolean) : [];
  if (diagnosticosTecnicos.length) {
    const chips = diagnosticosTecnicos.slice(0, 6).map(d => {
      const raw = String(d);
      const title = escapeHtml(raw);
      const truncated = raw.length > 80 ? `${raw.slice(0, 77)}…` : raw;
      const label = escapeHtml(truncated);
      return `<span class="diagnostico-tag" title="${title}">${label}</span>`;
    }).join('');
    const extra = diagnosticosTecnicos.length > 6
      ? `<span class="diagnostico-tag diagnostico-tag--more">+${diagnosticosTecnicos.length - 6}</span>`
      : '';
    html += `<div class="diagnostico-summary">
      <div class="diagnostico-summary__title">Diagnóstico técnico en la zona</div>
      <div class="diagnostico-tags">${chips}${extra}</div>
    </div>`;
  } else {
    html += `<div class="diagnostico-summary diagnostico-summary--empty">Diagnóstico técnico: sin datos registrados en las órdenes de la zona.</div>`;
  }

  html += '<div class="table-container"><div style="max-height: 400px; overflow-y: auto;"><table class="detail-table"><thead><tr>';
  html += '<th style="position: sticky; top: 0; z-index: 10;"><input type="checkbox" id="selectAllCheckbox" onchange="toggleSelectAll()"></th>';
  
  const allCols = new Set();
  ordenes.forEach(o => {
    Object.keys(o).forEach(k => {
      if (!k.startsWith('_')) allCols.add(k);
    });
  });
  
  const columns = Array.from(allCols);
  const extraDeviceColumns = ['MAC', 'Modelo', 'Tipo'];
  extraDeviceColumns.forEach(col => {
    if (!columns.includes(col)) {
      columns.push(col);
    }
  });

  columns.forEach(col => {
    html += `<th style="position: sticky; top: 0; z-index: 10; white-space: nowrap;">${col}</th>`;
  });
  
  html += '</tr></thead><tbody>';
  
  ordenes.forEach((o, idx) => {
    const cita = o['Número de cita'] || `row_${idx}`;
    const checked = selectedOrders.has(cita) ? 'checked' : '';
    const device = ensureOrderDeviceMeta(o) || {};
    const rawMac = pickFirstValue(o, ORDER_FIELD_KEYS.mac);
    let fallbackMac = rawMac;
    const dispositivoPrincipal = Array.isArray(o.__meta?.dispositivos) && o.__meta.dispositivos.length
      ? o.__meta.dispositivos[0]
      : null;
    if (!fallbackMac && dispositivoPrincipal) {
      fallbackMac = dispositivoPrincipal.macAddress || dispositivoPrincipal.mac || '';
    }
    fallbackMac = TextUtils.formatMac(fallbackMac) || '';
    const fallbackTipo = pickFirstValue(o, ORDER_FIELD_KEYS.tipo);
    const deviceData = {
      MAC: device.mac || fallbackMac,
      Modelo: device.model || '',
      Tipo: device.type || device.technology || fallbackTipo || ''
    };
    html += `<tr>`;
    html += `<td><input type="checkbox" class="order-checkbox" data-cita="${cita}" ${checked} onchange="updateSelection()"></td>`;

    columns.forEach(col => {
      let value;

      if (extraDeviceColumns.includes(col)) {
        value = deviceData[col] || '';
      } else {
        value = o[col] || '';
      }

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

function debounce(func, wait) {
  let timeout;
  return function (...args) {
    clearTimeout(timeout);
    timeout = setTimeout(() => func.apply(this, args), wait);
  };
}

console.log('✅ Panel v5.0 COMPLETO inicializado');
console.log('🎯 FUNCIONALIDADES:');
console.log('   • Territorios Críticos con agrupación Flex');
console.log('   • Filtros múltiples de equipos (modelo/marca/territorio)');
console.log('   • Stats clickeables con modales interactivos');
console.log('   • Exportaciones completas (Zonas/CMTS/Equipos/BEFAN)');
console.log('   • Planillas funcionales en nueva ventana');
console.log('📊 Estados activos:', CONFIG.estadosPermitidos);
