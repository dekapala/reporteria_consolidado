console.log('🚀 Panel v5.0 COMPLETO - Sparklines normalizadas + histórico del modal unificado');

(function(){
  const CONFIG = {
    defaultDays: 3,
    codigo_befan: 'FR461',
    
    estadosPermitidosEstado1: [
      'DERIVADA',
      'EN ESPERA DE EJECUCION',
      'NUEVA',
      'PENDIENTE DE CONTACTO',
      'PENDIENTE EVENTO MASIVO',
      ''
    ],
    
    estadosPermitidosEstado2: [
      'CERRADA',
      'EN PROGRESO',
      'NUEVA',
      'PENDIENTE DE ACCION',
      'PROGRAMADA'
    ],
    
    estadosOcultosPorDefecto: ['CANCELADA']
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

  function stripAccents(s=''){
    return s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');
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
      if (!jsonStr) return [];

      const resultados = [];
      const firmas = new Set();

      const normalizarClave = (clave) => stripAccents(String(clave || '')).toLowerCase().replace(/[^a-z0-9]/g, '');
      const normalizarValor = (valor) => {
        if (valor === null || valor === undefined) return '';
        return String(valor).trim();
      };

      const agregar = (entrada = {}) => {
        const limpio = {
          category: normalizarValor(entrada.category),
          description: normalizarValor(entrada.description),
          model: normalizarValor(entrada.model),
          serialNumber: normalizarValor(entrada.serialNumber),
          macAddress: normalizarValor(entrada.macAddress),
          type: normalizarValor(entrada.type)
        };

        const tieneDatosClave = limpio.macAddress || limpio.serialNumber || limpio.model || limpio.description;
        if (!tieneDatosClave) return;

        const firma = `${limpio.macAddress}||${limpio.serialNumber}||${limpio.model}||${limpio.description}`;
        if (firmas.has(firma)) return;
        firmas.add(firma);
        resultados.push(limpio);
      };

      const obtenerDesdeObjeto = (obj) => {
        if (!obj || typeof obj !== 'object') return;

        const mapa = {};
        Object.keys(obj).forEach(key => {
          mapa[normalizarClave(key)] = obj[key];
        });

        const obtener = (...variantes) => {
          for (const variante of variantes) {
            const clave = normalizarClave(variante);
            if (Object.prototype.hasOwnProperty.call(mapa, clave)) {
              const valor = normalizarValor(mapa[clave]);
              if (valor) return valor;
            }
          }
          return '';
        };

        const entrada = {
          category: obtener('category', 'categoria'),
          description: obtener('description', 'descripcion', 'detalle'),
          model: obtener('model', 'modelo'),
          serialNumber: obtener('serialnumber', 'serial', 'serie', 'numerodeserie', 'nroserie', 'sn', 'serialid'),
          macAddress: obtener('macaddress', 'mac', 'macaddr', 'direccionmac', 'direccionmacaddress', 'macid', 'deviceid'),
          type: obtener('type', 'tipo')
        };

        agregar(entrada);
      };

      const recorrer = (valor) => {
        if (!valor) return;
        if (Array.isArray(valor)) {
          valor.forEach(recorrer);
          return;
        }
        if (typeof valor === 'object') {
          obtenerDesdeObjeto(valor);
          Object.values(valor).forEach(recorrer);
        }
      };

      try {
        const parsed = JSON.parse(jsonStr);
        recorrer(parsed);
      } catch (err) {
        // Ignorado, se usará la extracción mediante expresiones regulares
      }

      if (!resultados.length) {
        const texto = String(jsonStr);

        const macRegex = /(?:^|[\s,;\|])(?:mac|mac address|direcci[oó]n\s*mac|mac\s*id)\s*[:=\-]?\s*([0-9a-f]{2}(?:[:\-]?[0-9a-f]{2}){5,})/gi;
        const serialRegex = /(?:^|[\s,;\|])(?:sn|serial|serie|nro\s*serie|numero\s*serie|n[uú]mero\s*serie)\s*[:=\-]?\s*([a-z0-9\-]{4,})/gi;

        let match;
        const macs = new Set();
        while ((match = macRegex.exec(texto))) {
          const mac = match[1]?.replace(/[^0-9a-f:]/gi, '').toUpperCase();
          if (mac && !macs.has(mac)) {
            macs.add(mac);
            agregar({macAddress: mac});
          }
        }

        const seriales = new Set();
        while ((match = serialRegex.exec(texto))) {
          const serial = normalizarValor(match[1]);
          if (serial && !seriales.has(serial)) {
            seriales.add(serial);
            agregar({serialNumber: serial});
          }
        }
      }

      return resultados;
    },
    detectarSistema(numCaso) {
      const numStr = String(numCaso || '').trim();
      if (numStr.startsWith('8')) return 'OPEN';
      if (numStr.startsWith('3')) return 'FAN';
      return '';
    }
  };

  function toast(msg) {
    const t = document.createElement('div');
    t.className = 'toast';
    t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(() => t.remove(), 4000);
  }

  function debounce(func, wait) {
    let timeout;
    return function(...args) {
      const later = () => {
        timeout = null;
        func.apply(this, args);
      };
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
    };
  }

  function openPlanillasNewTab() {
    document.getElementById('utilitiesMenu').classList.remove('show');
    window.open('planillas.html', '_blank');
    toast('📋 Abriendo herramientas de Planillas...');
  }

  window.CONFIG = CONFIG;
  window.FMS_TIPOS = FMS_TIPOS;
  window.stripAccents = stripAccents;
  window.findDispositivosColumn = findDispositivosColumn;
  window.DateUtils = DateUtils;
  window.NumberUtils = NumberUtils;
  window.TextUtils = TextUtils;
  window.toast = toast;
  window.debounce = debounce;
  window.openPlanillasNewTab = openPlanillasNewTab;

  console.log('✅ utils.js cargado correctamente');
  console.log('📊 Estados Estado.1:', CONFIG.estadosPermitidosEstado1);
  console.log('📊 Estados Estado.2:', CONFIG.estadosPermitidosEstado2);
})();
