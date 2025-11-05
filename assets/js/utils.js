console.log('🚀 Panel v4.9 MEJORADO - Filtros de equipos + Ordenamiento por ingresos');

(function(){
  const CONFIG = {
    defaultDays: 3,
    codigo_befan: 'FR461',
    estadosPermitidos: [
      'NUEVA',
      'EN PROGRESO',
      'PENDIENTE DE ACCION',
      'PROGRAMADA',
      'PENDIENTE DE CONTACTO',
      'EN ESPERA DE EJECUCION'
    ],
    estadosOcultosPorDefecto: ['CANCELADA','CERRADA']
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
      try {
        if (!jsonStr) return [];
        try {
          const parsed = JSON.parse(jsonStr);
          const arr = parsed?.openResp || parsed?.open_response || parsed?.data || [];
          if (Array.isArray(arr)) {
            return arr.map(d => ({
              category: d.category || '',
              description: d.description || '',
              model: d.model || '',
              serialNumber: d.serialNumber || d.serial || '',
              macAddress: d.macAddress || d.mac || '',
              type: d.type || ''
            }));
          }
        } catch {
          // fallback handled below
        }
        const get = (f) => {
          const re = new RegExp(`"${f}"\\s*:\\s*"([^"]*)"`, 'i');
          const m = String(jsonStr).match(re);
          return m ? m[1] : '';
        };
        if (/macAddress|serialNumber|model|category|description/i.test(jsonStr)){
          return [{
            category: get('category'),
            description: get('description'),
            model: get('model'),
            serialNumber: get('serialNumber'),
            macAddress: get('macAddress'),
            type: get('type')
          }];
        }
      } catch {}
      return [];
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
})();
