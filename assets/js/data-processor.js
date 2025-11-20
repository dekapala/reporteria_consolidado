/**
 * DataProcessor v5.0 CORREGIDO
 * Compatible con app_fixed.js
 * 
 * Mejoras aplicadas:
 * - readFileAsUint8Array con fallback
 * - Plan B para headers (fila 0)
 * - Asignación de __meta.dispositivos
 * - Mejor logging y manejo de errores
 */

(function(){
  if (!window.XLSX) {
    console.error('❌ XLSX no está disponible. Verifica el <script> del CDN antes de data-processor.js');
    return;
  }

  // === FUNCIONES AUXILIARES ===
  
  // Helper para leer archivos con fallback
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

  // Quitar acentos
  function stripAccents(s = '') {
    return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  // Buscar columna de dispositivos
  function findDispositivosColumn(rowObj) {
    const keys = Object.keys(rowObj || {});
    for (const k of keys) {
      const nk = stripAccents(k.toLowerCase());
      if (nk.includes('informacion') && nk.includes('dispositivo')) return k;
    }
    return rowObj['Informacion Dispositivos'] ? 'Informacion Dispositivos'
         : rowObj['Información Dispositivos'] ? 'Información Dispositivos'
         : null;
  }

  // Parser de dispositivos JSON
  function parseDispositivosJSON(jsonStr) {
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
      } catch { }
      
      const get = (f) => {
        const re = new RegExp(`"${f}"\\s*:\\s*"([^"]*)"`, 'i');
        const m = String(jsonStr).match(re);
        return m ? m[1] : '';
      };

      if (/macAddress|serialNumber|model|category|description/i.test(jsonStr)) {
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
  }

  // Parsear fecha
  function parseDate(str) {
    if (!str) return null;
    const match = String(str).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (match) return new Date(match[3], match[2] - 1, match[1]);
    return new Date(str);
  }

  // Días entre fechas
  function daysBetween(date1, date2) {
    const oneDay = 24 * 60 * 60 * 1000;
    return Math.round(Math.abs((date1 - date2) / oneDay));
  }

  // Formatear fecha
  function formatDate(date) {
    if (!date) return '';
    const d = String(date.getDate()).padStart(2, '0');
    const m = String(date.getMonth() + 1).padStart(2, '0');
    return `${d}/${m}/${date.getFullYear()}`;
  }

  // Clave de día
  function toDayKey(date) {
    const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    return d.getTime();
  }

  // Normalizar territorio
  function normalizarTerritorio(territorio) {
    if (!territorio) return '';
    return territorio
      .replace(/\s*flex\s*\d+/gi, ' Flex')
      .replace(/\s{2,}/g, ' ')
      .trim();
  }

  function extraerFlex(territorio) {
    if (!territorio) return null;
    const match = territorio.match(/flex\s*(\d+)/i);
    return match ? parseInt(match[1]) : null;
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
      console.log(`📊 Procesando Excel tipo ${tipo}: ${file.name}`);
      
      const data = await readFileAsUint8Array(file);

      const wb = XLSX.read(data, {
        type: 'array',
        cellDates: false,
        cellText: false,
        cellFormula: false,
        raw: false
      });

      if (!wb || !wb.SheetNames || !wb.SheetNames.length) {
        return {success: false, error: 'Libro sin hojas'};
      }
      
      const ws = wb.Sheets[wb.SheetNames[0]];
      if (!ws || !ws['!ref']) {
        const rangeGuess = XLSX.utils.decode_range(XLSX.utils.encode_range({s:{r:0,c:0}, e:{r:9999,c:99}}));
        ws['!ref'] = ws['!ref'] || XLSX.utils.encode_range(rangeGuess);
      }
      if (!ws['!ref']) {
        return {success: false, error: 'Hoja sin rango (!ref) detectable'};
      }
      
      const range = XLSX.utils.decode_range(ws['!ref']);
      
      // Detector de headers con PLAN B
      let headerRow = 0;
      let headerFound = false;
      
      for (let R = 0; R <= 20; R++) {
        for (let C = 0; C <= 15; C++) {
          const cellAddr = XLSX.utils.encode_cell({r: R, c: C});
          const cell = ws[cellAddr];
          if (cell && cell.v) {
            const cellValue = String(cell.v).toLowerCase();
            if (cellValue.includes('zona') && cellValue.includes('tecnica')) {
              headerRow = R;
              console.log(`✅ Headers detectados en fila ${R + 1} (columna ${C}): "${cell.v}"`);
              headerFound = true;
              break;
            }
            if (cellValue.includes('numero') && (cellValue.includes('caso') || cellValue.includes('cuenta'))) {
              headerRow = R;
              console.log(`✅ Headers detectados en fila ${R + 1} (columna ${C}): "${cell.v}"`);
              headerFound = true;
              break;
            }
          }
        }
        if (headerFound) break;
      }
      
      // PLAN B: Si no encuentra headers, usar fila 0
      if (!headerFound) {
        headerRow = 0;
        console.warn(`⚠️ No se detectaron headers conocidos, usando fila 1 como headers (Plan B)`);
      }
      
      console.log(`📋 Usando fila ${headerRow + 1} como headers`);
      
      range.s.r = headerRow;
      ws['!ref'] = XLSX.utils.encode_range(range);
      
      const jsonData = XLSX.utils.sheet_to_json(ws, { 
        defval: '', 
        raw: false, 
        dateNF: 'dd/mm/yyyy' 
      });
      
      // NUEVO: Procesar dispositivos y asignar __meta a cada orden
      if (tipo === 1 || tipo === 2) {
        jsonData.forEach(order => {
          const colInfo = findDispositivosColumn(order);
          const infoDispositivos = colInfo ? order[colInfo] : '';
          const dispositivos = parseDispositivosJSON(infoDispositivos);
          
          order.__meta = {
            dispositivos: dispositivos,
            zonaHFC: order['Zona Tecnica HFC'] || '',
            zonaFTTH: order['Zona Tecnica FTTH'] || '',
            territorio: order['Territorio de servicio: Nombre'] || '',
            numeroCaso: order['Número del caso'] || order['Caso Externo'] || '',
            fecha: parseDate(order['Fecha de creación'] || order['Fecha/Hora de apertura'] || order['Fecha de inicio'])
          };
        });
        console.log(`📱 Dispositivos parseados para ${jsonData.length} órdenes`);
      }
      
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
      console.error(`❌ Error loading Excel (tipo ${tipo}):`, e);
      return {success: false, error: e.message};
    }
  }
  
  async loadCSV(file) {
    try {
      console.log(`📊 Procesando CSV: ${file.name}`);
      
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
      console.error('❌ Error loading CSV:', e);
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
    
    console.log('📊 Procesando alarmas FMS...');
    
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

    console.log(`📊 Analizando zonas con ventana de ${daysWindow} días`);

    return zones.map(z => {
      const seenOTperDay = new Map();
      const ordenesPerDay = new Map();
      let ordenesEnVentana = 0;

      const diagnosticos = new Set();

      z.ordenes.forEach(o => {
        const fecha = o['Fecha de creación'] || o['Fecha/Hora de apertura'] || o['Fecha de inicio'];
        const dt = parseDate(fecha);
        if (!dt) return;

        const daysAgo = daysBetween(today, dt);

        if (daysAgo <= daysWindow) {
          ordenesEnVentana++;

          const dk = toDayKey(dt);
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

      const todayKey = toDayKey(today);
      const yesterdayKey = toDayKey(new Date(today.getTime() - 86400000));

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
        const dk = toDayKey(date);
        last7Days.push(formatDate(date).slice(0, 5));
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
      
      const territorioNormalizado = normalizarTerritorio(territorioOriginal);
      const numeroFlex = extraerFlex(territorioOriginal);
      
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

// Exponer globalmente
const dataProcessor = new DataProcessor();
window.DataProcessor = DataProcessor;
window.dataProcessor = dataProcessor;

console.log('✅ DataProcessor v5.0 cargado correctamente');

})();
