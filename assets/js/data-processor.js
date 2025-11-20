(function(){
  if (!window.XLSX) {
    console.error('❌ XLSX no está disponible. Verifica el <script> del CDN antes de data-processor.js');
    return;
  }

  const stripAccentsLocal = (s = '') => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

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
      const normalized = stripAccentsLocal(k).toLowerCase();
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

  normalizeId(value) {
    if (value === undefined || value === null) return '';
    const str = String(value).trim();
    if (!str) return '';

    // Unifica IDs que vienen como números o con sufijo .0 desde Excel
    if (/^-?\d+(\.0+)?$/.test(str)) {
      return str.replace(/\.0+$/, '');
    }

    return str;
  }

  buildRowSignature(row) {
    if (!row || typeof row !== 'object') return '';
    const keys = Object.keys(row)
      .filter(k => row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '')
      .sort();

    if (!keys.length) return '';

    return keys
      .map(k => `${stripAccentsLocal(k).toLowerCase()}=${String(row[k]).trim()}`)
      .join('|');
  }
  
  async loadExcel(file, tipo) {
    try {
      const data = typeof readFileAsUint8Array === 'function'
        ? await readFileAsUint8Array(file)
        : new Uint8Array(await file.arrayBuffer());

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
      if (!ws || !ws['!ref']) {
        // intento recuperar rango inferido
        const rangeGuess = XLSX.utils.decode_range(XLSX.utils.encode_range({s:{r:0,c:0}, e:{r:9999,c:99}}));
        ws['!ref'] = ws['!ref'] || XLSX.utils.encode_range(rangeGuess);
      }
      if (!ws['!ref']) {
        return {success:false, error:'Hoja sin rango (!ref) detectable'};
      }
      const range = XLSX.utils.decode_range(ws['!ref']);
      
      let headerRow = 0;
      let foundHeaders = false;
      for (let R = 0; R <= 20; R++) {
        let hasContent = false;
        for (let C = 0; C <= 15; C++) {
          const cellAddr = XLSX.utils.encode_cell({r: R, c: C});
          const cell = ws[cellAddr];
          if (cell && cell.v) {
            const cellValue = String(cell.v).toLowerCase();
            if (cellValue.includes('zona') && cellValue.includes('tecnica')) {
              headerRow = R;
              foundHeaders = true;
              console.log(`✅ Headers detectados en fila ${R + 1} (columna ${C}): "${cell.v}"`);
              hasContent = true;
              break;
            }
            if (cellValue.includes('numero') && (cellValue.includes('caso') || cellValue.includes('cuenta'))) {
              headerRow = R;
              foundHeaders = true;
              console.log(`✅ Headers detectados en fila ${R + 1} (columna ${C}): "${cell.v}"`);
              hasContent = true;
              break;
            }
          }
        }
        if (hasContent) break;
      }
      
      if (!foundHeaders) {
        console.warn('⚠️ No se detectaron headers explícitos; usando fila 1 como fallback');
        headerRow = 0;
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
      const bytes = typeof readFileAsUint8Array === 'function'
        ? await readFileAsUint8Array(file)
        : new Uint8Array(await file.arrayBuffer());
      let text = '';
      if (typeof decodeText === 'function') {
        text = decodeText(bytes);
      } else if (typeof TextDecoder !== 'undefined') {
        text = new TextDecoder('utf-8').decode(bytes);
      } else {
        text = Array.from(bytes).map(b => String.fromCharCode(b)).join('');
      }
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
    let dedupCount = 0;

    const mergeInto = (target, source) => {
      Object.keys(source).forEach(key => {
        const incoming = source[key];
        const current = target[key];
        if ((current === undefined || current === null || current === '') && incoming !== undefined && incoming !== null && String(incoming).trim() !== '') {
          target[key] = incoming;
        }
      });
    };

    const upsertRow = (rows, tag) => {
      if (!Array.isArray(rows)) return;

      rows.forEach((r, idx) => {
        const orderId = this.normalizeId(this.getOrderId(r));
        let key = orderId;

        if (!key) {
          const signature = this.buildRowSignature(r);
          if (signature) {
            key = `sig:${signature}`;
          } else {
            missingIdCount++;
            key = `${tag}-auto-${idx + 1}`;
          }
        }

        if (map.has(key)) {
          mergeInto(map.get(key), r);
          map.get(key)._merged = true;
          dedupCount++;
        } else {
          map.set(key, { ...r, _merged: false });
        }
      });
    };

    upsertRow(this.consolidado1, 'c1');
    upsertRow(this.consolidado2, 'c2');

    const result = Array.from(map.values());
    const suffix = missingIdCount > 0 ? ` • IDs asignados por fallback: ${missingIdCount}` : '';
    const dedupSuffix = dedupCount > 0 ? ` • Duplicados fusionados: ${dedupCount}` : '';
    console.log(`✅ Total órdenes únicas (deduplicadas): ${result.length}${suffix}${dedupSuffix}`);

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

const dataProcessor = new DataProcessor();
window.DataProcessor = DataProcessor;
window.dataProcessor = dataProcessor;
})();
