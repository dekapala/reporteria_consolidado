(function(){
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
      const data = new Uint8Array(await file.arrayBuffer());

      const wb = XLSX.read(data, {
        type: 'array',
        cellDates: false,
        cellText: false,
        cellFormula: false,
        raw: false
      });

      const ws = wb.Sheets[wb.SheetNames[0]];
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
      console.log(`📋 Primeras 10 columnas:`, jsonData.length > 0 ? Object.keys(jsonData[0]).slice(0, 10) : []);

      if (jsonData.length > 0) {
        const zonaCols = Object.keys(jsonData[0]).filter(k => k.toLowerCase().includes('zona'));
        if (zonaCols.length > 0) {
          console.log(`✅ Columnas de ZONA encontradas:`, zonaCols);
          zonaCols.forEach(col => {
            const ejemplo = jsonData[0][col];
            console.log(`   "${col}" = "${ejemplo}"`);
          });
        }
      }

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
          map.set(cita, {...r, _merged: false});
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
          map.set(cita, {...r, _merged: false});
        }
      }
    });

    const result = Array.from(map.values());
    console.log(`✅ Total órdenes únicas (deduplicadas): ${result.length}`);

    return result;
  }

  prepareOrders(rows) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    return rows.map(row => {
      const order = {...row};
      const meta = {};

      const fechaStr = row['Fecha de creación'] || row['Fecha/Hora de apertura'] || row['Fecha de inicio'];
      const fecha = DateUtils.parse(fechaStr);
      meta.fecha = fecha || null;
      meta.dayKey = fecha ? DateUtils.toDayKey(fecha) : null;
      meta.daysFromToday = fecha ? DateUtils.daysBetween(today, fecha) : Number.POSITIVE_INFINITY;

      const estado1Raw = (row['Estado.1'] ?? row['Estado'] ?? '').toString().toUpperCase().trim();
      const estado2Raw = (row['Estado.2'] ?? row['Estado final'] ?? '').toString().toUpperCase().trim();

      meta.estado1 = estado1Raw;
      meta.estado2 = estado2Raw;
      meta.estado = estado1Raw || estado2Raw;

      const estado1Valido = CONFIG.estadosPermitidosEstado1.includes(estado1Raw);
      const estado2Valido = CONFIG.estadosPermitidosEstado2.includes(estado2Raw);
      meta.estadoValido = estado1Valido && estado2Valido;
      const estadoRaw = (row['Estado.1'] || row['Estado'] || row['Estado.2'] || '').toString().toUpperCase().trim();
      meta.estado = estadoRaw;
      meta.estadoValido = estadoRaw && (!CONFIG.estadosOcultosPorDefecto.includes(estadoRaw));

      const {zonaPrincipal, tipo} = this.getZonaPrincipal(row);
      meta.zonaPrincipal = zonaPrincipal;
      meta.zonaTipo = tipo;
      meta.zonaHFC = row['Zona Tecnica HFC'] || row['Zona Tecnica'] || '';
      meta.zonaFTTH = row['Zona Tecnica FTTH'] || '';
      meta.esFTTH = /9\d{2}/.test(meta.zonaHFC);

      meta.territorio = row['Territorio de servicio: Nombre'] || row['Territorio'] || '';

      const numCaso = row['Número del caso'] || row['Caso Externo'] || row['External Case Id'] || '';
      meta.numeroCaso = numCaso;
      meta.sistema = TextUtils.detectarSistema(numCaso);

      const tipoTrabajo = row['Tipo de trabajo: Nombre de tipo de trabajo'] || '';
      meta.tipoTrabajo = tipoTrabajo.toUpperCase();

      const searchableParts = [
        meta.zonaHFC,
        meta.zonaFTTH,
        row['Calle'] || '',
        row['Caso Externo'] || '',
        row['External Case Id'] || '',
        row['Número del caso'] || ''
      ];
      meta.searchable = searchableParts.join(' ');

      const dispositivosColumn = findDispositivosColumn(row);
      meta.dispositivosColumna = dispositivosColumn;
      meta.dispositivos = dispositivosColumn ? TextUtils.parseDispositivosJSON(row[dispositivosColumn]) : [];

      order.__meta = meta;
      return order;
    });
  }

  processZones(rows) {
    const zoneGroups = new Map();

    rows.forEach(r => {
      const meta = r.__meta || {};
      const zonaPrincipal = meta.zonaPrincipal;
      const tipo = meta.zonaTipo;
      if (!zonaPrincipal) return;

      if (!zoneGroups.has(zonaPrincipal)) {
        zoneGroups.set(zonaPrincipal, {
          zona: zonaPrincipal,
          tipo: tipo,
          zonaHFC: meta.zonaHFC || '',
          zonaFTTH: meta.zonaFTTH || '',
          territorio: meta.territorio || '',
          territorios: new Set(),
          ordenes: []
        });
      }

      const zone = zoneGroups.get(zonaPrincipal);
      zone.ordenes.push(r);
      if (meta.territorio) zone.territorios.add(meta.territorio);
    });

    return Array.from(zoneGroups.values()).map(z => ({
      ...z,
      territorios: Array.from(z.territorios)
    }));

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
    today.setHours(0,0,0,0);

    console.log(`✅ Analizando zonas con ventana de ${daysWindow} días`);

    return zones.map(z => {
      const seenOTperDay = new Map();
      const ordenesPerDay = new Map();
      let ordenesEnVentana = 0;

      const diagnosticos = new Set();

      z.ordenes.forEach(o => {
        const meta = o.__meta || {};
        if (!meta.fecha || meta.daysFromToday > daysWindow) return;

        ordenesEnVentana++;

        const dk = meta.dayKey;
        if (dk == null) return;

        if (!seenOTperDay.has(dk)) seenOTperDay.set(dk, new Set());
        if (!ordenesPerDay.has(dk)) ordenesPerDay.set(dk, []);

        const cita = o['Número de cita'];
        if (cita) {
          seenOTperDay.get(dk).add(cita);
          ordenesPerDay.get(dk).push(o);
        }

        const diag = o['Diagnostico Tecnico'] || o['Diagnóstico Técnico'] || '';
        if (diag) diagnosticos.add(diag);
      });

      const todayKey = DateUtils.toDayKey(today);
      const yesterdayKey = DateUtils.toDayKey(new Date(today.getTime() - 86400000));

      const ingresoN = seenOTperDay.get(todayKey)?.size || 0;
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
    }).sort((a,b) => b.score - a.score);
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

    return Array.from(cmtsGroups.values()).sort((a,b) => b.totalOTs - a.totalOTs);
  }
}

const dataProcessor = new DataProcessor();
window.dataProcessor = dataProcessor;
})();
