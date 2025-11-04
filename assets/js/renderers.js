(function(){
const UIRenderer = {
  renderStats(data) {
    return `
      <div class="stat-card"><div class="stat-label">Total Órdenes</div><div class="stat-value">${NumberUtils.format(data.total)}</div></div>
      <div class="stat-card"><div class="stat-label">Zonas Analizadas</div><div class="stat-value">${NumberUtils.format(data.zonas)}</div></div>
      <div class="stat-card"><div class="stat-label">Zonas Críticas</div><div class="stat-value">${NumberUtils.format(data.criticas)}</div></div>
      <div class="stat-card"><div class="stat-label">Zonas FTTH</div><div class="stat-value">${NumberUtils.format(data.ftth)}</div></div>
      <div class="stat-card"><div class="stat-label">Con Alarmas</div><div class="stat-value">${NumberUtils.format(data.conAlarmas)}</div></div>
      <div class="stat-card"><div class="stat-label">Nodos Críticos</div><div class="stat-value">${NumberUtils.format(data.nodosCriticos)}</div></div>
    `;
  },

  renderSparkline(counts, labels) {
    const max = Math.max(...counts, 1);
    const width = 120;
    const height = 30;
    const barWidth = width / counts.length;

    let svg = `<svg class="sparkline" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">`;

    counts.forEach((count, i) => {
      const barHeight = (count / max) * height;
      const x = i * barWidth;
      const y = height - barHeight;
      const color = count >= max * 0.7 ? '#D13438' : count >= max * 0.4 ? '#F7630C' : '#0078D4';

      svg += `<rect x="${x}" y="${y}" width="${barWidth - 2}" height="${barHeight}" fill="${color}" rx="2"/>`;
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
      const badgeTipo = z.tipo === 'FTTH' ? '<span class="badge badge-ftth">FTTH</span>' : '<span class="badge badge-hfc">HFC</span>';
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
        <td><button class="btn btn-primary" style="padding: 6px 12px; font-size: 0.75rem;" onclick="openModal(${idx})">👁️ Ver</button></td>
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
        <td><button class="btn btn-primary" style="padding: 6px 12px; font-size: 0.75rem;" onclick="openCMTSDetail('${c.cmts.replace(/'/g, "\\'")}')">👁️ Ver Zonas</button></td>
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
      .sort((a,b) => b.casos.length - a.casos.length)
      .slice(0, 50);

    if (!sorted.length) return '<div class="loading-message"><p>No hay edificios con 2+ incidencias</p></div>';

    let html = '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
    html += '<th>Dirección</th><th>Zona</th><th>Territorio</th><th class="number">Total OTs</th>';
    html += '</tr></thead><tbody>';

    sorted.forEach((e, idx) => {
      html += `<tr class="clickable" onclick="showEdificioDetail(${idx})">
        <td><strong>${e.direccion}</strong></td>
        <td>${e.zona}</td>
        <td>${e.territorio}</td>
        <td class="number">${e.casos.length}</td>
      </tr>`;
    });

    html += '</tbody></table></div></div>';

    window.edificiosData = sorted;

    return html;
  },

  renderEquipos(ordenes) {
    if (!ordenes.length) {
      return '<div class="loading-message"><p>⚠️ Aplica filtros para ver equipos asociados.</p></div>';
    }

    if (!equipmentCache || !equipmentCache.size) {
      return '<div class="loading-message"><p>⚠️ No se pudo construir la cache de equipos. Verifica la columna de Información de Dispositivos.</p></div>';
    }

    const allowedCases = new Set();
    const allowedZones = new Set();

    ordenes.forEach(order => {
      const meta = order.__meta || {};
      if (meta.numeroCaso) allowedCases.add(meta.numeroCaso);
      else if (order['Número de cita']) allowedCases.add(order['Número de cita']);
      if (meta.zonaPrincipal) allowedZones.add(meta.zonaPrincipal);
    });

    const modelos = new Set();
    const marcas = new Set();
    const grupos = new Map();

    allowedZones.forEach(zona => {
      const items = equipmentCache.get(zona) || [];
      const filtradoPorOrden = items.filter(item => {
        if (!item.numCaso) return true;
        return allowedCases.has(item.numCaso);
      });
      if (!filtradoPorOrden.length) return;

      filtradoPorOrden.forEach(item => {
        if (item.modelo) modelos.add(item.modelo);
        if (item.marca) marcas.add(item.marca);
      });

      grupos.set(zona, filtradoPorOrden);
    });

    if (!grupos.size) {
      return '<div class="loading-message"><p>⚠️ No se encontraron equipos para la vista filtrada.</p></div>';
    }

    if (!window.equiposOpen) window.equiposOpen = new Set();

    let html = '<div class="equipos-filters">';
    html += '<div class="filter-group">';
    html += '<div class="filter-label">Filtrar por Modelo</div>';
    html += '<select id="filterEquipoModelo" class="input-select" onchange="applyEquiposFilters()">';
    html += '<option value="">Todos los modelos</option>';
    Array.from(modelos).sort().forEach(m => {
      const selected = Filters.equipoModelo === m ? 'selected' : '';
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

    html += '<div class="filter-actions">';
    html += '<button class="btn btn-secondary" onclick="clearEquiposFilters()">Limpiar filtros</button>';
    html += '</div>';
    html += '</div>';

    const gruposFiltrados = new Map();
    grupos.forEach((items, zona) => {
      const filtrado = items.filter(item => {
        if (Filters.equipoModelo && item.modelo !== Filters.equipoModelo) return false;
        if (Filters.equipoMarca && item.marca !== Filters.equipoMarca) return false;
        return true;
      });
      if (filtrado.length) gruposFiltrados.set(zona, filtrado);
    });

    const zonasOrdenadas = Array.from(gruposFiltrados.keys()).sort((a, b) => a.localeCompare(b));

    html += '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
    html += '<th>Zona</th><th class="number">Equipos</th><th>Acciones</th>';
    html += '</tr></thead><tbody>';

    zonasOrdenadas.forEach(zona => {
      const arr = gruposFiltrados.get(zona);
      const isOpen = window.equiposOpen && window.equiposOpen.has(zona);

      html += `<tr class="${isOpen ? 'expanded' : ''}">`;
      html += `<td><button class="link-btn" onclick="toggleEquiposGrupo('${zona.replace(/'/g, "\\'")}')">${isOpen ? '▼' : '▶'} ${zona}</button></td>`;
      html += `<td class="number">${arr.length}</td>`;
      html += '<td>';
      html += `<button class="btn btn-primary btn-sm" onclick="exportEquiposGrupoExcel('${zona.replace(/'/g, "\\'")}', false)">Exportar zona</button>`;
      html += `<button class="btn btn-secondary btn-sm" onclick="exportEquiposGrupoExcel('${zona.replace(/'/g, "\\'")}', true)" style="margin-left:6px">Exportar vista</button>`;
      html += '</td>';
      html += '</tr>';

      if (isOpen) {
        html += '<tr class="equipos-detalle"><td colspan="3">';
        html += '<div class="table-container"><div class="table-wrapper"><table class="detail-table"><thead><tr>';
        html += '<th>OT/Caso</th><th>Sistema</th><th>Serial</th><th>MAC</th><th>Tipo</th><th>Marca</th><th>Modelo</th>';
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
        html += '</td></tr>';
      }
    });

    html += '</tbody></table></div></div>';
    html += `<p style="text-align:center;margin-top:12px;color:var(--text-secondary)">Total: ${zonasOrdenadas.length} zonas • Click para expandir/colapsar • Exporta por zona o global</p>`;

    window.equiposPorZona = gruposFiltrados;
    window.equiposPorZonaCompleto = equipmentCache;

    return html;
  }
};
window.UIRenderer = UIRenderer;
})();
