/**
 * UI Renderer v5.0 - Módulo de Presentación
 * Se encarga de generar todo el HTML de las tablas y stats.
 */
(function(){

    const UIRenderer = {
      // 1. Renderizado de Tarjetas Superiores (Stats)
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
    
      // 2. Renderizado de Tabla FMS (NUEVO y CORREGIDO)
      renderFMS(alarmas) {
        if (!alarmas || !alarmas.length) {
          return '<div class="loading-message"><p>✅ No hay alarmas FMS activas para la selección actual.</p></div>';
        }
    
        let html = '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
        html += '<th>ID</th><th>Zona</th><th>Elemento</th><th>Daño</th><th>Descripción</th><th>Inicio</th><th>Duración</th><th>Estado</th>';
        html += '</tr></thead><tbody>';
    
        // Ordenar por fecha de creación descendente
        const sorted = alarmas.sort((a, b) => {
            return new Date(b.creationDate) - new Date(a.creationDate);
        });
    
        sorted.slice(0, 500).forEach(a => { // Limite de renderizado por performance
          const badgeEstado = a.isActive 
            ? '<span class="badge badge-critical">ACTIVA</span>' 
            : '<span class="badge badge-up">CERRADA</span>';
          
          let duracion = '-';
          if (a.creationDate) {
             const start = DateUtils.parse(a.creationDate); 
             const end = a.isActive ? new Date() : (a.recoveryDate ? DateUtils.parse(a.recoveryDate) : new Date());
             if (start && end) {
                 const diffHrs = Math.round((end - start) / (1000 * 60 * 60));
                 duracion = `${diffHrs} hs`;
             }
          }
    
          html += `<tr>
            <td style="font-family:monospace; font-size:0.75rem;">${a.eventId || '-'}</td>
            <td><strong>${a.zonaTecnica || '-'}</strong></td> 
            <td>${a.elementCode || '-'} <br><small style="color:var(--text-secondary)">${a.elementType}</small></td>
            <td>${a.damage || '-'}</td>
            <td title="${a.description || ''}">${(a.description || '').substring(0, 60)}...</td>
            <td>${a.creationDate || '-'}</td>
            <td>${duracion}</td>
            <td>${badgeEstado}</td>
          </tr>`;
        });
    
        html += '</tbody></table></div></div>';
        html += `<p style="text-align:center;margin-top:12px;color:var(--text-secondary)">Mostrando ${Math.min(sorted.length, 500)} de ${sorted.length} alarmas</p>`;
        
        return html;
      },
      
      // 3. Helpers gráficos
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
      
      // 4. Renderizado de Zonas
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
    
      // 5. Renderizado de CMTS
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
      
      // 6. Renderizado de Edificios
      renderEdificios(ordenes) {
        const edificios = new Map();
        
        ordenes.forEach(o => {
          const dir = TextUtils.normalize(o['Calle']);
          if (!dir || dir.length < 5) return;
          
          const zona = o['Zona Tecnica HFC'] || o['Zona Tecnica FTTH'] || '';
          
          if (!edificios.has(dir)) {
            edificios.set(dir, {
              direccion: o['Calle'],
              zona: zona,
              territorio: o['Territorio de servicio: Nombre'] || '',
              casos: [],
              zonas: new Set()
            });
          }
          edificios.get(dir).casos.push(o);
          if (zona) edificios.get(dir).zonas.add(zona);
        });
        
        // Agregar información de alarmas a cada edificio
        edificios.forEach(edificio => {
          let alarmasTotal = 0;
          let alarmasActivas = 0;
          const alarmasDetalle = [];
          
          // Buscar alarmas por cada zona del edificio
          edificio.zonas.forEach(zona => {
            const alarmas = dataProcessor.fmsMap.get(zona) || [];
            alarmasTotal += alarmas.length;
            const activas = alarmas.filter(a => a.isActive);
            alarmasActivas += activas.length;
            
            if (alarmas.length > 0) {
              alarmasDetalle.push({
                zona: zona,
                alarmas: alarmas,
                activas: activas.length
              });
            }
          });
          
          edificio.alarmasTotal = alarmasTotal;
          edificio.alarmasActivas = alarmasActivas;
          edificio.alarmasDetalle = alarmasDetalle;
          edificio.tieneAlarma = alarmasActivas > 0;
        });
        
        const sorted = Array.from(edificios.values())
          .filter(e => e.casos.length >= 2)
          .sort((a, b) => b.casos.length - a.casos.length)
          .slice(0, 50);
        
        if (!sorted.length) return '<div class="loading-message"><p>No hay edificios con 2+ incidencias</p></div>';
        
        let html = '<div class="table-container"><div class="table-wrapper"><table><thead><tr>';
        html += '<th>Dirección</th><th>Zona</th><th>Territorio</th><th class="number">Total OTs</th><th class="number">Alarmas</th><th>Acciones</th>';
        html += '</tr></thead><tbody>';
        
        sorted.forEach((e, idx) => {
          const alarmaClass = e.alarmasActivas > 0 ? 'alarma-activa' : '';
          const alarmaIcon = e.alarmasActivas > 0 ? '🚨' : (e.alarmasTotal > 0 ? '⚠️' : '✅');
          const alarmaText = e.alarmasActivas > 0 ? `${e.alarmasActivas} activas` : (e.alarmasTotal > 0 ? `${e.alarmasTotal} hist.` : 'Sin alarmas');
          
          html += `<tr class="clickable ${alarmaClass}">
            <td onclick="showEdificioDetail(${idx})"><strong>${e.direccion}</strong></td>
            <td onclick="showEdificioDetail(${idx})">${e.zona}</td>
            <td onclick="showEdificioDetail(${idx})">${e.territorio}</td>
            <td class="number" onclick="showEdificioDetail(${idx})">${e.casos.length}</td>
            <td class="number ${alarmaClass}" onclick="showEdificioDetail(${idx})">${alarmaIcon} ${alarmaText}</td>
            <td>
              <button class="btn btn-sm btn-info" onclick="event.stopPropagation(); showEdificioAlarmas(${idx})" ${e.alarmasTotal === 0 ? 'disabled' : ''}>
                🔍 Ver FMS
              </button>
            </td>
          </tr>`;
        });
        
        html += '</tbody></table></div></div>';
        window.edificiosData = sorted;
        return html;
      },
      
      // 7. Renderizado de Equipos
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
          // Se asume que Filters es global, pero aquí en renderer es mejor leer el estado si es necesario,
          // o simplemente renderizar lo que llega. Por simplicidad, renderizamos opciones.
          const selected = (window.Filters && window.Filters.equipoModelo && window.Filters.equipoModelo.includes(m)) ? 'selected' : '';
          html += `<option value="${m}" ${selected}>${m || '(Sin modelo)'}</option>`;
        });
        html += '</select>';
        html += '</div>';
    
        html += '<div class="filter-group">';
        html += '<div class="filter-label">Filtrar por Marca</div>';
        html += '<select id="filterEquipoMarca" class="input-select" onchange="applyEquiposFilters()">';
        html += '<option value="">Todas las marcas</option>';
        Array.from(marcas).sort().forEach(m => {
           const selected = (window.Filters && window.Filters.equipoMarca === m) ? 'selected' : '';
          html += `<option value="${m}" ${selected}>${m || '(Sin marca)'}</option>`;
        });
        html += '</select>';
        html += '</div>';
    
        html += '<div class="filter-group">';
        html += '<div class="filter-label">Filtrar por Territorio</div>';
        html += '<select id="filterEquipoTerritorio" class="input-select" onchange="applyEquiposFilters()">';
        html += '<option value="">Todos los territorios</option>';
        Array.from(territorios).sort().forEach(t => {
           const selected = (window.Filters && window.Filters.equipoTerritorio === t) ? 'selected' : '';
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
          
          // Leemos filtros globales si existen
          if (window.Filters) {
              if (window.Filters.equipoModelo && window.Filters.equipoModelo.length > 0) {
                filtrado = filtrado.filter(e => window.Filters.equipoModelo.includes(e.modelo));
              }
              if (window.Filters.equipoMarca) {
                filtrado = filtrado.filter(e => e.marca === window.Filters.equipoMarca);
              }
              if (window.Filters.equipoTerritorio) {
                filtrado = filtrado.filter(e => e.territorio === window.Filters.equipoTerritorio);
              }
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

    // Exponer globalmente
    window.UIRenderer = UIRenderer;
    console.log('✅ UIRenderer v5.0 cargado (FMS Fixed)');

})();
