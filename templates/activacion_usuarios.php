<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="es">
<head>
<?php
$recursosBase = __DIR__ . '/../recursos';
$recursosDisponibles = is_dir($recursosBase);
if ($recursosDisponibles) {
    @session_start();
    $herramienta = 'activacion_usuarios';
    if (file_exists($recursosBase . '/recursos.php')) {
        include $recursosBase . '/recursos.php';
    }
    if (file_exists($recursosBase . '/encabezado.php')) {
        include $recursosBase . '/encabezado.php';
    }
    if (function_exists('validar_sesion')) {
        validar_sesion($herramienta);
    } elseif (session_status() !== PHP_SESSION_ACTIVE) {
        session_start();
    }
    if (function_exists('headerBasico')) {
        headerBasico();
    }
    if (function_exists('headerBootstrap')) {
        headerBootstrap(1);
    }
    if (function_exists('headerDatatables')) {
        headerDatatables();
    }
    if (function_exists('log_visitas_a_web') && isset($con_w)) {
        log_visitas_a_web($con_w);
    }
} else {
    if (session_status() !== PHP_SESSION_ACTIVE) {
        session_start();
    }
}
$aleatorio = rand(1, 100000);
?>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Activacion Usuarios</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js" integrity="sha512-8bB/y10p+V6YxIcW1gk0CbCI3eNsAEBnAv1hJ3FtIoNvJ5HfLj1nHeWnKwtRA6gMzDX6IfElA9I6wO9iVbW3pQ==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
  <link rel="stylesheet" href="/assets/css/styles.css?v=<?php echo $aleatorio; ?>" />
<?php if ($recursosDisponibles) : ?>
  <script type="text/javascript" src="funciones.js?<?php echo $aleatorio; ?>"></script>
  <script type="text/javascript" src="../recursos/js.js?<?php echo $aleatorio; ?>"></script>
  <link rel="stylesheet" type="text/css" href="style.css?<?php echo $aleatorio; ?>" />
<?php endif; ?>
</head>
<body>
  <div class="card border-0">
    <div class="card-heading">
      <nav class="navbar navbar-expand-lg bg-dark navbar-dark">
        <h4 class="navbar-brand m-0" href="#">Nombre del MicroSitio</h4>
        <div class="collapse navbar-collapse" id="navbarSupportedContent">
          <ul class="navbar-nav mr-auto"></ul>
          <ul class="navbar-nav">
            <li class="nav-item">
              <span class="navbar-text">
                <span class="fa fa-user-circle" aria-hidden="true"></span>
                <?php if (!empty($_SESSION['user'])) : ?>
                  <?php echo 'Bienvenido: ' . htmlspecialchars($_SESSION['user'], ENT_QUOTES, 'UTF-8'); ?>
                <?php else : ?>
                  Usuario invitado
                <?php endif; ?>
              </span>
            </li>
            <li class="nav-item">
              <a href="../recursos/sesion/desconectar.php?pag=hubs" class="nav-link"><span class="fa fa-sign-out" aria-hidden="true"></span>CERRAR SESION</a>
            </li>
          </ul>
        </div>
      </nav>
    </div>

    <div class="card-body" style="padding: 0;">
      <div class="app-container">

        <header class="app-header">
          <div class="header-content">
            <div class="header-title">
              <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path>
                <polyline points="3.27 6.96 12 12.01 20.73 6.96"></polyline>
                <line x1="12" y1="22.08" x2="12" y2="12"></line>
              </svg>
              <div>
                <h1>📊 Panel Fulfillment <span class="version-badge">v5.0 COMPLETO</span></h1>
                <p>✨ Exportes detallados por zona • Filtro selectivo de zonas • Sparklines normalizadas • Históricos consistentes • Alarmas FMS • Equipos • Filtro anti-CATEC</p>
              </div>
            </div>
            <div class="header-actions">
              <div class="utilities-dropdown">
                <button class="btn-utilities" onclick="toggleUtilities()">
                  🔧 Utilidades
                </button>
                <div class="utilities-menu" id="utilitiesMenu">
                  <div class="utilities-menu-group">
                    <div class="utilities-menu-header">Recursos rápidos</div>
                    <div class="utilities-menu-item" onclick="openPlanillasNewTab()">
                      📋 Planillas
                    </div>
                    <div class="utilities-menu-item" onclick="openUsefulLinks()">
                      🔗 Links útiles
                    </div>
                  </div>
                </div>
              </div>
              <button class="btn-help" onclick="toast('💡 v5.0 Completo: Exportes detallados por zona + Filtro selectivo tipo Excel, Links útiles renovados, sparklines normalizadas y filtro anti-CATEC (últimas 24h)')">
                💡 Info
              </button>
            </div>
          </div>
        </header>

        <section class="load-panel">
          <div class="load-section">
            <div class="section-header">📁 CARGAR REPORTES</div>

            <div class="load-buttons">
              <div class="load-item">
                <label for="fileConsolidado1" class="btn btn-primary">📄 Consolidado 1</label>
                <input type="file" id="fileConsolidado1" accept=".xlsx,.xls,.csv" hidden>
                <span class="file-status" id="status1">Sin cargar</span>
              </div>

              <div class="load-item">
                <label for="fileConsolidado2" class="btn btn-primary">📄 Consolidado 2</label>
                <input type="file" id="fileConsolidado2" accept=".xlsx,.xls,.csv" hidden>
                <span class="file-status" id="status2">Sin cargar</span>
              </div>

              <div class="load-item">
                <label for="fileNodos" class="btn btn-primary">🌐 Nodos UP/DOWN</label>
                <input type="file" id="fileNodos" accept=".xlsx,.xls,.csv" hidden>
                <span class="file-status" id="status3">Sin cargar</span>
              </div>

              <div class="load-item">
                <label for="fileFMS" class="btn btn-primary">🚨 Alarmas FMS</label>
                <input type="file" id="fileFMS" accept=".xlsx,.xls,.csv" hidden>
                <span class="file-status" id="status4">Sin cargar</span>
              </div>

              <div class="merge-status" id="mergeStatus" style="display: none;">
                ✓ <span id="mergeStatusText">Procesando...</span>
              </div>
            </div>
          </div>
        </section>

        <nav class="filters-panel">
          <div class="filter-group">
            <div class="filter-label">🔥 Ordenar Zonas</div>
            <select id="ordenarPorIngreso" class="input-select">
              <option value="desc">Mayor a Menor Ingreso</option>
              <option value="asc">Menor a Mayor Ingreso</option>
              <option value="">Sin ordenar (Score)</option>
            </select>
          </div>

          <div class="filter-group">
            <div class="filter-label">Filtros Base</div>
            <label class="checkbox-label">
              <input type="checkbox" id="filterCATEC">
              <span>Solo CATEC</span>
            </label>
            <label class="toggle-control">
              <input type="checkbox" id="filterExcludeCATEC">
              <span class="toggle-track">
                <span class="toggle-thumb"></span>
              </span>
              <span class="toggle-text">Excluir CATEC</span>
            </label>
            <label class="checkbox-label">
              <input type="checkbox" id="showAllStates">
              <span>Ver todos los estados</span>
            </label>
          </div>

          <div class="filter-group">
            <div class="filter-label">Tecnología</div>
            <label class="checkbox-label">
              <input type="checkbox" id="filterFTTH">
              <span>Solo FTTH (9xx)</span>
            </label>
            <label class="checkbox-label">
              <input type="checkbox" id="filterExcludeFTTH">
              <span>Excluir FTTH</span>
            </label>
          </div>

          <div class="filter-group">
            <div class="filter-label">Estado Nodo</div>
            <select id="filterNodoEstado" class="input-select">
              <option value="">Todos</option>
              <option value="up">Solo UP</option>
              <option value="down">Solo DOWN</option>
              <option value="critical">Críticos (>50 DOWN)</option>
            </select>
          </div>

          <div class="filter-group">
            <div class="filter-label">CMTS</div>
            <select id="filterCMTS" class="input-select">
              <option value="">Todos</option>
            </select>
          </div>

          <div class="filter-group">
            <div class="filter-label">⭐ Ventana Temporal</div>
            <select id="daysWindow" class="input-select">
              <option value="3" selected>3 días</option>
              <option value="7">7 días</option>
              <option value="14">14 días</option>
              <option value="30">30 días</option>
            </select>
          </div>

          <div class="filter-group">
            <div class="filter-label">Territorio</div>
            <select id="filterTerritorio" class="input-select">
              <option value="">Todos</option>
            </select>
          </div>

          <div class="filter-group">
            <div class="filter-label">CMTS Vendor</div>
            <select id="filterCMTSVendor" class="input-select">
              <option value="">Todos</option>
            </select>
          </div>

          <div class="filter-group">
            <div class="filter-label">CAC / Zona</div>
            <div class="zone-select" id="zoneSelect"></div>
          </div>

          <div class="filter-group">
            <div class="filter-label">Equipos</div>
            <label class="checkbox-label">
              <input type="checkbox" id="filterEquipoONUs">
              <span>ONU (ONT)</span>
            </label>
            <label class="checkbox-label">
              <input type="checkbox" id="filterEquipoCablemodems">
              <span>CM</span>
            </label>
            <label class="checkbox-label">
              <input type="checkbox" id="filterEquipoDecos">
              <span>Decos</span>
            </label>
          </div>

          <div class="filter-group">
            <div class="filter-label">Búsqueda rápida</div>
            <div class="quick-search">
              <input type="text" id="quickSearchInput" class="input-text" placeholder="Buscar caso, orden, técnico, zona...">
              <button class="btn btn-secondary" onclick="resetFilters()">Limpiar</button>
            </div>
          </div>
        </nav>

        <section class="stats-panel">
          <div class="stats-grid" id="statsGrid"></div>
        </section>

        <section class="content-panel">
          <div class="left-column">
            <div class="panel">
              <div class="panel-header">
                <div>
                  <h2>Zonas críticas</h2>
                  <p id="zonesSummary">Cargá archivos para ver las zonas</p>
                </div>
                <div class="panel-actions">
                  <button class="btn btn-primary" id="btnExportVista" onclick="exportExcelVista()">📤 Exportar vista (Excel)</button>
                </div>
              </div>
              <div class="panel-body">
                <div class="zones-table-header">
                  <div class="col-zone">Zona / CAC</div>
                  <div class="col-ingresos">Ingresos</div>
                  <div class="col-porcentaje">%</div>
                  <div class="col-pendientes">Pendientes</div>
                  <div class="col-sparkline">Últimos 7 días</div>
                  <div class="col-acciones"></div>
                </div>
                <div class="zones-table" id="zonesTable"></div>
              </div>
            </div>

            <div class="panel" id="panelAlarmas">
              <div class="panel-header">
                <div>
                  <h2>Alarmas relevantes</h2>
                  <p>Alertas críticas de FMS combinadas con estado del nodo</p>
                </div>
                <div class="panel-actions">
                  <button class="btn btn-secondary" onclick="openAlarmaModal()">Ver todas</button>
                </div>
              </div>
              <div class="panel-body">
                <div class="alarms-list" id="alarmsList"></div>
              </div>
            </div>

            <div class="panel">
              <div class="panel-header">
                <div>
                  <h2>Estadísticas generales</h2>
                  <p>Resumen de KPIs y performance por tecnología</p>
                </div>
              </div>
              <div class="panel-body">
                <div class="stats-cards" id="kpiCards"></div>
              </div>
            </div>
          </div>

          <div class="right-column">
            <div class="panel">
              <div class="panel-header">
                <div>
                  <h2>Detalle de zona</h2>
                  <p id="detailSummary">Seleccioná una zona para ver los detalles</p>
                </div>
                <div class="panel-actions">
                  <button class="btn btn-primary" id="btnExportZona" disabled onclick="exportExcelZona()">📥 Exportar zona</button>
                  <button class="btn btn-secondary" onclick="exportarEquipos()">📦 Exportar equipos</button>
                </div>
              </div>
              <div class="panel-body">
                <div class="zone-details" id="zoneDetails"></div>
              </div>
            </div>

            <div class="panel">
              <div class="panel-header">
                <div>
                  <h2>Diagnósticos técnicos destacados</h2>
                  <p>Principales motivos por zona y tecnología</p>
                </div>
              </div>
              <div class="panel-body">
                <div class="diagnosticos-grid" id="diagnosticosGrid"></div>
              </div>
            </div>

            <div class="panel">
              <div class="panel-header">
                <div>
                  <h2>Nodos críticos</h2>
                  <p>Coincidencias entre nodos críticos y órdenes abiertas</p>
                </div>
              </div>
              <div class="panel-body">
                <div class="nodes-table" id="nodesTable"></div>
              </div>
            </div>
          </div>
        </section>

        <section class="footer-panel">
          <div class="footer-content">
            <div>
              <h3>Ayudas rápidas</h3>
              <ul>
                <li>💡 Usá el filtro de zonas para seleccionar múltiples CAC al estilo Excel.</li>
                <li>📤 Exportá la vista completa para generar una pestaña por zona filtrada.</li>
                <li>🚨 Revisá las alarmas FMS combinadas con el estado del nodo para priorizar visitas.</li>
              </ul>
            </div>
            <div>
              <h3>Atajos</h3>
              <div class="shortcut-buttons">
                <button class="btn btn-outline" onclick="filtrarHoy()">Hoy</button>
                <button class="btn btn-outline" onclick="filtrarAyer()">Ayer</button>
                <button class="btn btn-outline" onclick="filtrarUltimos(3)">Últimos 3 días</button>
                <button class="btn btn-outline" onclick="filtrarUltimos(7)">Últimos 7 días</button>
              </div>
            </div>
            <div>
              <h3>Diagnósticos frecuentes</h3>
              <div class="diagnosticos-chips" id="diagnosticosChips"></div>
            </div>
          </div>
        </section>
      </div>
    </div>

    <div class="card-footer">
      <label>Performance Eficiencia y Mejora</label>
      <a href="mailto:dataofficepem@teco.com.ar" target="_top">dataofficepem@teco.com.ar</a>
    </div>
  </div>

  <div id="toastContainer"></div>

  <div id="modalBackdrop" class="modal-backdrop" onclick="closeModal()">
    <div class="modal" onclick="event.stopPropagation()">
      <div class="modal-header">
        <div class="modal-title" id="modalTitle">Detalle de Zona</div>
        <button class="modal-close" onclick="closeModal()">×</button>
      </div>

      <div class="modal-body" id="modalBody">
        <div class="loading-message">
          <div class="spinner"></div>
          <p>Cargando detalles...</p>
        </div>
      </div>

      <div class="modal-footer" id="modalFooter">
        <div class="selection-info" id="selectionInfo">
          0 órdenes seleccionadas
        </div>
        <div style="display: flex; gap: 8px;">
          <button class="btn btn-warning" id="btnExportBEFAN" disabled onclick="exportBEFAN()">
            📤 Exportar BEFAN (TXT)
          </button>
          <button class="btn btn-secondary" onclick="closeModal()">
            Cerrar
          </button>
        </div>
      </div>
    </div>
  </div>

  <div id="alarmaBackdrop" class="modal-backdrop" onclick="closeAlarmaModal()">
    <div class="modal" onclick="event.stopPropagation()" style="max-width: 1000px;">
      <div class="modal-header">
        <div class="modal-title">🚨 Información de Alarmas</div>
        <button class="modal-close" onclick="closeAlarmaModal()">×</button>
      </div>

      <div class="modal-body" id="alarmaModalBody">
        <div class="loading-message">
          <div class="spinner"></div>
          <p>Cargando datos de alarmas...</p>
        </div>
      </div>

      <div class="modal-footer">
        <button class="btn btn-primary" onclick="closeAlarmaModal()">
          Cerrar
        </button>
      </div>
    </div>
  </div>

  <div id="edificioBackdrop" class="modal-backdrop" onclick="closeEdificioModal()">
    <div class="modal" onclick="event.stopPropagation()" style="max-width: 1200px;">
      <div class="modal-header">
        <div class="modal-title" id="edificioModalTitle">Detalle de Edificio</div>
        <button class="modal-close" onclick="closeEdificioModal()">×</button>
      </div>

      <div class="modal-body" id="edificioModalBody">
        <div class="loading-message">
          <div class="spinner"></div>
          <p>Cargando órdenes del edificio...</p>
        </div>
      </div>

      <div class="modal-footer">
        <button class="btn btn-primary" onclick="closeEdificioModal()">
          Cerrar
        </button>
      </div>
    </div>
  </div>

  <iframe class="d-none" id="miframe" name="miframe"></iframe>
  <script src="/assets/js/app.js?v=<?php echo $aleatorio; ?>"></script>
</body>
</html>
