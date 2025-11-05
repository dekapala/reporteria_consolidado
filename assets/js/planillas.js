(function(){
  const STORAGE_KEY = 'planillasNotas';
  const templates = [
    {
      nombre: 'Asignación diaria',
      descripcion: 'Listado base de OTs a despachar con datos de contacto y horario.',
      columnas: ['N° Orden', 'Cliente', 'Dirección', 'Distrito', 'Teléfono', 'Horario', 'Observaciones']
    },
    {
      nombre: 'Checklist de ejecución',
      descripcion: 'Control de acciones realizadas por cada técnico en la visita.',
      columnas: ['N° Orden', 'Técnico', 'Hora llegada', 'Hora cierre', 'Acción realizada', 'Requiere escalamiento', 'Comentarios']
    },
    {
      nombre: 'Incidencias',
      descripcion: 'Registro de incidencias o restricciones detectadas en campo.',
      columnas: ['Fecha', 'Zona', 'N° Orden', 'Motivo', 'Contacto', 'Acción tomada', 'Estado']
    }
  ];

  function loadNotes(){
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return [];
      const notes = JSON.parse(raw);
      return Array.isArray(notes) ? notes : [];
    } catch (err) {
      console.warn('No se pudieron cargar las notas de planillas', err);
      return [];
    }
  }

  function saveNotes(notes){
    localStorage.setItem(STORAGE_KEY, JSON.stringify(notes));
  }

  function renderTemplates(container){
    container.innerHTML = '';
    templates.forEach(t => {
      const card = document.createElement('div');
      card.className = 'template-card';
      card.innerHTML = `
        <h3>${t.nombre}</h3>
        <p class="template-meta">${t.descripcion}</p>
        <div class="template-columns">
          ${t.columnas.map(col => `<span>${col}</span>`).join('')}
        </div>
      `;
      container.appendChild(card);
    });
  }

  function exportTemplates(){
    const wb = XLSX.utils.book_new();
    templates.forEach(t => {
      const sheet = XLSX.utils.aoa_to_sheet([t.columnas]);
      XLSX.utils.book_append_sheet(wb, sheet, t.nombre.slice(0, 28));
    });
    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Planillas_operativas_${fecha}.xlsx`);
    if (typeof toast === 'function') {
      toast('✅ Planillas generadas');
    }
    const lastGenerated = document.getElementById('lastGenerated');
    if (lastGenerated) {
      lastGenerated.textContent = `Última descarga: ${new Date().toLocaleString('es-PE', { dateStyle: 'medium', timeStyle: 'short' })}`;
    }
  }

  function renderNotes(listEl, notes, onDelete){
    listEl.innerHTML = '';
    if (!notes.length) {
      const empty = document.createElement('p');
      empty.textContent = 'No tienes notas registradas aún.';
      empty.style.color = 'rgba(226, 232, 240, 0.65)';
      listEl.appendChild(empty);
      return;
    }

    notes.forEach((note, idx) => {
      const item = document.createElement('li');
      item.innerHTML = `<span>${note}</span>`;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = 'Eliminar';
      btn.addEventListener('click', () => onDelete(idx));
      item.appendChild(btn);
      listEl.appendChild(item);
    });
  }

  function init(){
    const templatesList = document.getElementById('templatesList');
    const btnDescargar = document.getElementById('btnDescargarPlanilla');
    const fechaActual = document.getElementById('fechaActual');
    const shareLinkInput = document.getElementById('shareLink');
    const btnCopyLink = document.getElementById('btnCopyLink');
    const notesForm = document.getElementById('notesForm');
    const notesInput = document.getElementById('noteText');
    const notesList = document.getElementById('notesList');

    if (fechaActual) {
      fechaActual.textContent = `Actualizado el ${new Date().toLocaleDateString('es-PE', { dateStyle: 'full' })}`;
    }

    if (templatesList) {
      renderTemplates(templatesList);
    }

    if (btnDescargar) {
      btnDescargar.addEventListener('click', exportTemplates);
    }

    if (shareLinkInput) {
      const { origin, href, pathname } = window.location;
      const normalizedOrigin = origin && origin !== 'null'
        ? origin
        : href.replace(pathname, '');
      const basePath = pathname.replace(/\/[^/]*$/, '');
      const rawLink = `${normalizedOrigin}${basePath}/planillas.html`;
      shareLinkInput.value = rawLink.replace(/\/+planillas\.html$/, '/planillas.html');
    }

    if (btnCopyLink && shareLinkInput) {
      const copyLink = async () => {
        const link = shareLinkInput.value;
        if (!link) return;
        if (navigator.clipboard?.writeText) {
          try {
            await navigator.clipboard.writeText(link);
            if (typeof toast === 'function') {
              toast('🔗 Enlace copiado');
            }
            return;
          } catch (err) {
            console.warn('No se pudo usar navigator.clipboard', err);
          }
        }
        shareLinkInput.select();
        document.execCommand('copy');
        if (typeof toast === 'function') {
          toast('🔗 Enlace copiado');
        }
      };

      btnCopyLink.addEventListener('click', copyLink);
    }

    if (notesForm && notesInput && notesList) {
      let notes = loadNotes();

      const refreshNotes = () => {
        renderNotes(notesList, notes, (idx) => {
          notes = notes.filter((_, i) => i !== idx);
          saveNotes(notes);
          refreshNotes();
        });
      };

      refreshNotes();

      notesForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const value = notesInput.value.trim();
        if (!value) return;
        notes = [value, ...notes].slice(0, 20);
        saveNotes(notes);
        refreshNotes();
        notesInput.value = '';
        notesInput.focus();
      });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
