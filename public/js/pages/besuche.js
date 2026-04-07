// ========================================
// Besuche-Seite (Liste + Detail-Modal)
// Mit branchenspezifischen Icons
// ========================================

async function renderBesuche() {
  const main = document.getElementById('mainContent');

  main.innerHTML = `
    <div class="page-header">
      <div>
        <h1>Besuchsberichte</h1>
        <p class="subtitle">Alle erfassten Kundenbesuche</p>
      </div>
      <a href="#/neuer-besuch" class="btn btn-dark">+ Neuer Bericht</a>
    </div>

    <div class="search-bar">
      <input type="text" class="search-input" id="besucheSuche" placeholder="Nach Kunde, Produkt, Kontakt suchen..." onkeyup="besucheSuchen()">
      <select class="filter-select" id="besucheSegmentFilter" onchange="besucheSuchen()">
        <option value="">Alle Segmente</option>
      </select>
      <select class="filter-select" id="besucheKWFilter" onchange="besucheSuchen()">
        <option value="">Alle Kalenderwochen</option>
      </select>
    </div>

    <div class="result-count" id="besucheCount"></div>
    <div id="besucheListe">
      <div class="loading"><span class="spinner"></span> Besuche laden...</div>
    </div>
  `;

  await ladeBesucheListe();

  // Segment-Filter von Dashboard/Wochenbericht uebernehmen
  if (window.pendingSegmentFilter) {
    const segSelect = document.getElementById('besucheSegmentFilter');
    if (segSelect) {
      segSelect.value = window.pendingSegmentFilter;
    }
    window.pendingSegmentFilter = null;
    await besucheSuchen();
  }
}

let alleBesucheCache = [];

async function ladeBesucheListe(filter = {}) {
  try {
    const besuche = await API.besuche.laden(filter);
    alleBesucheCache = besuche;

    // KW-Dropdown befuellen
    const kws = [...new Set(besuche.map(b => b.kw).filter(Boolean))].sort((a, b) => b - a);
    const kwSelect = document.getElementById('besucheKWFilter');
    if (kwSelect) {
      const currentVal = kwSelect.value;
      kwSelect.innerHTML = '<option value="">Alle Kalenderwochen</option>' +
        kws.map(kw => `<option value="${kw}" ${kw == currentVal ? 'selected' : ''}>KW ${kw}</option>`).join('');
    }

    // Segment-Dropdown befuellen
    const segmente = [...new Set(besuche.map(b => b.segment || 'Sonstiges'))].sort();
    const segSelect = document.getElementById('besucheSegmentFilter');
    if (segSelect) {
      const currentSeg = segSelect.value;
      segSelect.innerHTML = '<option value="">Alle Segmente</option>' +
        segmente.map(s => `<option value="${s}" ${s === currentSeg ? 'selected' : ''}>${s}</option>`).join('');
    }

    renderBesucheListe(besuche);
  } catch (error) {
    document.getElementById('besucheListe').innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">\u26A0\uFE0F</div>
        <p class="empty-state-text">Fehler beim Laden: ${error.message}</p>
      </div>
    `;
  }
}

function renderBesucheListe(besuche) {
  const container = document.getElementById('besucheListe');
  const countEl = document.getElementById('besucheCount');

  if (countEl) countEl.textContent = `${besuche.length} Besuche`;

  if (besuche.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">\u{1F4CB}</div>
        <p class="empty-state-text">Keine Besuche gefunden</p>
      </div>
    `;
    return;
  }

  container.innerHTML = besuche.map(b => {
    const themenPreview = b.themen ? b.themen.substring(0, 120).replace(/\n/g, ' ') + (b.themen.length > 120 ? '...' : '') : '';
    const icon = getBranchenIcon(b.firma, b.segment);
    const farbe = getBranchenFarbe(b.firma, b.segment);

    return `
      <div class="besuch-card" onclick="zeigeBesuchDetail('${b.id}')">
        <div class="besuch-card-header">
          <span class="branchen-icon-circle" style="background:${farbe}; width:36px; height:36px; font-size:18px;">${icon}</span>
          <span class="besuch-card-firma">${escapeHtml(b.firma)}</span>
          ${b.kw ? `<span class="badge badge-kw">KW ${b.kw}</span>` : ''}
          ${b.besuchstyp ? `<span class="badge badge-typ">${escapeHtml(b.besuchstyp)}</span>` : ''}
        </div>
        <div class="besuch-card-meta">
          ${b.datum ? `<span>\u{1F4C5} ${formatDatum(b.datum)}</span>` : ''}
          ${b.ort ? `<span>\u{1F4CD} ${escapeHtml(b.ort)}</span>` : ''}
          ${b.kontaktperson ? `<span>\u{1F464} ${escapeHtml(b.kontaktperson)}</span>` : ''}
        </div>
        ${b.produkte ? `<div class="besuch-card-produkte"><strong>Produkte:</strong> ${escapeHtml(b.produkte)}</div>` : ''}
        ${themenPreview ? `<div class="besuch-card-themen">${escapeHtml(themenPreview)}</div>` : ''}
        <span class="besuch-card-arrow">\u203A</span>
      </div>
    `;
  }).join('');
}

async function besucheSuchen() {
  const suche = document.getElementById('besucheSuche')?.value?.trim() || '';
  const kw = document.getElementById('besucheKWFilter')?.value || '';
  const segment = document.getElementById('besucheSegmentFilter')?.value || '';

  const filter = {};
  if (suche) filter.suche = suche;
  if (kw) filter.kw = kw;

  // Segment-Filter client-seitig anwenden (wird nicht an API geschickt)
  await ladeBesucheListe(filter);

  if (segment && alleBesucheCache.length > 0) {
    const gefiltert = alleBesucheCache.filter(b => (b.segment || 'Sonstiges') === segment);
    renderBesucheListe(gefiltert);
  }
}

// Navigation von Dashboard/Wochenbericht: Segment vorfiltern
function filterBesucheNachSegment(segment) {
  window.pendingSegmentFilter = segment;
  window.location.hash = '#/besuche';
}

// ---- Detail-Modal ----

async function zeigeBesuchDetail(id) {
  const besuch = alleBesucheCache.find(b => b.id === id);
  if (!besuch) return;

  const fotos = parseFotos(besuch.fotos);

  const content = `
    <div class="modal-title">${escapeHtml(besuch.firma)}</div>

    <div class="flex gap-8 flex-wrap mb-16">
      ${besuch.kw ? `<span class="badge badge-kw">KW ${besuch.kw}</span>` : ''}
      ${besuch.besuchstyp ? `<span class="badge badge-typ">${escapeHtml(besuch.besuchstyp)}</span>` : ''}
    </div>

    <div class="modal-field">
      <div class="modal-field-label">Datum</div>
      <div class="modal-field-value">${formatDatum(besuch.datum)}</div>
    </div>

    ${besuch.ort ? `
    <div class="modal-field">
      <div class="modal-field-label">Ort</div>
      <div class="modal-field-value">${escapeHtml(besuch.ort)}</div>
    </div>` : ''}

    ${besuch.kontaktperson ? `
    <div class="modal-field">
      <div class="modal-field-label">Kontaktperson</div>
      <div class="modal-field-value">${escapeHtml(besuch.kontaktperson)}</div>
    </div>` : ''}

    ${besuch.segment ? `
    <div class="modal-field">
      <div class="modal-field-label">Segment</div>
      <div class="modal-field-value">${escapeHtml(besuch.segment)}</div>
    </div>` : ''}

    ${besuch.produkte ? `
    <div class="modal-field">
      <div class="modal-field-label">Produkte</div>
      <div class="modal-field-value">${escapeHtml(besuch.produkte)}</div>
    </div>` : ''}

    ${besuch.themen ? `
    <div class="modal-field">
      <div class="modal-field-label">Themen</div>
      <div class="modal-field-value">${textToBullets(besuch.themen)}</div>
    </div>` : ''}

    ${besuch.ergebnis ? `
    <div class="modal-field">
      <div class="modal-field-label">Ergebnis</div>
      <div class="modal-field-value">${textToBullets(besuch.ergebnis)}</div>
    </div>` : ''}

    ${besuch.naechsteSchritte ? `
    <div class="modal-field">
      <div class="modal-field-label">Naechste Schritte</div>
      <div class="modal-field-value">${textToBullets(besuch.naechsteSchritte)}</div>
    </div>` : ''}

    ${besuch.stellungnahme ? `
    <div class="modal-field">
      <div class="modal-field-label">Stellungnahme</div>
      <div class="modal-field-value">${escapeHtml(besuch.stellungnahme)}</div>
    </div>` : ''}

    ${besuch.projekt ? `
    <div class="modal-field">
      <div class="modal-field-label">Projekt</div>
      <div class="modal-field-value">${escapeHtml(besuch.projekt)}</div>
    </div>` : ''}

    ${fotos.length > 0 ? `
    <div class="modal-field">
      <div class="modal-field-label">Fotos</div>
      <div class="fotos-grid">
        ${fotos.map(f => {
          const url = typeof f === 'string' ? f : f.url;
          const thumb = typeof f === 'string' ? f : (f.thumbnail || f.url);
          return `<div class="foto-thumb">
            <img src="${thumb}" onclick="openLightbox('${url}')" alt="Foto">
          </div>`;
        }).join('')}
      </div>
    </div>` : ''}

    <div class="modal-actions">
      <button class="btn btn-outline" onclick="besuchBearbeiten('${besuch.id}')">\u270F\uFE0F Bearbeiten</button>
      <button class="btn btn-danger" onclick="besuchLoeschen('${besuch.id}')">\u{1F5D1} Loeschen</button>
    </div>
  `;

  createModal(content);
}

function besuchBearbeiten(id) {
  // Modal schliessen
  document.querySelector('.modal-overlay')?.remove();
  window.location.hash = `#/neuer-besuch/edit/${id}`;
}

async function besuchLoeschen(id) {
  // Modal schliessen
  document.querySelector('.modal-overlay')?.remove();

  const confirmed = await confirmDialog('Wirklich loeschen? Dieser Besuch kann nicht wiederhergestellt werden.');
  if (!confirmed) return;

  try {
    await API.besuche.loeschen(id);
    showToast('Besuch geloescht', 'success');
    await ladeBesucheListe();
  } catch (error) {
    showToast('Loeschen fehlgeschlagen: ' + error.message, 'error');
  }
}
