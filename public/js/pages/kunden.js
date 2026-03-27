// ========================================
// Kunden-Seite (Liste + Detail)
// Branchenspezifische Icons und Besuchszaehlung
// ========================================

// Branchen-Icon basierend auf Segment oder Firmenname
function getBranchenIcon(firma, segment) {
  // Segment-basiert
  if (segment === 'Architekt') return '\u{1F4D0}';
  if (segment === 'Maler') return '\u{1F3A8}';
  if (segment === 'Schreiner') return '\u{1FAB5}';
  if (segment === 'Zimmerei') return '\u{1F3D7}';
  if (segment === 'Fensterbau') return '\u{1FA9F}';
  if (segment === 'Fassadenbau') return '\u{1F3E2}';
  if (segment === 'Handel') return '\u{1F3EA}';
  if (segment === 'Generalunternehmer') return '\u{1F3DB}';

  // Firmenname-basiert (Fallback)
  if (firma) {
    const f = firma.toLowerCase();
    if (f.includes('architekt') || f.includes('atelier')) return '\u{1F4D0}';
    if (f.includes('maler') || f.includes('malergesch') || f.includes('art of')) return '\u{1F3A8}';
    if (f.includes('schreiner') || f.includes('holz') || f.includes('holzbau')) return '\u{1FAB5}';
    if (f.includes('zimmer')) return '\u{1F3D7}';
    if (f.includes('fenster') || f.includes('klarer')) return '\u{1FA9F}';
    if (f.includes('fassade')) return '\u{1F3E2}';
    if (f.includes('kabe') || f.includes('sika')) return '\u{1F3EA}';
  }

  return '\u{1F3E2}'; // Default
}

// Pastell-Farbe fuer das Branchen-Icon
function getBranchenFarbe(firma, segment) {
  if (segment === 'Architekt') return '#f0e6ff';
  if (segment === 'Maler') return '#e6f0ff';
  if (segment === 'Schreiner') return '#e6ffe6';
  if (segment === 'Zimmerei') return '#fff8e6';
  if (segment === 'Fensterbau') return '#e6fffe';
  if (segment === 'Fassadenbau') return '#fff0e6';
  if (segment === 'Handel') return '#ffe6f0';
  if (segment === 'Generalunternehmer') return '#f0e6ff';

  if (firma) {
    const f = firma.toLowerCase();
    if (f.includes('architekt') || f.includes('atelier')) return '#f0e6ff';
    if (f.includes('maler') || f.includes('malergesch') || f.includes('art of')) return '#e6f0ff';
    if (f.includes('schreiner') || f.includes('holz') || f.includes('holzbau')) return '#e6ffe6';
    if (f.includes('zimmer')) return '#fff8e6';
    if (f.includes('fenster') || f.includes('klarer')) return '#e6fffe';
    if (f.includes('fassade')) return '#fff0e6';
    if (f.includes('kabe') || f.includes('sika')) return '#ffe6f0';
  }

  return '#f5f0e8'; // Default Beige
}

// Globaler State fuer Kunden-Sortierung und Filter
let kundenState = {
  sortBy: 'alpha',      // 'alpha', 'besuche', 'segment'
  filterSegment: '',     // '' = alle
  alleKunden: [],        // Cache fuer Sortierung ohne Neuladen
  alleBesuche: []
};

async function renderKunden(params) {
  const main = document.getElementById('mainContent');

  // Kunden-Detailansicht
  if (params && params.length > 0 && !params.startsWith('edit')) {
    await renderKundeDetail(params);
    return;
  }

  main.innerHTML = `
    <div class="page-header">
      <div>
        <h1>Kundenverzeichnis</h1>
        <p class="subtitle">Alle erfassten Kunden und ihre Historie</p>
      </div>
    </div>

    <div class="search-bar" style="display:flex; gap:12px; flex-wrap:wrap; align-items:center;">
      <input type="text" class="search-input" id="kundenSuche" placeholder="Nach Firmenname oder Ort suchen..." onkeyup="kundenFiltern()" style="flex:1; min-width:200px;">

      <select id="kundenSegmentFilter" onchange="kundenFiltern()" style="padding:10px 14px; border:1px solid #e8e0d6; border-radius:12px; font-size:14px; background:#fff; color:#3d3024; min-width:160px;">
        <option value="">Alle Branchen</option>
        <option value="Schreiner">🪵 Schreiner</option>
        <option value="Maler">🎨 Maler</option>
        <option value="Zimmerei">🏗️ Zimmerei</option>
        <option value="Fensterbau">🪟 Fensterbau</option>
        <option value="Fassadenbau">🏢 Fassadenbau</option>
        <option value="Architekt">📐 Architekt</option>
        <option value="Generalunternehmer">🏛️ Generalunternehmer</option>
        <option value="Handel">🏪 Handel</option>
        <option value="Sonstiges">Sonstiges</option>
      </select>

      <select id="kundenSortierung" onchange="kundenFiltern()" style="padding:10px 14px; border:1px solid #e8e0d6; border-radius:12px; font-size:14px; background:#fff; color:#3d3024; min-width:160px;">
        <option value="alpha">A–Z Alphabetisch</option>
        <option value="alpha_desc">Z–A Alphabetisch</option>
        <option value="besuche">Meiste Besuche</option>
        <option value="segment">Nach Branche</option>
      </select>
    </div>

    <div id="kundenZaehler" style="padding:8px 0; color:#8a7a6a; font-size:14px;"></div>

    <div id="kundenListe">
      <div class="loading"><span class="spinner"></span> Kunden laden...</div>
    </div>
  `;

  await ladeKundenListe();
}

async function ladeKundenListe(filter = {}) {
  try {
    // Kunden und Besuche parallel laden
    const [kunden, besuche] = await Promise.all([
      API.kunden.laden(filter),
      API.besuche.laden()
    ]);

    // Besuchsanzahl per Firmenname zaehlen
    const besuchsMap = {};
    (besuche || []).forEach(b => {
      if (b.firma) {
        const key = b.firma.trim().toLowerCase();
        besuchsMap[key] = (besuchsMap[key] || 0) + 1;
      }
    });

    // Besuchsanzahl an Kunden zuweisen
    kunden.forEach(k => {
      const key = (k.firma || '').trim().toLowerCase();
      const matchedCount = besuchsMap[key] || 0;
      k.besucheAnzahl = Math.max(k.besucheAnzahl || 0, matchedCount);

      // Segment aus Firmenname ableiten wenn nicht gesetzt
      if (!k.segment) {
        k.segment = detectSegment(k.firma);
      }
    });

    // Im State cachen fuer Sortierung/Filter ohne Neuladen
    kundenState.alleKunden = kunden;
    kundenState.alleBesuche = besuche;

    kundenFiltern();
  } catch (error) {
    document.getElementById('kundenListe').innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">\u26A0\uFE0F</div>
        <p class="empty-state-text">Fehler beim Laden: ${error.message}</p>
      </div>
    `;
  }
}

// Segment aus Firmenname erkennen
function detectSegment(firma) {
  if (!firma) return '';
  const f = firma.toLowerCase();
  if (f.includes('architekt') || f.includes('atelier') || f.includes('ing.') || f.includes('planung')) return 'Architekt';
  if (f.includes('maler') || f.includes('malergesch') || f.includes('art of') || f.includes('anstrich')) return 'Maler';
  if (f.includes('schreiner') || f.includes('schreinerei')) return 'Schreiner';
  if (f.includes('zimmer') || f.includes('holzbau') || f.includes('holz ')) return 'Zimmerei';
  if (f.includes('fenster')) return 'Fensterbau';
  if (f.includes('fassade')) return 'Fassadenbau';
  if (f.includes('kabe') || f.includes('sika') || f.includes('handel') || f.includes(' ag,') || f.includes('farben')) return 'Handel';
  return '';
}

// Filtern und Sortieren (ohne API-Call, aus Cache)
function kundenFiltern() {
  const suche = (document.getElementById('kundenSuche')?.value || '').trim().toLowerCase();
  const segment = document.getElementById('kundenSegmentFilter')?.value || '';
  const sortBy = document.getElementById('kundenSortierung')?.value || 'alpha';

  let kunden = [...kundenState.alleKunden];

  // Textsuche
  if (suche) {
    kunden = kunden.filter(k => {
      const text = [k.firma, k.ort, k.hauptkontakt, k.segment].filter(Boolean).join(' ').toLowerCase();
      return text.includes(suche);
    });
  }

  // Segment-Filter
  if (segment) {
    kunden = kunden.filter(k => {
      // Geprueft gegen gesetztes Segment oder aus Name erkanntes
      const kundeSeg = k.segment || detectSegment(k.firma);
      return kundeSeg === segment;
    });
  }

  // Sortierung
  const segmentReihenfolge = ['Schreiner', 'Maler', 'Zimmerei', 'Fensterbau', 'Fassadenbau', 'Architekt', 'Generalunternehmer', 'Handel', 'Sonstiges', ''];

  switch (sortBy) {
    case 'alpha':
      kunden.sort((a, b) => (a.firma || '').localeCompare(b.firma || '', 'de'));
      break;
    case 'alpha_desc':
      kunden.sort((a, b) => (b.firma || '').localeCompare(a.firma || '', 'de'));
      break;
    case 'besuche':
      kunden.sort((a, b) => (b.besucheAnzahl || 0) - (a.besucheAnzahl || 0));
      break;
    case 'segment':
      kunden.sort((a, b) => {
        const segA = a.segment || detectSegment(a.firma) || '';
        const segB = b.segment || detectSegment(b.firma) || '';
        const idxA = segmentReihenfolge.indexOf(segA);
        const idxB = segmentReihenfolge.indexOf(segB);
        if (idxA !== idxB) return (idxA === -1 ? 999 : idxA) - (idxB === -1 ? 999 : idxB);
        return (a.firma || '').localeCompare(b.firma || '', 'de');
      });
      break;
  }

  // Zaehler aktualisieren
  const zaehler = document.getElementById('kundenZaehler');
  if (zaehler) {
    const total = kundenState.alleKunden.length;
    const filtered = kunden.length;
    zaehler.textContent = filtered === total
      ? `${total} Kunden`
      : `${filtered} von ${total} Kunden`;
  }

  renderKundenGrid(kunden);
}

function renderKundenGrid(kunden) {
  const container = document.getElementById('kundenListe');

  if (kunden.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">\u{1F465}</div>
        <p class="empty-state-text">Keine Kunden gefunden</p>
      </div>
    `;
    return;
  }

  container.innerHTML = `
    <div class="kunden-grid">
      ${kunden.map(k => {
        const icon = getBranchenIcon(k.firma, k.segment);
        const farbe = getBranchenFarbe(k.firma, k.segment);
        return `
        <div class="kunde-card" onclick="window.location.hash='#/kunden/${k.id}'">
          <div class="branchen-icon-circle" style="background:${farbe};">${icon}</div>
          <div class="kunde-card-firma">${escapeHtml(k.firma)}</div>
          <div class="kunde-card-info">
            ${k.ort ? `<span>\u{1F4CD} ${escapeHtml(k.ort)}</span>` : ''}
            ${k.hauptkontakt ? `<span>\u{1F464} Kontakt: ${escapeHtml(k.hauptkontakt)}</span>` : ''}
            ${k.segment ? `<span>\u{1F3F7} ${escapeHtml(k.segment)}</span>` : ''}
          </div>
          <div class="kunde-card-besuche">${k.besucheAnzahl || 0} Besuche</div>
        </div>
      `;}).join('')}
    </div>
  `;
}

// kundenSuchen wird nicht mehr benoetigt - ersetzt durch kundenFiltern()

// ---- Kunden-Detail ----

async function renderKundeDetail(kundeId) {
  const main = document.getElementById('mainContent');
  main.innerHTML = '<div class="loading"><span class="spinner"></span> Kundendaten laden...</div>';

  try {
    // Kunde laden und parallel alle Besuche laden (fuer Frontend-Match)
    const [kunde, alleBesuche] = await Promise.all([
      API.kunden.einzeln(kundeId),
      API.besuche.laden()
    ]);

    // Wenn die Notion-Relation leer ist, Besuche per Firmenname matchen
    if (!kunde.besuche || kunde.besuche.length === 0) {
      const kundeNameLower = (kunde.firma || '').trim().toLowerCase();
      kunde.besuche = (alleBesuche || []).filter(b => {
        return (b.firma || '').trim().toLowerCase() === kundeNameLower;
      });
      kunde.besucheAnzahl = kunde.besuche.length;
    }

    const icon = getBranchenIcon(kunde.firma, kunde.segment);
    const farbe = getBranchenFarbe(kunde.firma, kunde.segment);

    main.innerHTML = `
      <div class="kunde-detail">
        <a class="kunde-detail-back" onclick="window.location.hash='#/kunden'">\u2190 Zurueck zur Liste</a>

        <div class="page-header">
          <h1><span class="branchen-icon-circle" style="background:${farbe}; display:inline-flex; width:40px; height:40px; font-size:20px; vertical-align:middle; margin-right:8px;">${icon}</span> ${escapeHtml(kunde.firma)}</h1>
          <a href="#/neuer-besuch/kunde/${kundeId}" class="btn btn-orange">+ Bericht fuer Kunde</a>
        </div>

        <div class="kunde-detail-grid">
          <!-- Links: Kundendaten -->
          <div class="card">
            <div class="card-title flex justify-between items-center">
              <span>Kundendaten</span>
              <button class="btn btn-outline btn-sm" onclick="kundeBearbeiten('${kundeId}')">\u270F\uFE0F Bearbeiten</button>
            </div>
            <div class="kunde-info-grid mt-16">
              <div class="kunde-info-item">
                <div class="label">Segment</div>
                <div>${escapeHtml(kunde.segment || '-')}</div>
              </div>
              <div class="kunde-info-item">
                <div class="label">Ort</div>
                <div>${escapeHtml(kunde.ort || '-')}</div>
              </div>
              <div class="kunde-info-item">
                <div class="label">Hauptkontakt</div>
                <div>${escapeHtml(kunde.hauptkontakt || '-')}</div>
              </div>
              <div class="kunde-info-item">
                <div class="label">Telefon</div>
                <div>${escapeHtml(kunde.telefon || '-')}</div>
              </div>
              <div class="kunde-info-item">
                <div class="label">Email</div>
                <div>${kunde.email ? `<a href="mailto:${kunde.email}">${escapeHtml(kunde.email)}</a>` : '-'}</div>
              </div>
              <div class="kunde-info-item">
                <div class="label">Status</div>
                <div>${escapeHtml(kunde.status || '-')}</div>
              </div>
              ${kunde.webseite ? `
              <div class="kunde-info-item">
                <div class="label">Webseite</div>
                <div><a href="${kunde.webseite}" target="_blank">${escapeHtml(kunde.webseite)}</a></div>
              </div>` : ''}
              ${kunde.notizen ? `
              <div class="kunde-info-item" style="grid-column: span 2;">
                <div class="label">Notizen</div>
                <div>${escapeHtml(kunde.notizen)}</div>
              </div>` : ''}
            </div>
          </div>

          <!-- Rechts: Besuchshistorie -->
          <div class="card">
            <div class="card-title">Besuchshistorie (${kunde.besuche ? kunde.besuche.length : 0})</div>
            ${(!kunde.besuche || kunde.besuche.length === 0) ? `
              <div class="empty-state" style="padding:20px;">
                <p class="empty-state-text">Noch keine Besuche erfasst</p>
              </div>
            ` : `
              <ul class="activity-list mt-16">
                ${kunde.besuche.map(b => `
                  <li class="activity-item" style="cursor:pointer;" onclick="window.location.hash='#/besuche'; setTimeout(() => zeigeBesuchDetail('${b.id}'), 500);">
                    <div class="activity-icon">\u{1F4CB}</div>
                    <div class="activity-text">
                      <div>
                        <strong>${formatDatumKurz(b.datum)}</strong>
                        ${b.besuchstyp ? `<span class="badge badge-typ" style="font-size:10px; padding:2px 6px; margin-left:6px;">${escapeHtml(b.besuchstyp)}</span>` : ''}
                      </div>
                      ${b.themen ? `<div style="font-size:13px; color:var(--text-light); margin-top:4px;">${escapeHtml(b.themen.substring(0, 100))}${b.themen.length > 100 ? '...' : ''}</div>` : ''}
                    </div>
                  </li>
                `).join('')}
              </ul>
            `}
          </div>
        </div>
      </div>
    `;

  } catch (error) {
    main.innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">\u26A0\uFE0F</div>
        <p class="empty-state-text">Kunde konnte nicht geladen werden: ${error.message}</p>
        <a href="#/kunden" class="btn btn-outline mt-16">Zurueck zur Liste</a>
      </div>
    `;
  }
}

// ---- Kunde bearbeiten ----

async function kundeBearbeiten(kundeId) {
  let kunde;
  try {
    kunde = await API.kunden.einzeln(kundeId);
  } catch (e) {
    showToast('Kunde konnte nicht geladen werden', 'error');
    return;
  }

  const content = `
    <div class="modal-title">Kunde bearbeiten</div>
    <div class="form-row">
      <div class="form-group">
        <label>Firma</label>
        <input type="text" id="editKundeFirma" value="${escapeHtml(kunde.firma || '')}">
      </div>
      <div class="form-group">
        <label>Ort</label>
        <input type="text" id="editKundeOrt" value="${escapeHtml(kunde.ort || '')}">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Hauptkontakt</label>
        <input type="text" id="editKundeKontakt" value="${escapeHtml(kunde.hauptkontakt || '')}">
      </div>
      <div class="form-group">
        <label>Segment</label>
        <select id="editKundeSegment">
          <option value="">-- Waehlen --</option>
          ${selectOptions(SEGMENTE, kunde.segment || '')}
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Telefon</label>
        <input type="text" id="editKundeTelefon" value="${escapeHtml(kunde.telefon || '')}">
      </div>
      <div class="form-group">
        <label>Email</label>
        <input type="email" id="editKundeEmail" value="${escapeHtml(kunde.email || '')}">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Webseite</label>
        <input type="url" id="editKundeWeb" value="${escapeHtml(kunde.webseite || '')}">
      </div>
      <div class="form-group">
        <label>Status</label>
        <select id="editKundeStatus">
          ${selectOptions(STATUS_OPTIONEN, kunde.status || 'Aktiv')}
        </select>
      </div>
    </div>
    <div class="form-group mb-16">
      <label>Notizen</label>
      <textarea id="editKundeNotizen" rows="3">${escapeHtml(kunde.notizen || '')}</textarea>
    </div>
    <div class="flex gap-12" style="justify-content:flex-end;">
      <button class="btn btn-outline" onclick="this.closest('.modal-overlay').remove()">Abbrechen</button>
      <button class="btn btn-green" onclick="kundeAktualisieren('${kundeId}')">\u{1F4BE} Speichern</button>
    </div>
  `;

  createModal(content);
}

async function kundeAktualisieren(kundeId) {
  const daten = {
    id: kundeId,
    firma: document.getElementById('editKundeFirma').value.trim(),
    ort: document.getElementById('editKundeOrt').value.trim(),
    hauptkontakt: document.getElementById('editKundeKontakt').value.trim(),
    segment: document.getElementById('editKundeSegment').value,
    telefon: document.getElementById('editKundeTelefon').value.trim(),
    email: document.getElementById('editKundeEmail').value.trim(),
    webseite: document.getElementById('editKundeWeb').value.trim(),
    status: document.getElementById('editKundeStatus').value,
    notizen: document.getElementById('editKundeNotizen').value.trim()
  };

  try {
    await API.kunden.bearbeiten(daten);
    document.querySelector('.modal-overlay')?.remove();
    showToast('Kunde aktualisiert!', 'success');
    // Seite neu laden
    await renderKundeDetail(kundeId);
  } catch (error) {
    showToast('Speichern fehlgeschlagen: ' + error.message, 'error');
  }
}
