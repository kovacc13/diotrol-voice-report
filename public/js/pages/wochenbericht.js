// ========================================
// Wochenbericht-Seite
// Mit To-Do-Split (Daniel / Marco)
// und Segment-Ranking
// ========================================

let wochenberichtData = null;

// To-Do Status (Checkboxen) - wird pro Session im Speicher gehalten
let todoStatus = {};

// Besuchs-Sortierung innerhalb eines Tages:
// 1. Office-Eintraege zuerst
// 2. Normale Kundenbesuche in der Mitte
// 3. Interne Eintraege (Diotrol-Intern, Home-Office) zuletzt
function besuchPrioritaet(b) {
  const firma = (b.firma || '').toLowerCase();
  if (/office|buero|büro/i.test(firma)) return 0;       // Office zuerst
  if (/intern|home.?office|dio/i.test(firma)) return 2;  // Intern zuletzt
  return 1;                                               // Kunden in der Mitte
}

function sortBesuche(besuche) {
  return [...besuche].sort((a, b) => besuchPrioritaet(a) - besuchPrioritaet(b));
}

async function renderWochenbericht() {
  const main = document.getElementById('mainContent');
  const kw = getCurrentKW();
  const jahr = getCurrentJahr();

  main.innerHTML = `
    <div class="page-header">
      <div>
        <h1>Wochenbericht</h1>
        <p class="subtitle">Zusammenfassung Ihrer Vertriebsaktivitaeten</p>
      </div>
      <div class="flex gap-8">
        <select class="filter-select" id="wbKWSelect" onchange="wochenberichtLaden()">
          <option value="${kw}">KW ${kw}</option>
        </select>
        <div class="dropdown">
          <button class="btn btn-dark" onclick="toggleExportDropdown()">📤 Exportieren ▾</button>
          <div class="dropdown-content" id="exportDropdown">
            <button class="dropdown-item" onclick="exportPDF()">📄 PDF herunterladen (.pdf)</button>
            <button class="dropdown-item" onclick="exportWord()">📘 Word herunterladen (.docx)</button>
            <button class="dropdown-item" onclick="exportText()">📝 Text herunterladen (.txt)</button>
            <button class="dropdown-item" onclick="exportEmail()">✉️ Per E-Mail senden</button>
            <hr style="margin:4px 0; border:none; border-top:1px solid #e5e7eb;">
            <button class="dropdown-item" onclick="exportMarcoReport()">👔 Marco-Bericht senden</button>
          </div>
        </div>
      </div>
    </div>

    <div class="stat-cards" id="wbStats">
      <a class="stat-card" href="#/besuche" title="Alle Besuche anzeigen">
        <div class="stat-icon">\u{1F4CB}</div>
        <div class="stat-info">
          <div class="stat-label">Total Besuche</div>
          <div class="stat-value" id="wbTotal">-</div>
        </div>
      </a>
      <a class="stat-card" href="#/kunden" title="Neukunden anzeigen">
        <div class="stat-icon">\u{1F195}</div>
        <div class="stat-info">
          <div class="stat-label">Neukunden</div>
          <div class="stat-value" id="wbNeukunden">-</div>
        </div>
      </a>
      <a class="stat-card" href="#/kunden" title="Bestandskunden anzeigen">
        <div class="stat-icon">\u{1F504}</div>
        <div class="stat-info">
          <div class="stat-label">Bestandskunden</div>
          <div class="stat-value" id="wbBestandskunden">-</div>
        </div>
      </a>
      <a class="stat-card" href="#/kunden" title="Besuchte Kunden anzeigen">
        <div class="stat-icon">\u{1F465}</div>
        <div class="stat-info">
          <div class="stat-label">Kunden besucht</div>
          <div class="stat-value" id="wbKundenBesucht">-</div>
        </div>
      </a>
    </div>

    <!-- Segment-Ranking -->
    <div class="card" id="wbSegmentRanking" style="margin-bottom:16px;">
      <h3 class="card-title">📊 Segmentverteilung</h3>
      <div id="segmentRankingContent">
        <div class="loading"><span class="spinner"></span> Laden...</div>
      </div>
    </div>

    <!-- To-Do Listen (Daniel + Marco) -->
    <div id="wbTodoSection" style="display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-bottom:16px;">
      <div class="card" id="wbTodosDaniel">
        <div class="todo-header-daniel">
          <h3 class="card-title" style="margin:0;">⭐ Meine To-Dos (Daniel)</h3>
          <span class="todo-counter" id="todoDanielCounter">-</span>
        </div>
        <div id="todoDanielContent">
          <div class="loading"><span class="spinner"></span> Laden...</div>
        </div>
        <div class="todo-actions" id="todoDanielActions" style="display:none;">
          <button class="btn btn-sm btn-outline" onclick="toggleErledigte('daniel')">Erledigte anzeigen</button>
          <button class="btn btn-sm btn-outline" onclick="copyTodos('daniel')">📋 Kopieren</button>
          <button class="btn btn-sm btn-outline" onclick="downloadTodos('daniel')">💾 Download</button>
        </div>
      </div>
      <div class="card" id="wbTodosMarco">
        <div class="todo-header-marco">
          <h3 class="card-title" style="margin:0;">👔 Aufgaben fuer Marco</h3>
          <span class="todo-counter" id="todoMarcoCounter">-</span>
        </div>
        <div id="todoMarcoContent">
          <div class="loading"><span class="spinner"></span> Laden...</div>
        </div>
        <div class="todo-actions" id="todoMarcoActions" style="display:none;">
          <button class="btn btn-sm btn-outline" onclick="toggleErledigte('marco')">Erledigte anzeigen</button>
          <button class="btn btn-sm btn-outline" onclick="copyTodos('marco')">📋 Kopieren</button>
          <button class="btn btn-sm btn-outline" onclick="downloadTodos('marco')">💾 Download</button>
          <button class="btn btn-sm btn-orange" onclick="emailTodos('marco')">✉️ An Marco senden</button>
        </div>
      </div>
    </div>

    <h3 class="mb-16">📅 Details der Woche</h3>
    <div id="wbBesuche">
      <div class="loading"><span class="spinner"></span> Wochenbericht laden...</div>
    </div>
  `;

  await wochenberichtLaden();
}

async function wochenberichtLaden() {
  const kwSelect = document.getElementById('wbKWSelect');
  const kw = kwSelect ? kwSelect.value : getCurrentKW();
  const jahr = getCurrentJahr();

  try {
    const data = await API.wochenbericht.laden(kw, jahr);
    wochenberichtData = data;

    // KW-Dropdown mit verfuegbaren KWs befuellen
    if (data.verfuegbareKWs && data.verfuegbareKWs.length > 0 && kwSelect) {
      const current = parseInt(kwSelect.value);
      kwSelect.innerHTML = data.verfuegbareKWs.map(k =>
        `<option value="${k}" ${k === current ? 'selected' : ''}>KW ${k}</option>`
      ).join('');
      if (!data.verfuegbareKWs.includes(getCurrentKW())) {
        kwSelect.innerHTML = `<option value="${getCurrentKW()}">KW ${getCurrentKW()} (aktuell)</option>` + kwSelect.innerHTML;
      }
    }

    // Stats aktualisieren
    const stats = data.stats || {};
    document.getElementById('wbTotal').textContent = stats.totalBesuche || 0;
    document.getElementById('wbNeukunden').textContent = stats.neukunden || 0;
    document.getElementById('wbBestandskunden').textContent = stats.bestandskunden || 0;
    document.getElementById('wbKundenBesucht').textContent = stats.firmenBesucht || 0;

    // Segment-Ranking rendern
    renderSegmentRanking(data);

    // To-Do Listen rendern
    renderTodoListen(data);

    // Besuche rendern
    renderWochenberichtBesuche(data.besuche || []);

  } catch (error) {
    document.getElementById('wbBesuche').innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">⚠️</div>
        <p class="empty-state-text">Wochenbericht konnte nicht geladen werden: ${error.message}</p>
      </div>
    `;
  }
}

// ============================================================
// SEGMENT-RANKING
// ============================================================
function renderSegmentRanking(data) {
  const container = document.getElementById('segmentRankingContent');
  const besuche = data.besuche || [];

  if (besuche.length === 0) {
    container.innerHTML = '<p class="muted">Keine Daten</p>';
    return;
  }

  // Segmentverteilung berechnen (auch client-seitig fuer Exporte)
  const segmente = {};
  besuche.forEach(b => {
    const seg = b.segment || 'Sonstiges';
    segmente[seg] = (segmente[seg] || 0) + 1;
  });

  const ranking = Object.entries(segmente)
    .sort((a, b) => b[1] - a[1]);

  const maxCount = ranking[0] ? ranking[0][1] : 1;

  // Segment-Farben fuer die Balken
  const SEG_FARBEN = {
    'Architekt': '#5b7fa6', 'Maler': '#b87a6b', 'Zimmerei': '#6a9b6a',
    'Fensterbau': '#8b7baa', 'Schreiner': '#a89060', 'Holzbau': '#6a9b9b',
    'Holzbauingenieure': '#5a8a7a', 'Hobelwerke': '#8a7a5a', 'Saegereien': '#7a6a5a',
    'Handel': '#c4956a',
    'Diotrol-Intern': '#4a7a6a', 'Office': '#6a8a9b', 'Sonstiges': '#8a8a8a'
  };

  container.innerHTML = `
    <div class="segment-ranking-list">
      ${ranking.map(([seg, cnt], idx) => {
        const prozent = Math.round(cnt / besuche.length * 100);
        const balkenBreite = Math.round(cnt / maxCount * 100);
        const farbe = SEG_FARBEN[seg] || '#6b7280';
        const icon = SEGMENT_ICONS[seg] || 'X';
        return `
          <div class="segment-ranking-row segment-ranking-clickable" onclick="filterBesucheNachSegment('${seg}')" title="Besuche fuer ${seg} anzeigen">
            <div class="segment-ranking-rank">${idx + 1}.</div>
            <div class="segment-ranking-icon" style="background:${farbe};">${icon}</div>
            <div class="segment-ranking-info">
              <div class="segment-ranking-name">${seg}</div>
              <div class="segment-ranking-bar-bg">
                <div class="segment-ranking-bar" style="width:${balkenBreite}%; background:${farbe};"></div>
              </div>
            </div>
            <div class="segment-ranking-count">${cnt}</div>
            <div class="segment-ranking-pct">${prozent}%</div>
          </div>
        `;
      }).join('')}
    </div>
  `;
}

// ============================================================
// TO-DO LISTEN (Daniel + Marco)
// ============================================================
// Erkennung ob ein To-Do fuer Marco ist (auch in alten Daten ohne todosMarco-Feld)
function istMarcoTodo(text) {
  const lower = text.toLowerCase();
  // Signalwoerter: Name "Marco" am Anfang oder als Subjekt, Chef-Bezug
  const marcoPatterns = [
    /\bmarco\b/i,
    /\bmh\b/i,
    /\bchef\b/i,
    /\bgeschaeftsleitung\b/i,
    /\bvorgesetzt/i,
    /\bfreigabe\b/i,
    /\bbudget\s*(genehmig|freigab|anford)/i,
    /\bintern\s+besprechen\b/i,
    /\bmanagement\b/i
  ];
  return marcoPatterns.some(p => p.test(text));
}

function renderTodoListen(data) {
  const besuche = data.besuche || [];

  // To-Dos sammeln und intelligent zuordnen
  const todosDaniel = [];
  const todosMarco = [];

  besuche.forEach(b => {
    if (b.naechsteSchritte) {
      // Jede Zeile als separates To-Do
      const lines = b.naechsteSchritte.split('\n')
        .map(l => l.trim().replace(/^[\u2022\-\*]\s*/, ''))
        .filter(l => l.length > 0);
      lines.forEach(line => {
        // Intelligente Zuordnung: Wenn "Marco" oder Chef-Signalwoerter drin stehen,
        // gehoert es in die Marco-Liste (auch bei alten Daten ohne todosMarco-Feld)
        if (istMarcoTodo(line)) {
          todosMarco.push({
            id: `m-${b.id}-${todosMarco.length}`,
            firma: b.firma || 'Intern',
            text: line,
            datum: b.datum,
            besuchstyp: b.besuchstyp || b.typ || 'Besuch'
          });
        } else {
          todosDaniel.push({
            id: `d-${b.id}-${todosDaniel.length}`,
            firma: b.firma || 'Intern',
            text: line,
            datum: b.datum,
            besuchstyp: b.besuchstyp || b.typ || 'Besuch'
          });
        }
      });
    }
    // Explizites todosMarco-Feld (neue Besuche nach dem Update)
    if (b.todosMarco) {
      const lines = b.todosMarco.split('\n')
        .map(l => l.trim().replace(/^[\u2022\-\*]\s*/, ''))
        .filter(l => l.length > 0);
      lines.forEach(line => {
        todosMarco.push({
          id: `m-${b.id}-${todosMarco.length}`,
          firma: b.firma || 'Intern',
          text: line,
          datum: b.datum,
          besuchstyp: b.besuchstyp || b.typ || 'Besuch'
        });
      });
    }
  });

  renderTodoList('daniel', todosDaniel);
  renderTodoList('marco', todosMarco);
}

function renderTodoList(typ, todos) {
  const container = document.getElementById(typ === 'daniel' ? 'todoDanielContent' : 'todoMarcoContent');
  const counter = document.getElementById(typ === 'daniel' ? 'todoDanielCounter' : 'todoMarcoCounter');
  const actions = document.getElementById(typ === 'daniel' ? 'todoDanielActions' : 'todoMarcoActions');

  const offene = todos.filter(t => !todoStatus[t.id]);
  const erledigte = todos.filter(t => todoStatus[t.id]);

  counter.textContent = `${offene.length} offen` + (erledigte.length > 0 ? ` · ${erledigte.length} erledigt` : '');

  if (todos.length === 0) {
    container.innerHTML = `
      <div class="empty-state" style="padding:16px;">
        <p class="empty-state-text" style="font-size:13px;">Keine ${typ === 'daniel' ? 'eigenen' : 'Marco-'}Aufgaben in dieser KW</p>
      </div>
    `;
    actions.style.display = 'none';
    return;
  }

  actions.style.display = 'flex';

  // Zeige nur offene, ausser "Erledigte anzeigen" ist aktiv
  const showErledigte = container.dataset.showErledigte === 'true';
  const visibleTodos = showErledigte ? todos : offene;

  container.innerHTML = `
    <div class="todo-checklist">
      ${visibleTodos.map(t => {
        const checked = todoStatus[t.id] ? 'checked' : '';
        const strikeClass = todoStatus[t.id] ? 'todo-done' : '';
        return `
          <label class="todo-item ${strikeClass}">
            <input type="checkbox" ${checked} onchange="toggleTodo('${t.id}', '${typ}')">
            <span class="todo-text">
              <strong>[${escapeHtml(t.firma)}]:</strong> ${escapeHtml(t.text)}
              <span class="todo-meta">(Quelle: ${escapeHtml(t.besuchstyp)} vom ${formatDatumKurz(t.datum)})</span>
            </span>
          </label>
        `;
      }).join('')}
    </div>
  `;

  // Speichere Todos fuer Export
  container.dataset.todos = JSON.stringify(todos);
}

function toggleTodo(todoId, typ) {
  todoStatus[todoId] = !todoStatus[todoId];
  // Re-render nur die betroffene Liste
  if (wochenberichtData) renderTodoListen(wochenberichtData);
}

function toggleErledigte(typ) {
  const container = document.getElementById(typ === 'daniel' ? 'todoDanielContent' : 'todoMarcoContent');
  const current = container.dataset.showErledigte === 'true';
  container.dataset.showErledigte = !current;
  if (wochenberichtData) renderTodoListen(wochenberichtData);
}

function getTodosFromContainer(typ) {
  const container = document.getElementById(typ === 'daniel' ? 'todoDanielContent' : 'todoMarcoContent');
  try {
    return JSON.parse(container.dataset.todos || '[]');
  } catch { return []; }
}

function copyTodos(typ) {
  const todos = getTodosFromContainer(typ);
  const label = typ === 'daniel' ? 'Meine To-Dos (Daniel)' : 'Aufgaben fuer Marco';
  let text = `${label} - KW ${wochenberichtData?.kw || ''}\n${'='.repeat(40)}\n\n`;
  todos.forEach(t => {
    const status = todoStatus[t.id] ? '[x]' : '[ ]';
    text += `${status} [${t.firma}]: ${t.text}\n`;
  });
  navigator.clipboard.writeText(text).then(() => {
    showToast(`${label} kopiert!`, 'success');
  }).catch(() => {
    showToast('Kopieren fehlgeschlagen', 'error');
  });
}

function downloadTodos(typ) {
  const todos = getTodosFromContainer(typ);
  const label = typ === 'daniel' ? 'Meine_ToDos_Daniel' : 'Aufgaben_Marco';
  let text = `${typ === 'daniel' ? 'MEINE TO-DOs (Daniel Pfister)' : 'AUFGABEN FUER MARCO'}\n`;
  text += `KW ${wochenberichtData?.kw || ''} / ${wochenberichtData?.jahr || ''}\n`;
  text += `${'='.repeat(50)}\n\n`;
  todos.forEach(t => {
    const status = todoStatus[t.id] ? '[x]' : '[ ]';
    text += `${status} [${t.firma}]: ${t.text}\n    Quelle: ${t.besuchstyp} vom ${formatDatumKurz(t.datum)}\n\n`;
  });
  const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
  downloadBlob(blob, `${label}_KW${wochenberichtData?.kw || ''}.txt`);
  showToast('To-Do-Liste heruntergeladen', 'success');
}

function emailTodos(typ) {
  const todos = getTodosFromContainer(typ);
  const kw = wochenberichtData?.kw || '';
  const jahr = wochenberichtData?.jahr || '';
  let text = `Aufgaben aus Wochenbericht KW ${kw} / ${jahr}\n\n`;
  todos.forEach(t => {
    text += `[ ] [${t.firma}]: ${t.text}\n`;
  });
  text += `\n---\nErstellt aus dem Wochenbericht von Daniel Pfister`;
  const subject = encodeURIComponent(`Aufgaben aus Wochenbericht KW ${kw}`);
  const body = encodeURIComponent(text);
  window.open(`mailto:?subject=${subject}&body=${body}`, '_self');
}

// ============================================================
// BESUCHS-DETAIL-ANSICHT
// ============================================================
function renderWochenberichtBesuche(besuche) {
  const container = document.getElementById('wbBesuche');

  if (besuche.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">📅</div>
        <p class="empty-state-text">Keine Besuche in dieser Kalenderwoche</p>
      </div>
    `;
    return;
  }

  container.innerHTML = sortBesuche(besuche).map(b => `
    <div class="wb-besuch">
      <div class="wb-besuch-header">
        <span class="wb-besuch-firma">${escapeHtml(b.firma)}</span>
        ${b.segment ? `<span class="badge badge-segment">${escapeHtml(b.segment)}</span>` : ''}
        ${b.besuchstyp ? `<span class="badge badge-typ">${escapeHtml(b.besuchstyp)}</span>` : ''}
        ${b.datum ? `<span style="font-size:13px; color:var(--text-light);">📅 ${formatDatum(b.datum)}</span>` : ''}
        ${b.ort ? `<span style="font-size:13px; color:var(--text-light);">📍 ${escapeHtml(b.ort)}</span>` : ''}
      </div>

      <div class="wb-columns">
        <div>
          <div class="wb-column-title">📋 THEMEN</div>
          <div>${textToBullets(b.themen) || '<span style="color:var(--text-light)">Keine Themen erfasst</span>'}</div>
        </div>
        <div>
          <div class="wb-column-title">⚙️ ERGEBNIS & SCHRITTE</div>
          ${b.ergebnis ? `<div class="wb-ergebnis"><strong>Ergebnis:</strong> ${textToBullets(b.ergebnis)}</div>` : ''}
          ${b.naechsteSchritte ? `<div class="wb-next-steps"><strong>Daniel:</strong> ${textToBullets(b.naechsteSchritte)}</div>` : ''}
          ${b.todosMarco ? `<div class="wb-marco-steps"><strong>Marco:</strong> ${textToBullets(b.todosMarco)}</div>` : ''}
        </div>
      </div>
    </div>
  `).join('');
}

// ---- Export ----

// Segment-Icons (Kuerzel fuer PDF-Kreise)
const SEGMENT_ICONS = {
  'Schreiner': 'T', 'Maler': 'M', 'Zimmerei': 'H',
  'Fensterbau': 'F', 'Holzbau': 'Hb', 'Architekt': 'A',
  'Holzbauingenieure': 'HI', 'Hobelwerke': 'Hw', 'Saegereien': 'Sä',
  'Handel': 'D',
  'Diotrol-Intern': 'DI', 'Office': 'O', 'Sonstiges': 'X'
};

// Segment-Farben (gleich wie im Ranking)
const SEG_FARBEN_RGB = {
  'Architekt': [91, 127, 166], 'Maler': [184, 122, 107], 'Zimmerei': [106, 155, 106],
  'Fensterbau': [139, 123, 170], 'Schreiner': [168, 144, 96], 'Holzbau': [106, 155, 155],
  'Holzbauingenieure': [90, 138, 122], 'Hobelwerke': [138, 122, 90], 'Saegereien': [122, 106, 90],
  'Handel': [196, 149, 106],
  'Diotrol-Intern': [74, 122, 106], 'Office': [106, 138, 155], 'Sonstiges': [138, 138, 138]
};

function toggleExportDropdown() {
  const dd = document.getElementById('exportDropdown');
  dd.classList.toggle('show');
  const closeHandler = (e) => {
    if (!e.target.closest('.dropdown')) {
      dd.classList.remove('show');
      document.removeEventListener('click', closeHandler);
    }
  };
  setTimeout(() => document.addEventListener('click', closeHandler), 10);
}

// Daten fuer Executive Summary aufbereiten
function berechneSummary(data) {
  const besuche = data.besuche || [];
  const firmen = new Set(besuche.map(b => b.firma).filter(Boolean));
  const produkteAll = [];
  const followUpsDaniel = [];
  const followUpsMarco = [];
  const marktinfos = [];
  const segmente = {};

  besuche.forEach(b => {
    if (b.produkte) {
      b.produkte.split(',').forEach(p => {
        const t = p.trim();
        if (t) produkteAll.push(t);
      });
    }
    if (b.naechsteSchritte) {
      // Auch in der Summary: Zeilen intelligent auf Daniel/Marco aufteilen
      const lines = b.naechsteSchritte.split('\n')
        .map(l => l.trim().replace(/^[\u2022\-\*]\s*/, ''))
        .filter(l => l.length > 0);
      const danielLines = [];
      const marcoLines = [];
      lines.forEach(line => {
        if (istMarcoTodo(line)) {
          marcoLines.push(line);
        } else {
          danielLines.push(line);
        }
      });
      if (danielLines.length > 0) {
        followUpsDaniel.push({ firma: b.firma || 'Intern', text: danielLines.join('\n'), datum: b.datum });
      }
      if (marcoLines.length > 0) {
        followUpsMarco.push({ firma: b.firma || 'Intern', text: marcoLines.join('\n'), datum: b.datum });
      }
    }
    if (b.todosMarco) {
      followUpsMarco.push({ firma: b.firma || 'Intern', text: b.todosMarco, datum: b.datum });
    }
    if (b.ergebnis) {
      marktinfos.push({ firma: b.firma || '', text: b.ergebnis });
    }
    const seg = b.segment || 'Sonstiges';
    segmente[seg] = (segmente[seg] || 0) + 1;
  });

  // Produkt-Haeufigkeit
  const produktCount = {};
  produkteAll.forEach(p => { produktCount[p] = (produktCount[p] || 0) + 1; });
  const topProdukte = Object.entries(produktCount)
    .sort((a, b) => b[1] - a[1]).slice(0, 5);

  // Segment-Ranking sortiert
  const segmentRanking = Object.entries(segmente)
    .sort((a, b) => b[1] - a[1]);

  return { firmen, produkteAll, followUpsDaniel, followUpsMarco, marktinfos, segmente, segmentRanking, topProdukte };
}

// ============================================================
// PDF EXPORT (jsPDF) - mit To-Do-Split + Segment-Ranking
// ============================================================
function exportPDF() {
  if (!wochenberichtData) { showToast('Keine Daten geladen', 'error'); return; }
  if (typeof jspdf === 'undefined') { showToast('PDF-Bibliothek nicht geladen', 'error'); return; }

  const data = wochenberichtData;
  const stats = data.stats || {};
  const besuche = data.besuche || [];
  const summary = berechneSummary(data);
  const { jsPDF } = jspdf;
  const pdf = new jsPDF('p', 'mm', 'a4');

  // Farben
  const GRUEN = [26, 92, 58];
  const GRUEN_DUNKEL = [14, 61, 37];
  const GRUEN_HELL = [232, 245, 237];
  const AKZENT = [245, 166, 35];
  const AKZENT_HELL = [254, 249, 220];
  const BLAU = [91, 127, 166];
  const BLAU_HELL = [240, 245, 250];
  const TEXT_DUNKEL = [26, 26, 26];
  const TEXT_GRAU = [107, 114, 128];
  const WEISS = [255, 255, 255];
  const BORDER = [229, 231, 235];

  let y = 0;

  function checkPage(needed) {
    if (y + needed > 275) {
      pdf.addPage();
      pdf.setFillColor(...GRUEN);
      pdf.rect(0, 0, 210, 3, 'F');
      pdf.setFontSize(8);
      pdf.setTextColor(...GRUEN);
      pdf.text(`Daniel Pfister | Wochenbericht KW ${data.kw}`, 15, 8);
      y = 14;
    }
  }

  // === SEITE 1: Deckblatt ===
  // Gruener Header
  pdf.setFillColor(...GRUEN);
  pdf.rect(0, 0, 210, 48, 'F');
  pdf.setFillColor(...AKZENT);
  pdf.rect(0, 48, 210, 2, 'F');

  // Name + Untertitel
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(26);
  pdf.setTextColor(...WEISS);
  pdf.text('Daniel Pfister', 15, 22);
  pdf.setFont('helvetica', 'normal');
  pdf.setFontSize(10);
  pdf.setTextColor(200, 230, 210);
  pdf.text('Kundenbesuchsbericht | Diotrol AG | Aussendienst Schweiz', 15, 32);

  // KW-Kreis rechts
  pdf.setFillColor(...WEISS);
  pdf.circle(175, 26, 15, 'F');
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(9);
  pdf.setTextColor(...GRUEN);
  pdf.text('KW', 175, 21, { align: 'center' });
  pdf.setFontSize(22);
  pdf.text(String(data.kw), 175, 32, { align: 'center' });

  // Berichtsinfo
  y = 58;
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(22);
  pdf.setTextColor(...TEXT_DUNKEL);
  pdf.text('Wochenbericht', 15, y);
  y += 9;
  pdf.setFont('helvetica', 'normal');
  pdf.setFontSize(10);
  pdf.setTextColor(...TEXT_GRAU);
  const today = new Date();
  pdf.text(`Daniel Pfister | Erstellt: ${today.getDate().toString().padStart(2,'0')}.${(today.getMonth()+1).toString().padStart(2,'0')}.${today.getFullYear()}`, 15, y);
  y += 12;

  // === ZUSAMMENFASSUNG ===
  pdf.setFillColor(...GRUEN);
  pdf.rect(15, y, 180, 9, 'F');
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(13);
  pdf.setTextColor(...WEISS);
  pdf.text('  ZUSAMMENFASSUNG', 15, y + 6.5);
  y += 14;

  // Metriken-Kacheln
  const metriken = [
    [String(stats.totalBesuche || 0), 'Aktivitaeten'],
    [String(stats.neukunden || 0), 'Neukunden'],
    [String(stats.bestandskunden || 0), 'Bestandskunden'],
    [String(summary.firmen.size), 'Kunden besucht'],
  ];
  const kW = 42, kH = 32, kGap = 3;

  metriken.forEach(([zahl, label], i) => {
    const x = 15 + i * (kW + kGap);
    pdf.setFillColor(...WEISS);
    pdf.setDrawColor(...BORDER);
    pdf.setLineWidth(0.4);
    pdf.rect(x, y, kW, kH, 'DF');
    pdf.setFillColor(...GRUEN);
    pdf.rect(x, y, kW, 2.5, 'F');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(24);
    pdf.setTextColor(...GRUEN);
    pdf.text(zahl, x + kW / 2, y + 17, { align: 'center' });
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(7);
    pdf.setTextColor(...TEXT_GRAU);
    pdf.text(label, x + kW / 2, y + 25, { align: 'center' });
  });
  y += kH + 10;

  // === SEGMENT-RANKING (NEU - als sortierte Liste mit Balken) ===
  if (summary.segmentRanking.length > 0) {
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(10);
    pdf.setTextColor(...GRUEN_DUNKEL);
    pdf.text('Segmentverteilung (Ranking)', 16, y);
    y += 7;

    const maxSeg = summary.segmentRanking[0][1];
    summary.segmentRanking.forEach(([seg, cnt], idx) => {
      checkPage(8);
      const prozent = Math.round(cnt / besuche.length * 100);
      const balkenBreite = Math.round(cnt / maxSeg * 80); // max 80mm Balken
      const farbe = SEG_FARBEN_RGB[seg] || [107, 114, 128];

      // Rang
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(8);
      pdf.setTextColor(...TEXT_GRAU);
      pdf.text(`${idx + 1}.`, 18, y);

      // Segment-Icon Kreis
      pdf.setFillColor(...farbe);
      pdf.circle(24, y - 1, 2.5, 'F');
      const ik = SEGMENT_ICONS[seg] || 'X';
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(5);
      pdf.setTextColor(...WEISS);
      pdf.text(ik, 24, y - 0.2, { align: 'center' });

      // Segmentname
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(8);
      pdf.setTextColor(...TEXT_DUNKEL);
      pdf.text(seg, 29, y);

      // Balken
      pdf.setFillColor(235, 237, 240);
      pdf.rect(60, y - 2.5, 80, 3.5, 'F');
      pdf.setFillColor(...farbe);
      pdf.rect(60, y - 2.5, balkenBreite, 3.5, 'F');

      // Zahl + Prozent
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(8);
      pdf.setTextColor(...farbe);
      pdf.text(`${cnt}`, 144, y);
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(7);
      pdf.setTextColor(...TEXT_GRAU);
      pdf.text(`(${prozent}%)`, 151, y);

      y += 6;
    });
    y += 4;
  }

  // Rechts daneben: Top-Produkte (nur wenn Platz)
  // (Produkte kommen jetzt unter dem Ranking)
  if (summary.topProdukte.length > 0) {
    checkPage(20);
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(10);
    pdf.setTextColor(...GRUEN_DUNKEL);
    pdf.text('Top Produkte', 16, y);
    y += 7;
    summary.topProdukte.forEach(([prod, cnt]) => {
      pdf.setFillColor(...GRUEN);
      pdf.circle(19, y - 1.2, 1.2, 'F');
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(9);
      pdf.setTextColor(...TEXT_DUNKEL);
      pdf.text(`${prod} (${cnt}x)`, 23, y);
      y += 5;
    });
    y += 6;
  }

  // === TO-DOs DANIEL ===
  if (summary.followUpsDaniel.length > 0) {
    checkPage(30);

    pdf.setFillColor(...GRUEN_HELL);
    pdf.rect(15, y, 180, 10, 'F');
    pdf.setFillColor(...GRUEN);
    pdf.rect(15, y, 3, 10, 'F');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(11);
    pdf.setTextColor(...GRUEN_DUNKEL);
    pdf.text('MEINE TO-DOs (Daniel)', 22, y + 7);
    y += 14;

    // Tabellenkopf
    pdf.setFillColor(...GRUEN);
    pdf.rect(15, y - 3, 180, 7, 'F');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(8);
    pdf.setTextColor(...WEISS);
    pdf.text('Kunde', 18, y + 1);
    pdf.text('Aktion', 62, y + 1);
    pdf.text('Datum', 172, y + 1);
    y += 7;

    summary.followUpsDaniel.slice(0, 12).forEach((fu, idx) => {
      checkPage(9);
      if (idx % 2 === 0) {
        pdf.setFillColor(248, 250, 252);
        pdf.rect(15, y - 4, 180, 7, 'F');
      }
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(8);
      pdf.setTextColor(...GRUEN_DUNKEL);
      pdf.text(fu.firma.substring(0, 28), 18, y);
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(8);
      pdf.setTextColor(...TEXT_DUNKEL);
      const txt = fu.text.length > 65 ? fu.text.substring(0, 65) + '...' : fu.text;
      pdf.text(txt.replace(/\n/g, ' '), 62, y);
      pdf.setFontSize(7.5);
      pdf.setTextColor(...TEXT_GRAU);
      pdf.text(formatDatumKurz(fu.datum), 172, y);
      y += 7;
    });
    y += 6;
  }

  // === TO-DOs MARCO ===
  if (summary.followUpsMarco.length > 0) {
    checkPage(30);

    pdf.setFillColor(...BLAU_HELL);
    pdf.rect(15, y, 180, 10, 'F');
    pdf.setFillColor(...BLAU);
    pdf.rect(15, y, 3, 10, 'F');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(11);
    pdf.setTextColor(70, 95, 130);
    pdf.text('AUFGABEN FUER MARCO', 22, y + 7);
    y += 14;

    // Tabellenkopf
    pdf.setFillColor(...BLAU);
    pdf.rect(15, y - 3, 180, 7, 'F');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(8);
    pdf.setTextColor(...WEISS);
    pdf.text('Kunde', 18, y + 1);
    pdf.text('Aktion', 62, y + 1);
    pdf.text('Datum', 172, y + 1);
    y += 7;

    summary.followUpsMarco.slice(0, 12).forEach((fu, idx) => {
      checkPage(9);
      if (idx % 2 === 0) {
        pdf.setFillColor(245, 248, 252);
        pdf.rect(15, y - 4, 180, 7, 'F');
      }
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(8);
      pdf.setTextColor(70, 95, 130);
      pdf.text(fu.firma.substring(0, 28), 18, y);
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(8);
      pdf.setTextColor(...TEXT_DUNKEL);
      const txt = fu.text.length > 65 ? fu.text.substring(0, 65) + '...' : fu.text;
      pdf.text(txt.replace(/\n/g, ' '), 62, y);
      pdf.setFontSize(7.5);
      pdf.setTextColor(...TEXT_GRAU);
      pdf.text(formatDatumKurz(fu.datum), 172, y);
      y += 7;
    });
    y += 6;
  }

  // === DETAILBERICHT ===
  pdf.addPage();
  pdf.setFillColor(...GRUEN);
  pdf.rect(0, 0, 210, 3, 'F');
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(8);
  pdf.setTextColor(...GRUEN);
  pdf.text(`Daniel Pfister | Wochenbericht KW ${data.kw}`, 15, 8);
  y = 16;

  // Nach Tag gruppieren
  const tage = {};
  besuche.forEach(b => {
    const tag = b.datum ? getWochentag(b.datum) : (b.tag || 'Unbekannt');
    if (!tage[tag]) tage[tag] = [];
    tage[tag].push(b);
  });

  const tagReihenfolge = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag'];
  tagReihenfolge.forEach(tag => {
    if (!tage[tag]) return;

    checkPage(20);

    // Tag-Header
    pdf.setFillColor(...GRUEN);
    pdf.rect(15, y, 180, 9, 'F');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(10);
    pdf.setTextColor(...WEISS);
    pdf.text(`  ${tag.toUpperCase()}`, 18, y + 6);
    const anz = tage[tag].length;
    pdf.setFontSize(8);
    pdf.text(`${anz} Besuch${anz !== 1 ? 'e' : ''}`, 190, y + 6, { align: 'right' });
    y += 13;

    // Besuche (sortiert: Office zuerst, Intern zuletzt)
    sortBesuche(tage[tag]).forEach((b, idx) => {
      checkPage(25);

      // Icon-Kreis + Firma
      const seg = b.segment || 'Sonstiges';
      const ik = SEGMENT_ICONS[seg] || 'X';
      const segFarbe = SEG_FARBEN_RGB[seg] || [107, 114, 128];
      pdf.setFillColor(...segFarbe);
      pdf.circle(20, y, 2.5, 'F');
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(4.5);
      pdf.setTextColor(...WEISS);
      pdf.text(ik, 20, y + 0.5, { align: 'center' });

      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(10);
      pdf.setTextColor(...TEXT_DUNKEL);
      const firmaText = b.firma || 'Unbekannt';
      pdf.text(firmaText, 25, y + 1);

      if (b.kontaktperson) {
        pdf.setFont('helvetica', 'normal');
        pdf.setFontSize(8);
        pdf.setTextColor(...TEXT_GRAU);
        pdf.text(b.kontaktperson, 25 + pdf.getTextWidth(firmaText) + 3, y + 1);
      }
      if (b.segment) {
        pdf.setFontSize(7);
        pdf.setTextColor(...segFarbe);
        pdf.text(b.segment, 190, y + 1, { align: 'right' });
      }
      y += 7;

      // Detail-Felder
      const felder = [
        ['Themen', b.themen, null],
        ['Produkte', b.produkte, null],
        ['Ergebnis', b.ergebnis, [GRUEN_HELL, GRUEN]],
        ['Daniel To-Do', b.naechsteSchritte, [AKZENT_HELL, AKZENT]],
        ['Marco To-Do', b.todosMarco, [BLAU_HELL, BLAU]],
        ['Stellungnahme', b.stellungnahme, null],
      ];

      felder.forEach(([label, wert, highlight]) => {
        if (!wert) return;
        checkPage(12);

        const yVor = y;

        pdf.setFont('helvetica', 'bold');
        pdf.setFontSize(8);
        const labelColor = highlight ? highlight[1] : GRUEN_DUNKEL;
        pdf.setTextColor(...labelColor);
        pdf.text(`${label}:`, 25, y);
        y += 4.5;

        pdf.setFont('helvetica', 'normal');
        pdf.setFontSize(8);
        pdf.setTextColor(...TEXT_DUNKEL);
        const lines = pdf.splitTextToSize(wert, 163);
        lines.forEach(line => {
          checkPage(5);
          pdf.text(line, 25, y);
          y += 4;
        });
        y += 1;

        if (highlight) {
          const blockH = y - yVor + 1;
          pdf.setFillColor(...highlight[0]);
          pdf.rect(21, yVor - 3, 172, blockH, 'F');
          pdf.setFillColor(...highlight[1]);
          pdf.rect(21, yVor - 3, 2, blockH, 'F');
          pdf.setFont('helvetica', 'bold');
          pdf.setFontSize(8);
          pdf.setTextColor(...highlight[1]);
          pdf.text(`${label}:`, 25, yVor);
          pdf.setFont('helvetica', 'normal');
          pdf.setTextColor(...TEXT_DUNKEL);
          let reY = yVor + 4.5;
          lines.forEach(line => {
            pdf.text(line, 25, reY);
            reY += 4;
          });
        }

        y += 1;
      });

      // Trennlinie
      if (idx < tage[tag].length - 1) {
        pdf.setDrawColor(...BORDER);
        pdf.setLineWidth(0.2);
        pdf.line(22, y, 190, y);
        y += 5;
      } else {
        y += 3;
      }
    });
    y += 5;
  });

  // Footer auf jeder Seite
  const pageCount = pdf.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    pdf.setPage(i);
    pdf.setDrawColor(...BORDER);
    pdf.setLineWidth(0.3);
    pdf.line(15, 282, 195, 282);
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(7);
    pdf.setTextColor(...TEXT_GRAU);
    pdf.text('Daniel Pfister | Diotrol AG | Vertraulich', 15, 287);
    pdf.text(`Seite ${i} von ${pageCount}`, 190, 287, { align: 'right' });
  }

  pdf.save(`Wochenbericht_KW${data.kw}_${data.jahr}.pdf`);
  showToast('PDF-Export heruntergeladen', 'success');
}

// ============================================================
// WORD EXPORT - mit To-Do-Split + Segment-Ranking
// ============================================================
function exportWord() {
  const data = wochenberichtData;
  if (!data) return;

  const stats = data.stats || {};
  const summary = berechneSummary(data);

  // Segment-Emoji fuer Word
  const SEG_EMOJI = {
    'Schreiner': '\u{1F6CB}', 'Maler': '\u{1F3A8}', 'Zimmerei': '\u{1F3D7}',
    'Fensterbau': '\u{1F5BC}', 'Fassadenbau': '\u{1F3D7}', 'Architekt': '\u{1F4D0}',
    'Generalunternehmer': '\u{1F3E2}', 'Handel': '\u{1F3EA}',
    'Diotrol-Intern': '\u{1F3E2}', 'Office': '\u{1F4BC}', 'Sonstiges': '\u{1F4CB}'
  };

  let html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
<head><meta charset="utf-8"><title>Wochenbericht</title>
<!--[if gte mso 9]><xml>
<w:WordDocument><w:View>Print</w:View>
<w:DoNotOptimizeForBrowser/>
</w:WordDocument></xml><![endif]-->
<style>
@page { mso-page-orientation: portrait; }
@page Section1 {
  mso-header-margin: 0.5in;
  mso-footer-margin: 0.5in;
  mso-footer: f1;
}
div.Section1 { page: Section1; }
p.MsoFooter { font-family: Calibri; font-size: 8pt; color: #6b7280; text-align: center; }
body { font-family: Calibri, sans-serif; font-size: 11pt; color: #1a1a1a; }
h1 { font-size: 18pt; color: #1a5c3a; margin-bottom: 4pt; }
h2 { font-size: 13pt; color: #1a5c3a; border-bottom: 2px solid #1a5c3a; padding-bottom: 4pt; margin-top: 16pt; }
h3 { font-size: 11pt; color: #0e3d25; margin-top: 12pt; margin-bottom: 4pt; }
table { border-collapse: collapse; width: 100%; margin: 8pt 0; }
td, th { padding: 5pt 10pt; vertical-align: top; text-align: left; }
th { background-color: #1a5c3a; color: white; font-size: 9pt; }
th.marco { background-color: #5b7fa6; }
.header-bg { background-color: #1a5c3a; color: white; padding: 12pt 16pt; }
.header-bg h1 { color: white; margin: 0; }
.header-bg p { color: #c8e6d2; margin: 4pt 0 0 0; }
.accent-bar { background-color: #f5a623; height: 4pt; }
.stat-table td { text-align: center; background: #e8f5ed; padding: 8pt; }
.stat-zahl { font-size: 20pt; font-weight: bold; color: #1a5c3a; display: block; }
.stat-label { font-size: 8pt; color: #6b7280; }
.highlight { background: #e8f5ed; padding: 6pt 10pt; border-left: 4px solid #1a5c3a; margin: 6pt 0; }
.todo-highlight { background: #fef9c3; padding: 6pt 10pt; border-left: 4px solid #f5a623; margin: 6pt 0; }
.marco-highlight { background: #f0f5fa; padding: 6pt 10pt; border-left: 4px solid #5b7fa6; margin: 6pt 0; }
.todo-title-daniel { background: #e8f5ed; padding: 8pt; font-size: 13pt; font-weight: bold; color: #0e3d25; border-left: 4px solid #1a5c3a; }
.todo-title-marco { background: #f0f5fa; padding: 8pt; font-size: 13pt; font-weight: bold; color: #465f82; border-left: 4px solid #5b7fa6; }
.label-text { font-weight: bold; color: #0e3d25; font-size: 9pt; }
.muted { color: #6b7280; font-size: 9pt; }
.segment-badge { color: #6b7280; font-size: 9pt; }
.ranking-bar { display: inline-block; height: 12pt; border-radius: 2pt; margin-right: 6pt; vertical-align: middle; }
</style></head><body>`;

  // Footer
  html += `<div style="mso-element:footer" id="f1">
    <p class="MsoFooter"><span style="color:#6b7280;">Daniel Pfister | Diotrol AG | Vertraulich</span>
    <span style="float:right; color:#6b7280;">Seite <span style="mso-field-code:'PAGE'"></span><!--[if gte mso 9]><span style="mso-element:field-begin"></span> PAGE <span style="mso-element:field-end"></span><![endif]--></span></p>
  </div>`;

  html += '<div class="Section1">';

  // Header
  html += `<div class="header-bg">
    <h1>Daniel Pfister</h1>
    <p>Kundenbesuchsbericht | Diotrol AG | Aussendienst Schweiz</p>
  </div>
  <div class="accent-bar"></div>`;

  html += `<h1 style="margin-top:12pt;">Wochenbericht KW ${data.kw}</h1>`;
  html += `<p class="muted">Daniel Pfister | Erstellt: ${new Date().toLocaleDateString('de-CH')}</p>`;

  // Executive Summary
  html += `<h2>\u{1F4CA} ZUSAMMENFASSUNG</h2>`;
  html += `<table class="stat-table"><tr>`;
  [
    [stats.totalBesuche || 0, 'Aktivitaeten'],
    [stats.neukunden || 0, 'Neukunden'],
    [stats.bestandskunden || 0, 'Bestandskunden'],
    [summary.firmen.size, 'Kunden besucht']
  ].forEach(([z, l]) => {
    html += `<td><span class="stat-zahl">${z}</span><span class="stat-label">${l}</span></td>`;
  });
  html += `</tr></table>`;

  // Segment-Ranking (NEU - als sortiertes Ranking mit Balken)
  if (summary.segmentRanking.length > 0) {
    const SEG_FARBEN_HEX = {
      'Architekt': '#5b7fa6', 'Maler': '#b87a6b', 'Zimmerei': '#6a9b6a',
      'Fensterbau': '#8b7baa', 'Schreiner': '#a89060', 'Fassadenbau': '#6a9b9b',
      'Handel': '#c4956a', 'Generalunternehmer': '#7a7a9b',
      'Diotrol-Intern': '#4a7a6a', 'Office': '#6a8a9b', 'Sonstiges': '#8a8a8a'
    };
    const maxSeg = summary.segmentRanking[0][1];

    html += '<p class="label-text" style="font-size:11pt;">Segmentverteilung (Ranking):</p>';
    html += '<table style="width:100%;">';
    summary.segmentRanking.forEach(([seg, cnt], idx) => {
      const prozent = Math.round(cnt / (data.besuche || []).length * 100);
      const balkenBreite = Math.round(cnt / maxSeg * 100);
      const farbe = SEG_FARBEN_HEX[seg] || '#6b7280';
      const emoji = SEG_EMOJI[seg] || '\u{1F4CB}';
      html += `<tr>
        <td style="width:5%; font-weight:bold; color:#6b7280; font-size:9pt; padding:3pt;">${idx + 1}.</td>
        <td style="width:5%; font-size:9pt; padding:3pt;">${emoji}</td>
        <td style="width:25%; font-size:9pt; padding:3pt;">${seg}</td>
        <td style="width:45%; padding:3pt;">
          <span class="ranking-bar" style="width:${balkenBreite}%; background:${farbe};">&nbsp;</span>
        </td>
        <td style="width:10%; font-weight:bold; color:${farbe}; font-size:9pt; padding:3pt; text-align:right;">${cnt}</td>
        <td style="width:10%; color:#6b7280; font-size:8pt; padding:3pt;">(${prozent}%)</td>
      </tr>`;
    });
    html += '</table>';
  }

  // Top-Produkte
  if (summary.topProdukte.length > 0) {
    html += '<p class="label-text">Top Produkte:</p>';
    summary.topProdukte.forEach(([prod, cnt]) => {
      html += `<p style="margin:2pt 0 2pt 8pt; font-size:9pt; color:#1a5c3a;">\u25CF ${prod} (${cnt}x)</p>`;
    });
  }

  // To-Dos Daniel
  if (summary.followUpsDaniel.length > 0) {
    html += `<div class="todo-title-daniel">\u2B50 MEINE TO-DOs (Daniel)</div>`;
    html += `<table><tr><th>Kunde</th><th>Aktion</th><th>Datum</th></tr>`;
    summary.followUpsDaniel.forEach(fu => {
      const txt = fu.text.length > 100 ? fu.text.substring(0, 100) + '...' : fu.text;
      html += `<tr>
        <td style="font-weight:bold; color:#0e3d25; font-size:9pt; width:20%;">${escapeHtml(fu.firma)}</td>
        <td style="font-size:9pt;">${escapeHtml(txt.replace(/\n/g, ' '))}</td>
        <td style="font-size:9pt; color:#6b7280; width:12%;">${formatDatumKurz(fu.datum)}</td>
      </tr>`;
    });
    html += '</table>';
  }

  // To-Dos Marco
  if (summary.followUpsMarco.length > 0) {
    html += `<div class="todo-title-marco">\u{1F454} AUFGABEN FUER MARCO</div>`;
    html += `<table><tr><th class="marco">Kunde</th><th class="marco">Aktion</th><th class="marco">Datum</th></tr>`;
    summary.followUpsMarco.forEach(fu => {
      const txt = fu.text.length > 100 ? fu.text.substring(0, 100) + '...' : fu.text;
      html += `<tr>
        <td style="font-weight:bold; color:#465f82; font-size:9pt; width:20%;">${escapeHtml(fu.firma)}</td>
        <td style="font-size:9pt;">${escapeHtml(txt.replace(/\n/g, ' '))}</td>
        <td style="font-size:9pt; color:#6b7280; width:12%;">${formatDatumKurz(fu.datum)}</td>
      </tr>`;
    });
    html += '</table>';
  }

  // Seitenumbruch + Detailbericht
  html += '<br clear="all" style="page-break-before:always">';
  html += `<h2>\u{1F4C5} DETAILBERICHT KW ${data.kw}</h2>`;

  const tage = {};
  (data.besuche || []).forEach(b => {
    const tag = b.datum ? getWochentag(b.datum) : (b.tag || 'Unbekannt');
    if (!tage[tag]) tage[tag] = [];
    tage[tag].push(b);
  });

  ['Montag','Dienstag','Mittwoch','Donnerstag','Freitag','Samstag'].forEach(tag => {
    if (!tage[tag]) return;
    html += `<h2 style="background:#1a5c3a; color:white; padding:6pt 10pt; border:none;">
      ${tag.toUpperCase()} <span style="float:right; font-size:9pt; font-weight:normal;">
      ${tage[tag].length} Besuch${tage[tag].length !== 1 ? 'e' : ''}</span></h2>`;

    sortBesuche(tage[tag]).forEach(b => {
      const segEmoji = SEG_EMOJI[b.segment] || '\u{1F4CB}';
      html += `<h3>${segEmoji} ${escapeHtml(b.firma || 'Unbekannt')}`;
      if (b.kontaktperson) html += ` <span class="muted">| ${escapeHtml(b.kontaktperson)}</span>`;
      html += `</h3>`;
      if (b.segment) html += `<p class="segment-badge">${escapeHtml(b.segment)}</p>`;

      if (b.themen) html += `<p><span class="label-text">\u{1F4AC} Themen:</span><br>${escapeHtml(b.themen).replace(/\n/g,'<br>')}</p>`;
      if (b.produkte) html += `<p><span class="label-text">\u{1F4E6} Produkte:</span> ${escapeHtml(b.produkte)}</p>`;
      if (b.ergebnis) html += `<div class="highlight"><span class="label-text">\u{1F3AF} Ergebnis:</span><br>${escapeHtml(b.ergebnis).replace(/\n/g,'<br>')}</div>`;
      if (b.naechsteSchritte) html += `<div class="todo-highlight"><span class="label-text">\u2B50 Daniel To-Do:</span><br>${escapeHtml(b.naechsteSchritte).replace(/\n/g,'<br>')}</div>`;
      if (b.todosMarco) html += `<div class="marco-highlight"><span class="label-text">\u{1F454} Marco To-Do:</span><br>${escapeHtml(b.todosMarco).replace(/\n/g,'<br>')}</div>`;
      if (b.stellungnahme) html += `<p class="muted"><span class="label-text">\u{1F4DD} Stellungnahme:</span><br>${escapeHtml(b.stellungnahme).replace(/\n/g,'<br>')}</p>`;
      html += '<hr style="border:none; border-top:1px solid #e5e7eb; margin:8pt 0;">';
    });
  });

  html += '</div>';
  html += '</body></html>';

  const blob = new Blob(['\ufeff' + html], { type: 'application/msword' });
  downloadBlob(blob, `Wochenbericht_KW${data.kw}_${data.jahr}.doc`);
  showToast('Word-Export heruntergeladen', 'success');
}

// ============================================================
// TEXT + EMAIL EXPORT
// ============================================================
function generateReportText() {
  if (!wochenberichtData) return '';
  const data = wochenberichtData;
  const stats = data.stats || {};
  const summary = berechneSummary(data);

  let text = `WOCHENBERICHT KW ${data.kw} / ${data.jahr}\n`;
  text += `${'='.repeat(50)}\n\n`;
  text += `ZUSAMMENFASSUNG\n`;
  text += `Total Besuche: ${stats.totalBesuche || 0}\n`;
  text += `Neukunden: ${stats.neukunden || 0} | Bestandskunden: ${stats.bestandskunden || 0}\n`;
  text += `Kunden besucht: ${summary.firmen.size}\n`;
  if (stats.fokusProdukte) text += `Fokus-Produkte: ${stats.fokusProdukte}\n`;
  text += '\n';

  // Segment-Ranking
  if (summary.segmentRanking.length > 0) {
    text += `SEGMENTVERTEILUNG\n`;
    text += `${'-'.repeat(40)}\n`;
    summary.segmentRanking.forEach(([seg, cnt], idx) => {
      const prozent = Math.round(cnt / (data.besuche || []).length * 100);
      text += `${idx + 1}. ${seg}: ${cnt} Besuche (${prozent}%)\n`;
    });
    text += '\n';
  }

  // To-Dos Daniel
  if (summary.followUpsDaniel.length > 0) {
    text += `MEINE TO-DOs (Daniel)\n`;
    text += `${'-'.repeat(40)}\n`;
    summary.followUpsDaniel.forEach(fu => {
      text += `[ ] ${fu.firma}: ${fu.text.replace(/\n/g, ' ')}\n`;
    });
    text += '\n';
  }

  // To-Dos Marco
  if (summary.followUpsMarco.length > 0) {
    text += `AUFGABEN FUER MARCO\n`;
    text += `${'-'.repeat(40)}\n`;
    summary.followUpsMarco.forEach(fu => {
      text += `[ ] ${fu.firma}: ${fu.text.replace(/\n/g, ' ')}\n`;
    });
    text += '\n';
  }

  text += `${'='.repeat(50)}\nDETAILS\n${'='.repeat(50)}\n\n`;

  sortBesuche(data.besuche || []).forEach(b => {
    text += `--- ${b.firma} ---\n`;
    text += `Datum: ${formatDatum(b.datum)}\n`;
    if (b.ort) text += `Ort: ${b.ort}\n`;
    if (b.kontaktperson) text += `Kontakt: ${b.kontaktperson}\n`;
    if (b.segment) text += `Segment: ${b.segment}\n`;
    if (b.besuchstyp) text += `Besuchstyp: ${b.besuchstyp}\n`;
    if (b.produkte) text += `Produkte: ${b.produkte}\n`;
    if (b.themen) text += `\nThemen:\n${b.themen}\n`;
    if (b.ergebnis) text += `\nErgebnis:\n${b.ergebnis}\n`;
    if (b.naechsteSchritte) text += `\nDaniel To-Do:\n${b.naechsteSchritte}\n`;
    if (b.todosMarco) text += `\nMarco To-Do:\n${b.todosMarco}\n`;
    if (b.stellungnahme) text += `\nStellungnahme:\n${b.stellungnahme}\n`;
    text += '\n';
  });

  return text;
}

function exportText() {
  const text = generateReportText();
  const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
  downloadBlob(blob, `Wochenbericht_KW${wochenberichtData.kw}_${wochenberichtData.jahr}.txt`);
  showToast('Text-Export heruntergeladen', 'success');
}

function exportEmail() {
  const text = generateReportText();
  const subject = encodeURIComponent(`Wochenbericht KW ${wochenberichtData?.kw || ''} / ${wochenberichtData?.jahr || ''}`);
  const body = encodeURIComponent(text);
  window.open(`mailto:?subject=${subject}&body=${body}`, '_self');
}

// === MARCO-BERICHT (kompakt, nur fuer den Chef) ===
function exportMarcoReport() {
  if (!wochenberichtData) { showToast('Keine Daten geladen', 'error'); return; }
  const data = wochenberichtData;
  const stats = data.stats || {};
  const summary = berechneSummary(data);

  let text = `WOCHENBERICHT KW ${data.kw} - Zusammenfassung fuer Marco\n`;
  text += `${'='.repeat(50)}\n\n`;
  text += `Total Besuche: ${stats.totalBesuche || 0} | Neukunden: ${stats.neukunden || 0} | Kunden besucht: ${summary.firmen.size}\n\n`;

  // Segment-Ranking
  if (summary.segmentRanking.length > 0) {
    text += `SEGMENTVERTEILUNG:\n`;
    summary.segmentRanking.forEach(([seg, cnt], idx) => {
      const prozent = Math.round(cnt / (data.besuche || []).length * 100);
      text += `  ${idx + 1}. ${seg}: ${cnt} (${prozent}%)\n`;
    });
    text += '\n';
  }

  // Marcos Aufgaben
  if (summary.followUpsMarco.length > 0) {
    text += `DEINE AUFGABEN:\n`;
    text += `${'-'.repeat(40)}\n`;
    summary.followUpsMarco.forEach(fu => {
      text += `[ ] [${fu.firma}]: ${fu.text.replace(/\n/g, ' ')}\n`;
    });
    text += '\n';
  }

  // Top Highlights
  text += `TOP-HIGHLIGHTS:\n`;
  text += `${'-'.repeat(40)}\n`;
  (data.besuche || []).slice(0, 5).forEach(b => {
    if (b.ergebnis) {
      const kurz = b.ergebnis.length > 80 ? b.ergebnis.substring(0, 80) + '...' : b.ergebnis;
      text += `- ${b.firma}: ${kurz.replace(/\n/g, ' ')}\n`;
    }
  });

  text += `\n---\nDaniel Pfister | Diotrol AG | KW ${data.kw}`;

  const subject = encodeURIComponent(`Wochenbericht KW ${data.kw} - Zusammenfassung`);
  const body = encodeURIComponent(text);
  window.open(`mailto:?subject=${subject}&body=${body}`, '_self');
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
