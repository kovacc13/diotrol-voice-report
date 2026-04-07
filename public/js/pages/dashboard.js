// ========================================
// Dashboard-Seite
// Replit-Referenz-Design: 2x2 Stat-Cards
// ========================================

async function renderDashboard() {
  const main = document.getElementById('mainContent');
  const kw = getCurrentKW();

  // Daten parallel laden
  let besuche = [], kunden = [], besucheKW = [];
  try {
    const [besucheData, kundenData] = await Promise.all([
      API.besuche.laden(),
      API.kunden.laden()
    ]);
    besuche = besucheData || [];
    kunden = kundenData || [];
    besucheKW = besuche.filter(b => b.kw === kw);
  } catch (e) {
    console.error('Dashboard laden fehlgeschlagen:', e);
  }

  // Offene Projekte zaehlen
  const offeneProjekte = new Set(besuche.filter(b => b.projekt && b.projekt.trim()).map(b => b.projekt.trim())).size;

  // Letzte 5 Aktivitaeten
  const letzteAktivitaeten = besuche.slice(0, 5);

  // Segment-Ranking fuer aktuelle KW
  const segmenteKW = {};
  besucheKW.forEach(b => {
    const seg = b.segment || 'Sonstiges';
    segmenteKW[seg] = (segmenteKW[seg] || 0) + 1;
  });
  const segmentRankingKW = Object.entries(segmenteKW).sort((a, b) => b[1] - a[1]);
  const maxSegKW = segmentRankingKW[0] ? segmentRankingKW[0][1] : 1;

  const SEG_FARBEN_DASH = {
    'Architekt': '#5b7fa6', 'Maler': '#b87a6b', 'Zimmerei': '#6a9b6a',
    'Fensterbau': '#8b7baa', 'Schreiner': '#a89060', 'Holzbau': '#6a9b9b',
    'Holzbauingenieure': '#5a8a7a', 'Hobelwerke': '#8a7a5a', 'Saegereien': '#7a6a5a',
    'Handel': '#c4956a',
    'Diotrol-Intern': '#4a7a6a', 'Office': '#6a8a9b', 'Sonstiges': '#8a8a8a'
  };

  const SEGMENT_ICONS_DASH = {
    'Schreiner': 'T', 'Maler': 'M', 'Zimmerei': 'H',
    'Fensterbau': 'F', 'Holzbau': 'Hb', 'Architekt': 'A',
    'Holzbauingenieure': 'HI', 'Hobelwerke': 'Hw', 'Saegereien': 'Sä',
    'Handel': 'D',
    'Diotrol-Intern': 'DI', 'Office': 'O', 'Sonstiges': 'X'
  };

  main.innerHTML = `
    <div class="page-header">
      <div>
        <h1>Dashboard</h1>
        <p class="subtitle">Willkommen zurueck. Hier ist Ihre Uebersicht.</p>
      </div>
      <a href="#/neuer-besuch" class="btn btn-orange">+ Neuer Besuch</a>
    </div>

    <div class="stat-cards">
      <a href="#/besuche" class="stat-card">
        <div class="stat-icon">\u{1F4C5}</div>
        <div class="stat-info">
          <div class="stat-label">Besuche KW ${kw}</div>
          <div class="stat-value">${besucheKW.length}</div>
        </div>
      </a>
      <a href="#/besuche" class="stat-card">
        <div class="stat-icon">\u{1F4C8}</div>
        <div class="stat-info">
          <div class="stat-label">Besuche gesamt</div>
          <div class="stat-value">${besuche.length}</div>
        </div>
      </a>
      <a href="#/kunden" class="stat-card">
        <div class="stat-icon">\u{1F465}</div>
        <div class="stat-info">
          <div class="stat-label">Kunden gesamt</div>
          <div class="stat-value">${kunden.length}</div>
        </div>
      </a>
      <a href="#/besuche" class="stat-card">
        <div class="stat-icon">\u{1F4CB}</div>
        <div class="stat-info">
          <div class="stat-label">Offene Projekte</div>
          <div class="stat-value">${offeneProjekte}</div>
        </div>
      </a>
    </div>

    ${segmentRankingKW.length > 0 ? `
    <div class="card" style="margin-bottom:16px;">
      <h3 class="card-title">📊 Segmentverteilung KW ${kw}</h3>
      <div class="segment-ranking-list">
        ${segmentRankingKW.map(([seg, cnt], idx) => {
          const prozent = besucheKW.length > 0 ? Math.round(cnt / besucheKW.length * 100) : 0;
          const balkenBreite = Math.round(cnt / maxSegKW * 100);
          const farbe = SEG_FARBEN_DASH[seg] || '#6b7280';
          const icon = SEGMENT_ICONS_DASH[seg] || 'X';
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
    </div>` : ''}

    <div class="card">
      <h3 class="card-title">Letzte Aktivitaeten</h3>
      ${letzteAktivitaeten.length === 0
        ? '<div class="empty-state"><div class="empty-state-icon">\u{1F4CB}</div><p class="empty-state-text">Noch keine Besuche erfasst</p></div>'
        : `<ul class="activity-list">
            ${letzteAktivitaeten.map(b => `
              <a href="#/besuche" class="activity-item">
                <div class="activity-icon">\u{1F4CB}</div>
                <div class="activity-text">
                  <div><strong>${escapeHtml(b.firma)}</strong> - ${escapeHtml(b.besuchstyp || b.typ || 'Besuch')}</div>
                  <div class="activity-time">${formatDatum(b.datum)} \u00B7 ${escapeHtml(b.ort || '')}</div>
                </div>
              </a>
            `).join('')}
          </ul>`
      }
    </div>
  `;
}
