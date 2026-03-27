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

  main.innerHTML = `
    <div class="page-header">
      <div>
        <h1>Dashboard</h1>
        <p class="subtitle">Willkommen zurueck. Hier ist Ihre Uebersicht.</p>
      </div>
      <a href="#/neuer-besuch" class="btn btn-orange">+ Neuer Besuch</a>
    </div>

    <div class="stat-cards">
      <div class="stat-card">
        <div class="stat-icon">\u{1F4C5}</div>
        <div class="stat-info">
          <div class="stat-label">Besuche KW ${kw}</div>
          <div class="stat-value">${besucheKW.length}</div>
        </div>
      </div>
      <div class="stat-card">
        <div class="stat-icon">\u{1F4C8}</div>
        <div class="stat-info">
          <div class="stat-label">Besuche gesamt</div>
          <div class="stat-value">${besuche.length}</div>
        </div>
      </div>
      <div class="stat-card">
        <div class="stat-icon">\u{1F465}</div>
        <div class="stat-info">
          <div class="stat-label">Kunden gesamt</div>
          <div class="stat-value">${kunden.length}</div>
        </div>
      </div>
      <div class="stat-card">
        <div class="stat-icon">\u{1F4CB}</div>
        <div class="stat-info">
          <div class="stat-label">Offene Projekte</div>
          <div class="stat-value">${offeneProjekte}</div>
        </div>
      </div>
    </div>

    <div class="card">
      <h3 class="card-title">Letzte Aktivitaeten</h3>
      ${letzteAktivitaeten.length === 0
        ? '<div class="empty-state"><div class="empty-state-icon">\u{1F4CB}</div><p class="empty-state-text">Noch keine Besuche erfasst</p></div>'
        : `<ul class="activity-list">
            ${letzteAktivitaeten.map(b => `
              <li class="activity-item">
                <div class="activity-icon">\u{1F4CB}</div>
                <div class="activity-text">
                  <div><strong>${escapeHtml(b.firma)}</strong> - ${escapeHtml(b.besuchstyp || b.typ || 'Besuch')}</div>
                  <div class="activity-time">${formatDatum(b.datum)} \u00B7 ${escapeHtml(b.ort || '')}</div>
                </div>
              </li>
            `).join('')}
          </ul>`
      }
    </div>
  `;
}
