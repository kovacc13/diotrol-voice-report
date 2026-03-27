// ========================================
// Einfacher Hash-Router
// ========================================

const Router = {
  routes: {},
  currentPage: null,

  // Route registrieren
  on(path, handler) {
    this.routes[path] = handler;
  },

  // Aktuelle Route ermitteln und ausfuehren
  async navigate() {
    const hash = window.location.hash || '#/dashboard';
    // Route und eventuelle Parameter trennen
    const [path, ...paramParts] = hash.split('/').slice(1);
    const route = '/' + path;
    const params = paramParts.join('/');

    // Sidebar aktiven Punkt aktualisieren
    document.querySelectorAll('.nav-item').forEach(item => {
      item.classList.remove('active');
      if (item.dataset.page === path) {
        item.classList.add('active');
      }
    });

    // Mobile Sidebar schliessen
    const sidebar = document.getElementById('sidebar');
    const overlay = document.getElementById('sidebarOverlay');
    if (sidebar) sidebar.classList.remove('open');
    if (overlay) overlay.classList.remove('show');

    // Route ausfuehren
    const handler = this.routes[route];
    if (handler) {
      this.currentPage = path;
      const mainContent = document.getElementById('mainContent');
      mainContent.innerHTML = '<div class="loading"><span class="spinner"></span> Laden...</div>';
      try {
        await handler(params);
      } catch (error) {
        console.error('Seitenfehler:', error);
        mainContent.innerHTML = `
          <div class="empty-state">
            <div class="empty-state-icon">⚠️</div>
            <p class="empty-state-text">Fehler beim Laden: ${error.message}</p>
          </div>
        `;
      }
    } else {
      // Fallback: Dashboard
      window.location.hash = '#/dashboard';
    }
  },

  // Router starten
  init() {
    window.addEventListener('hashchange', () => this.navigate());
    this.navigate();
  }
};
