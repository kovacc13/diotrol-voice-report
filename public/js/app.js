// ========================================
// App-Initialisierung
// ========================================

// Routen registrieren
Router.on('/dashboard', renderDashboard);
Router.on('/neuer-besuch', renderNeuerBesuch);
Router.on('/besuche', renderBesuche);
Router.on('/kunden', renderKunden);
Router.on('/wochenbericht', renderWochenbericht);

// Router starten
document.addEventListener('DOMContentLoaded', () => {
  Router.init();
});

// Keyboard-Shortcut: ESC schliesst Modals und Lightbox
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    // Lightbox schliessen
    closeLightbox();
    // Modal schliessen
    const modal = document.querySelector('.modal-overlay');
    if (modal) modal.remove();
    // Export-Dropdown schliessen
    const dd = document.getElementById('exportDropdown');
    if (dd) dd.classList.remove('show');
  }
});

// Klick ausserhalb der Sidebar (Mobile)
document.addEventListener('click', (e) => {
  const sidebar = document.getElementById('sidebar');
  const hamburger = document.getElementById('hamburgerBtn');
  if (sidebar && sidebar.classList.contains('open') &&
      !sidebar.contains(e.target) &&
      !hamburger.contains(e.target)) {
    toggleSidebar();
  }
});
