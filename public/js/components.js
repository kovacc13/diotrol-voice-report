// ========================================
// Wiederverwendbare UI-Komponenten
// ========================================

// Toast-Benachrichtigung
function showToast(message, type = 'success') {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(20px)';
    toast.style.transition = 'all 0.3s';
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}

// Lightbox
function openLightbox(url) {
  const lb = document.getElementById('lightbox');
  const img = document.getElementById('lightboxImg');
  img.src = url;
  lb.classList.add('show');
}

function closeLightbox() {
  document.getElementById('lightbox').classList.remove('show');
}

// Mobile Sidebar Toggle
function toggleSidebar() {
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('sidebarOverlay');
  sidebar.classList.toggle('open');
  overlay.classList.toggle('show');
}

// Text zu Bullet-Point HTML konvertieren
function textToBullets(text) {
  if (!text) return '';
  // Verschiedene Bullet-Formate unterstuetzen
  const lines = text.split('\n')
    .map(l => l.trim())
    .filter(l => l.length > 0)
    .map(l => l.replace(/^[\u2022\-\*]\s*/, '')); // Bullet-Zeichen entfernen

  if (lines.length === 0) return '';
  if (lines.length === 1 && !text.includes('\n')) return `<p>${escapeHtml(text)}</p>`;

  return '<ul class="bullet-list">' +
    lines.map(l => `<li>${escapeHtml(l)}</li>`).join('') +
    '</ul>';
}

// HTML-Escaping
function escapeHtml(text) {
  if (!text) return '';
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// Datum formatieren
function formatDatum(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  const wochentage = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
  const tag = wochentage[d.getDay()];
  return `${d.getDate().toString().padStart(2, '0')}.${(d.getMonth() + 1).toString().padStart(2, '0')}.${d.getFullYear()} (${tag})`;
}

// Kurz-Datum
function formatDatumKurz(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  return `${d.getDate().toString().padStart(2, '0')}.${(d.getMonth() + 1).toString().padStart(2, '0')}.${d.getFullYear()}`;
}

// Aktuelle Kalenderwoche berechnen
function getCurrentKW() {
  const now = new Date();
  const onejan = new Date(now.getFullYear(), 0, 1);
  return Math.ceil(((now - onejan) / 86400000 + onejan.getDay() + 1) / 7);
}

// Aktuelles Jahr
function getCurrentJahr() {
  return new Date().getFullYear();
}

// Heutiges Datum als ISO String
function getTodayISO() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

// Wochentag aus Datum
function getWochentag(dateStr) {
  if (!dateStr) return '';
  const wochentage = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
  return wochentage[new Date(dateStr).getDay()];
}

// Fotos aus JSON-String parsen
function parseFotos(fotosStr) {
  if (!fotosStr) return [];
  try {
    const parsed = JSON.parse(fotosStr);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

// Modal erstellen
function createModal(content, options = {}) {
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.onclick = (e) => {
    if (e.target === overlay) overlay.remove();
  };

  const modal = document.createElement('div');
  modal.className = 'modal';
  modal.innerHTML = `
    <button class="modal-close" onclick="this.closest('.modal-overlay').remove()">&times;</button>
    ${content}
  `;

  overlay.appendChild(modal);
  document.body.appendChild(overlay);
  return overlay;
}

// Bestaetigungsdialog
function confirmDialog(message) {
  return new Promise((resolve) => {
    const overlay = createModal(`
      <div class="confirm-dialog">
        <p>${message}</p>
        <div class="confirm-actions">
          <button class="btn btn-outline" id="confirmNo">Abbrechen</button>
          <button class="btn btn-danger" id="confirmYes">Loeschen</button>
        </div>
      </div>
    `);

    overlay.querySelector('#confirmNo').onclick = () => {
      overlay.remove();
      resolve(false);
    };
    overlay.querySelector('#confirmYes').onclick = () => {
      overlay.remove();
      resolve(true);
    };
  });
}

// Segment-Optionen
const SEGMENTE = ['Schreiner', 'Maler', 'Zimmerei', 'Fensterbau', 'Holzbau', 'Architekt', 'Holzbauingenieure', 'Hobelwerke', 'Saegereien', 'Handel', 'Diotrol-Intern', 'Office', 'Sonstiges'];

// Besuchstyp-Optionen
const BESUCHSTYPEN = ['Erstbesuch', 'Folgebesuch', 'Kaltbesuch', 'Beratung', 'Reklamation', 'Schulung', 'Messe', 'Telefonat', 'Besprechung/Meeting'];

// Typ-Optionen
const TYPEN = ['Kundenbesuch', 'Intern', 'Office', 'Reise'];

// Status-Optionen
const STATUS_OPTIONEN = ['Aktiv', 'Inaktiv', 'Interessent', 'Verloren'];

// Select-Optionen als HTML
function selectOptions(options, selected = '') {
  return options.map(o =>
    `<option value="${o}" ${o === selected ? 'selected' : ''}>${o}</option>`
  ).join('');
}
