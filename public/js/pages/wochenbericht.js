// ========================================
// Wochenbericht-Seite
// ========================================

let wochenberichtData = null;

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
          </div>
        </div>
      </div>
    </div>

    <div class="stat-cards" id="wbStats">
      <div class="stat-card">
        <div class="stat-icon">\u{1F4CB}</div>
        <div class="stat-info">
          <div class="stat-label">Total Besuche</div>
          <div class="stat-value" id="wbTotal">-</div>
        </div>
      </div>
      <div class="stat-card">
        <div class="stat-icon">\u{1F195}</div>
        <div class="stat-info">
          <div class="stat-label">Neukunden</div>
          <div class="stat-value" id="wbNeukunden">-</div>
        </div>
      </div>
      <div class="stat-card">
        <div class="stat-icon">\u{1F504}</div>
        <div class="stat-info">
          <div class="stat-label">Bestandskunden</div>
          <div class="stat-value" id="wbBestandskunden">-</div>
        </div>
      </div>
      <div class="stat-card">
        <div class="stat-icon">\u{1F4E6}</div>
        <div class="stat-info">
          <div class="stat-label">Fokus-Produkte</div>
          <div class="stat-value" id="wbProdukte">-</div>
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
      // Aktuelle KW hinzufuegen falls nicht vorhanden
      if (!data.verfuegbareKWs.includes(getCurrentKW())) {
        kwSelect.innerHTML = `<option value="${getCurrentKW()}">KW ${getCurrentKW()} (aktuell)</option>` + kwSelect.innerHTML;
      }
    }

    // Stats aktualisieren
    const stats = data.stats || {};
    document.getElementById('wbTotal').textContent = stats.totalBesuche || 0;
    document.getElementById('wbNeukunden').textContent = stats.neukunden || 0;
    document.getElementById('wbBestandskunden').textContent = stats.bestandskunden || 0;
    document.getElementById('wbProdukte').textContent = stats.fokusProdukte || '-';

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

  container.innerHTML = besuche.map(b => `
    <div class="wb-besuch">
      <div class="wb-besuch-header">
        <span class="wb-besuch-firma">${escapeHtml(b.firma)}</span>
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
          ${b.naechsteSchritte ? `<div class="wb-next-steps"><strong>Neu:</strong> ${textToBullets(b.naechsteSchritte)}</div>` : ''}
        </div>
      </div>
    </div>
  `).join('');
}

// ---- Export ----

// Segment-Icons (Kuerzel fuer PDF-Kreise)
const SEGMENT_ICONS = {
  'Schreiner': 'T', 'Maler': 'M', 'Zimmerei': 'H',
  'Fensterbau': 'F', 'Fassadenbau': 'Fa', 'Architekt': 'A',
  'Generalunternehmer': 'GU', 'Handel': 'D', 'Sonstiges': 'X'
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
  const followUps = [];
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
      followUps.push({ firma: b.firma || 'Intern', text: b.naechsteSchritte, datum: b.datum });
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

  return { firmen, produkteAll, followUps, marktinfos, segmente, topProdukte };
}

// ============================================================
// PDF EXPORT (jsPDF)
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
  const TEXT_DUNKEL = [26, 26, 26];
  const TEXT_GRAU = [107, 114, 128];
  const WEISS = [255, 255, 255];
  const BORDER = [229, 231, 235];

  let y = 0;

  function checkPage(needed) {
    if (y + needed > 275) {
      pdf.addPage();
      // Mini-Header ab Seite 2
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
  // Akzent-Streifen
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

  // Metriken-Kacheln (groesser, mehr Platz)
  const metriken = [
    [String(stats.totalBesuche || 0), 'Aktivitaeten'],
    [String(stats.neukunden || 0), 'Neukunden'],
    [String(stats.bestandskunden || 0), 'Bestandskunden'],
    [String(summary.firmen.size), 'Kunden besucht'],
  ];
  const kW = 42, kH = 32, kGap = 3;

  metriken.forEach(([zahl, label], i) => {
    const x = 15 + i * (kW + kGap);
    // Kachel
    pdf.setFillColor(...WEISS);
    pdf.setDrawColor(...BORDER);
    pdf.setLineWidth(0.4);
    pdf.rect(x, y, kW, kH, 'DF');
    // Gruener Akzent oben
    pdf.setFillColor(...GRUEN);
    pdf.rect(x, y, kW, 2.5, 'F');
    // Zahl
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(24);
    pdf.setTextColor(...GRUEN);
    pdf.text(zahl, x + kW / 2, y + 17, { align: 'center' });
    // Label
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(7);
    pdf.setTextColor(...TEXT_GRAU);
    pdf.text(label, x + kW / 2, y + 25, { align: 'center' });
  });
  y += kH + 10;

  // Zwei-Spalten: Segmente + Top-Produkte
  const spY = y;
  // Links: Segmente
  if (Object.keys(summary.segmente).length > 0) {
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(10);
    pdf.setTextColor(...GRUEN_DUNKEL);
    pdf.text('Segmentverteilung', 16, y);
    y += 7;
    Object.entries(summary.segmente)
      .sort((a, b) => b[1] - a[1])
      .forEach(([seg, cnt]) => {
        const k = SEGMENT_ICONS[seg] || 'X';
        pdf.setFillColor(...GRUEN);
        pdf.circle(20, y - 1, 2.5, 'F');
        pdf.setFont('helvetica', 'bold');
        pdf.setFontSize(5);
        pdf.setTextColor(...WEISS);
        pdf.text(k, 20, y - 0.2, { align: 'center' });
        pdf.setFont('helvetica', 'normal');
        pdf.setFontSize(9);
        pdf.setTextColor(...TEXT_DUNKEL);
        pdf.text(`${seg} (${cnt})`, 25, y);
        y += 6;
      });
  }

  // Rechts: Top-Produkte
  let prodY = spY;
  if (summary.topProdukte.length > 0) {
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(10);
    pdf.setTextColor(...GRUEN_DUNKEL);
    pdf.text('Top Produkte', 110, prodY);
    prodY += 7;
    summary.topProdukte.forEach(([prod, cnt]) => {
      // Gruener Punkt statt Unicode-Bullet
      pdf.setFillColor(...GRUEN);
      pdf.circle(113, prodY - 1.2, 1.2, 'F');
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(9);
      pdf.setTextColor(...TEXT_DUNKEL);
      pdf.text(`${prod} (${cnt}x)`, 117, prodY);
      prodY += 6;
    });
  }
  y = Math.max(y, prodY) + 10;

  // === OFFENE TO-DOs ===
  if (summary.followUps.length > 0) {
    checkPage(30);

    // Ueberschrift mit Akzent-Hintergrund (kein Unicode-Emoji!)
    pdf.setFillColor(...AKZENT_HELL);
    pdf.rect(15, y, 180, 10, 'F');
    pdf.setFillColor(...AKZENT);
    pdf.rect(15, y, 3, 10, 'F');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(11);
    pdf.setTextColor(146, 64, 14);
    pdf.text('OFFENE TO-DOs & FOLLOW-UPS', 22, y + 7);
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

    summary.followUps.slice(0, 10).forEach((fu, idx) => {
      checkPage(9);
      // Abwechselnde Zeilenhintergruende
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

    // Besuche
    tage[tag].forEach((b, idx) => {
      checkPage(25);

      // Icon-Kreis + Firma
      const seg = b.segment || 'Sonstiges';
      const ik = SEGMENT_ICONS[seg] || 'X';
      pdf.setFillColor(...GRUEN);
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

      // Kontakt + Segment rechts
      if (b.kontaktperson) {
        pdf.setFont('helvetica', 'normal');
        pdf.setFontSize(8);
        pdf.setTextColor(...TEXT_GRAU);
        pdf.text(b.kontaktperson, 25 + pdf.getTextWidth(firmaText) + 3, y + 1);
      }
      if (b.segment) {
        pdf.setFontSize(7);
        pdf.setTextColor(...TEXT_GRAU);
        pdf.text(b.segment, 190, y + 1, { align: 'right' });
      }
      y += 7;

      // Detail-Felder (identisch mit Word-Export)
      // marker: [hintergrund, streifen-links] oder null
      const felder = [
        ['Themen', b.themen, null],
        ['Produkte', b.produkte, null],
        ['Ergebnis', b.ergebnis, [GRUEN_HELL, GRUEN]],
        ['Naechste Schritte', b.naechsteSchritte, [AKZENT_HELL, AKZENT]],
        ['Stellungnahme', b.stellungnahme, null],
      ];

      felder.forEach(([label, wert, highlight]) => {
        if (!wert) return;
        checkPage(12);

        const yVor = y;

        // Label
        const labelColor = highlight ? highlight[1] : GRUEN_DUNKEL;
        pdf.setFont('helvetica', 'bold');
        pdf.setFontSize(8);
        pdf.setTextColor(...labelColor);
        pdf.text(`${label}:`, 25, y);
        y += 4.5;

        // Wert (mehrzeilig)
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

        // Hintergrund + Streifen links (wie Word .highlight / .todo-highlight)
        if (highlight) {
          const blockH = y - yVor + 1;
          // Hintergrund
          pdf.setFillColor(...highlight[0]);
          pdf.rect(21, yVor - 3, 172, blockH, 'F');
          // Streifen links
          pdf.setFillColor(...highlight[1]);
          pdf.rect(21, yVor - 3, 2, blockH, 'F');
          // Text nochmal drueber (weil Hintergrund den Text ueberdeckt)
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

  // Footer auf jeder Seite: "Seite X von Y"
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

  // Download
  pdf.save(`Wochenbericht_KW${data.kw}_${data.jahr}.pdf`);
  showToast('PDF-Export heruntergeladen', 'success');
}

// ============================================================
// WORD EXPORT (mit Executive Summary)
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
    'Generalunternehmer': '\u{1F3E2}', 'Handel': '\u{1F3EA}', 'Sonstiges': '\u{1F4CB}'
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
.header-bg { background-color: #1a5c3a; color: white; padding: 12pt 16pt; }
.header-bg h1 { color: white; margin: 0; }
.header-bg p { color: #c8e6d2; margin: 4pt 0 0 0; }
.accent-bar { background-color: #f5a623; height: 4pt; }
.stat-table td { text-align: center; background: #e8f5ed; padding: 8pt; }
.stat-zahl { font-size: 20pt; font-weight: bold; color: #1a5c3a; display: block; }
.stat-label { font-size: 8pt; color: #6b7280; }
.highlight { background: #e8f5ed; padding: 6pt 10pt; border-left: 4px solid #1a5c3a; margin: 6pt 0; }
.todo-highlight { background: #fef9c3; padding: 6pt 10pt; border-left: 4px solid #f5a623; margin: 6pt 0; }
.todo-title { background: #fef9c3; padding: 8pt; font-size: 13pt; font-weight: bold; color: #92400e; }
.label-text { font-weight: bold; color: #0e3d25; font-size: 9pt; }
.muted { color: #6b7280; font-size: 9pt; }
.segment-badge { color: #6b7280; font-size: 9pt; }
</style></head><body>`;

  // Footer-Definition (Seitenzahlen fuer Word)
  html += `<div style="mso-element:footer" id="f1">
    <p class="MsoFooter"><span style="color:#6b7280;">Daniel Pfister | Diotrol AG | Vertraulich</span>
    <span style="float:right; color:#6b7280;">Seite <span style="mso-field-code:'PAGE'"></span><!--[if gte mso 9]><span style="mso-element:field-begin"></span> PAGE <span style="mso-element:field-end"></span><![endif]--></span></p>
  </div>`;

  // Content in Section wrappen
  html += '<div class="Section1">';

  // Header
  html += `<div class="header-bg">
    <h1>Daniel Pfister</h1>
    <p>Kundenbesuchsbericht | Diotrol AG | Aussendienst Schweiz</p>
  </div>
  <div class="accent-bar"></div>`;

  // Titel + Info
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

  // Segmente + Produkte nebeneinander
  html += '<table><tr><td style="width:50%; vertical-align:top;">';
  if (Object.keys(summary.segmente).length > 0) {
    html += '<p class="label-text">Segmentverteilung:</p>';
    Object.entries(summary.segmente).sort((a,b) => b[1]-a[1]).forEach(([seg, cnt]) => {
      const emoji = SEG_EMOJI[seg] || '\u{1F4CB}';
      html += `<p style="margin:2pt 0 2pt 8pt; font-size:9pt;">${emoji} ${seg} (${cnt})</p>`;
    });
  }
  html += '</td><td style="width:50%; vertical-align:top;">';
  if (summary.topProdukte.length > 0) {
    html += '<p class="label-text">Top Produkte:</p>';
    summary.topProdukte.forEach(([prod, cnt]) => {
      html += `<p style="margin:2pt 0 2pt 8pt; font-size:9pt; color:#1a5c3a;">\u25CF ${prod} (${cnt}x)</p>`;
    });
  }
  html += '</td></tr></table>';

  // To-Dos / Follow-ups
  if (summary.followUps.length > 0) {
    html += `<div class="todo-title">\u26A1 OFFENE TO-DOs & FOLLOW-UPS</div>`;
    html += `<table><tr><th>Kunde</th><th>Aktion</th><th>Datum</th></tr>`;
    summary.followUps.forEach(fu => {
      const txt = fu.text.length > 100 ? fu.text.substring(0, 100) + '...' : fu.text;
      html += `<tr>
        <td style="font-weight:bold; color:#0e3d25; font-size:9pt; width:20%;">${escapeHtml(fu.firma)}</td>
        <td style="font-size:9pt;">${escapeHtml(txt.replace(/\n/g, ' '))}</td>
        <td style="font-size:9pt; color:#6b7280; width:12%;">${formatDatumKurz(fu.datum)}</td>
      </tr>`;
    });
    html += '</table>';
  }

  // Seitenumbruch + Detailbericht
  html += '<br clear="all" style="page-break-before:always">';
  html += `<h2>\u{1F4C5} DETAILBERICHT KW ${data.kw}</h2>`;

  // Nach Tag gruppieren
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

    tage[tag].forEach(b => {
      const segEmoji = SEG_EMOJI[b.segment] || '\u{1F4CB}';
      html += `<h3>${segEmoji} ${escapeHtml(b.firma || 'Unbekannt')}`;
      if (b.kontaktperson) html += ` <span class="muted">| ${escapeHtml(b.kontaktperson)}</span>`;
      html += `</h3>`;
      if (b.segment) html += `<p class="segment-badge">${escapeHtml(b.segment)}</p>`;

      if (b.themen) html += `<p><span class="label-text">\u{1F4AC} Themen:</span><br>${escapeHtml(b.themen).replace(/\n/g,'<br>')}</p>`;
      if (b.produkte) html += `<p><span class="label-text">\u{1F4E6} Produkte:</span> ${escapeHtml(b.produkte)}</p>`;
      if (b.ergebnis) html += `<div class="highlight"><span class="label-text">\u{1F3AF} Ergebnis:</span><br>${escapeHtml(b.ergebnis).replace(/\n/g,'<br>')}</div>`;
      if (b.naechsteSchritte) html += `<div class="todo-highlight"><span class="label-text">\u26A1 Naechste Schritte:</span><br>${escapeHtml(b.naechsteSchritte).replace(/\n/g,'<br>')}</div>`;
      if (b.stellungnahme) html += `<p class="muted"><span class="label-text">\u{1F4DD} Stellungnahme:</span><br>${escapeHtml(b.stellungnahme).replace(/\n/g,'<br>')}</p>`;
      html += '<hr style="border:none; border-top:1px solid #e5e7eb; margin:8pt 0;">';
    });
  });

  // Section schliessen
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

  // To-Dos
  if (summary.followUps.length > 0) {
    text += `OFFENE TO-DOs\n`;
    text += `${'-'.repeat(40)}\n`;
    summary.followUps.forEach(fu => {
      text += `[ ] ${fu.firma}: ${fu.text.replace(/\n/g, ' ')}\n`;
    });
    text += '\n';
  }

  text += `${'='.repeat(50)}\nDETAILS\n${'='.repeat(50)}\n\n`;

  (data.besuche || []).forEach(b => {
    text += `--- ${b.firma} ---\n`;
    text += `Datum: ${formatDatum(b.datum)}\n`;
    if (b.ort) text += `Ort: ${b.ort}\n`;
    if (b.kontaktperson) text += `Kontakt: ${b.kontaktperson}\n`;
    if (b.besuchstyp) text += `Besuchstyp: ${b.besuchstyp}\n`;
    if (b.produkte) text += `Produkte: ${b.produkte}\n`;
    if (b.themen) text += `\nThemen:\n${b.themen}\n`;
    if (b.ergebnis) text += `\nErgebnis:\n${b.ergebnis}\n`;
    if (b.naechsteSchritte) text += `\nNaechste Schritte:\n${b.naechsteSchritte}\n`;
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
