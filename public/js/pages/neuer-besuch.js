// ========================================
// Neuer Besuch / Besuch bearbeiten
// Mit Pause-Taste beim Diktieren
// ========================================

// Globale Variablen fuer diese Seite
let besuchFormState = {
  transkript: '',
  fotos: [], // Array von { url, thumbnail, public_id }
  editId: null,
  editData: null,
  produkte: [],
  mediaRecorder: null,
  isRecording: false,
  isPaused: false,
  recognition: null,
  recordingTimer: null,
  recordingSeconds: 0
};

async function renderNeuerBesuch(params) {
  const main = document.getElementById('mainContent');
  besuchFormState = {
    transkript: '',
    fotos: [],
    editId: null,
    editData: null,
    produkte: [],
    mediaRecorder: null,
    isRecording: false,
    isPaused: false,
    recognition: null,
    recordingTimer: null,
    recordingSeconds: 0
  };

  // Pruefen ob Bearbeitungsmodus
  let editId = null;
  let prefillKundeId = null;

  if (params) {
    if (params.startsWith('edit/')) {
      editId = params.replace('edit/', '');
    } else if (params.startsWith('kunde/')) {
      prefillKundeId = params.replace('kunde/', '');
    }
  }

  if (editId) {
    try {
      besuchFormState.editData = await API.besuche.einzeln(editId);
      besuchFormState.editId = editId;
      besuchFormState.fotos = parseFotos(besuchFormState.editData.fotos);
      if (besuchFormState.editData.produkte) {
        besuchFormState.produkte = besuchFormState.editData.produkte.split(',').map(p => p.trim()).filter(Boolean);
      }
    } catch (e) {
      showToast('Besuch konnte nicht geladen werden', 'error');
    }
  }

  let prefillKunde = null;
  if (prefillKundeId) {
    try {
      prefillKunde = await API.kunden.einzeln(prefillKundeId);
    } catch (e) { /* ignorieren */ }
  }

  const d = besuchFormState.editData;
  const isEdit = !!d;

  main.innerHTML = `
    <div class="page-header">
      <div>
        <h1>${isEdit ? 'Besuch bearbeiten' : 'Neuer Besuch'}</h1>
        <p class="subtitle">${isEdit ? 'Besuchsdaten aendern' : 'Besuch diktieren, pruefen und speichern'}</p>
      </div>
    </div>

    <div class="steps">
      <div class="step active" id="step1Indicator"></div>
      <div class="step" id="step2Indicator"></div>
      <div class="step" id="step3Indicator"></div>
      <div class="step" id="step4Indicator"></div>
    </div>

    <!-- Schritt 1: Diktieren -->
    <div class="form-section" id="step1Section">
      <div class="form-section-title">\u{1F3A4} Schritt 1: Besuch diktieren</div>
      <div class="diktat-section">
        <!-- Mikrofon-Button (sichtbar wenn nicht aufnimmt) -->
        <div id="micStartArea">
          <button class="mic-btn" id="micBtn" onclick="startRecording()">\u{1F3A4}</button>
          <div class="diktat-hint" id="diktatHint">Antippen zum Aufnehmen</div>
        </div>

        <!-- Aufnahme-Steuerung (sichtbar waehrend Aufnahme) -->
        <div id="recordingArea" class="hidden">
          <div class="recording-controls">
            <button class="btn-stop" onclick="stopRecording()">\u23F9\uFE0F Stopp</button>
            <button class="btn-pause" id="pauseBtn" onclick="togglePause()">\u23F8\uFE0F Pause</button>
          </div>
          <div class="recording-status" id="recordingStatus">
            <span class="rec-dot" id="recDot"></span>
            <span id="recordingLabel">Aufnahme laeuft...</span>
            <span class="recording-timer" id="recordingTimer">00:00</span>
          </div>
        </div>

        <textarea class="diktat-textarea" id="transkriptText" placeholder="Hier erscheint der transkribierte Text. Sie koennen ihn auch direkt eingeben oder bearbeiten...">${isEdit && d.themen ? d.themen : ''}</textarea>
        <div class="diktat-actions">
          <button class="btn btn-outline" onclick="nochmalDiktieren()">\u{1F3A4} Nochmal diktieren (anhaengen)</button>
          <button class="btn btn-orange" onclick="kiAnalysieren()" id="analyseBtn">\u{1F52E} KI Analysieren & Ausfuellen</button>
        </div>
      </div>
    </div>

    <!-- Schritt 2: Daten pruefen -->
    <div class="form-section" id="step2Section">
      <div class="form-section-title">\u{1F4DD} Schritt 2: Daten pruefen und ergaenzen</div>

      <div class="form-row">
        <div class="form-group">
          <label>Firma *</label>
          <input type="text" id="feldFirma" value="${escapeHtml(d?.firma || prefillKunde?.firma || '')}" required>
        </div>
        <div class="form-group">
          <label>Ort</label>
          <input type="text" id="feldOrt" value="${escapeHtml(d?.ort || prefillKunde?.ort || '')}">
        </div>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label>Segment</label>
          <select id="feldSegment">
            <option value="">-- Waehlen --</option>
            ${selectOptions(SEGMENTE, d?.segment || prefillKunde?.segment || '')}
          </select>
        </div>
        <div class="form-group">
          <label>Kontaktperson</label>
          <input type="text" id="feldKontakt" value="${escapeHtml(d?.kontaktperson || prefillKunde?.hauptkontakt || '')}">
        </div>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label>Datum</label>
          <input type="date" id="feldDatum" value="${d?.datum || getTodayISO()}" onchange="updateWochentag()">
        </div>
        <div class="form-group">
          <label>Wochentag</label>
          <input type="text" id="feldWochentag" value="${d?.tag || getWochentag(d?.datum || getTodayISO())}" readonly style="background:#f9fafb;">
        </div>
        <div class="form-group">
          <label>Besuchstyp</label>
          <select id="feldBesuchstyp">
            <option value="">-- Waehlen --</option>
            ${selectOptions(BESUCHSTYPEN, d?.besuchstyp || '')}
          </select>
        </div>
      </div>

      <div class="form-group mb-16">
        <label>Besprochene Produkte</label>
        <div class="tag-input-container" id="produkteContainer" onclick="document.getElementById('produkteInput').focus()">
          ${besuchFormState.produkte.map(p => `<span class="tag">${escapeHtml(p)}<span class="tag-remove" onclick="removeProdukt(this, '${escapeHtml(p)}')">&times;</span></span>`).join('')}
          <input type="text" class="tag-input" id="produkteInput" placeholder="Produkt eingeben + Enter" onkeydown="handleProduktKey(event)">
        </div>
      </div>

      <div class="form-group mb-16">
        <label>Themen / Inhalt</label>
        <textarea id="feldThemen" rows="4">${escapeHtml(d?.themen || '')}</textarea>
      </div>

      <div class="form-group mb-16">
        <label>Ergebnis</label>
        <textarea id="feldErgebnis" rows="3">${escapeHtml(d?.ergebnis || '')}</textarea>
      </div>

      <div class="form-group mb-16">
        <label>Naechste Schritte</label>
        <textarea id="feldSchritte" rows="3">${escapeHtml(d?.naechsteSchritte || '')}</textarea>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label>Stellungnahme (optional)</label>
          <textarea id="feldStellungnahme" rows="2">${escapeHtml(d?.stellungnahme || '')}</textarea>
        </div>
        <div class="form-group">
          <label>Projekt (optional)</label>
          <input type="text" id="feldProjekt" value="${escapeHtml(d?.projekt || '')}">
        </div>
      </div>
    </div>

    <!-- Schritt 3: Fotos -->
    <div class="form-section" id="step3Section">
      <div class="form-section-title">\u{1F4F7} Schritt 3: Fotos</div>
      <div class="fotos-section" id="fotosDropzone" onclick="document.getElementById('fotoInput').click()">
        <input type="file" id="fotoInput" accept="image/*" multiple style="display:none" onchange="handleFotoUpload(event)">
        <div>\u{1F4F7} Fotos hinzufuegen</div>
        <div style="font-size:13px; color:var(--text-light); margin-top:4px;">Klicken oder Dateien hierher ziehen</div>
      </div>
      <div class="fotos-grid" id="fotosGrid">
        ${besuchFormState.fotos.map((f, i) => `
          <div class="foto-thumb" data-index="${i}">
            <img src="${typeof f === 'string' ? f : f.thumbnail || f.url}" onclick="openLightbox('${typeof f === 'string' ? f : f.url}')" alt="Foto">
            <button class="foto-thumb-remove" onclick="removeFoto(${i})">&times;</button>
          </div>
        `).join('')}
      </div>
    </div>

    <!-- Schritt 4: Speichern -->
    <div class="form-section text-center" id="step4Section">
      <div class="form-section-title" style="justify-content:center;">\u2705 Schritt 4: Speichern</div>
      <button class="btn btn-green" onclick="besuchSpeichern()" id="speichernBtn" style="font-size:16px; padding:14px 40px;">
        ${isEdit ? '\u{1F4BE} Aenderungen speichern' : '\u2705 Besuch speichern'}
      </button>
    </div>
  `;

  // Drag & Drop fuer Fotos
  setupFotoDragDrop();
  updateStepIndicators();
}

// Schritt-Indikatoren aktualisieren
function updateStepIndicators() {
  // Einfach: alle Steps als "done" markieren wenn Text vorhanden
  const has = (id) => document.getElementById(id)?.value?.trim();

  const s1 = document.getElementById('step1Indicator');
  const s2 = document.getElementById('step2Indicator');
  const s3 = document.getElementById('step3Indicator');
  const s4 = document.getElementById('step4Indicator');

  if (s1) s1.className = 'step active';
  if (s2) s2.className = has('feldFirma') ? 'step done' : 'step';
  if (s3) s3.className = besuchFormState.fotos.length > 0 ? 'step done' : 'step';
  if (s4) s4.className = 'step';
}

// ---- Sprachaufnahme mit OpenAI Whisper + Pause ----

async function startRecording() {
  const micStartArea = document.getElementById('micStartArea');
  const recordingArea = document.getElementById('recordingArea');

  try {
    // Mikrofon-Zugriff anfordern
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });

    // MediaRecorder fuer Audio-Aufnahme (wird an Whisper gesendet)
    const mediaRecorder = new MediaRecorder(stream, { mimeType: 'audio/webm' });
    const audioChunks = [];

    mediaRecorder.ondataavailable = (event) => {
      if (event.data.size > 0) {
        audioChunks.push(event.data);
      }
    };

    mediaRecorder.onstop = async () => {
      // Mikrofon-Stream stoppen
      stream.getTracks().forEach(track => track.stop());

      // Audio-Blob erstellen und an Whisper senden
      const audioBlob = new Blob(audioChunks, { type: 'audio/webm' });

      if (audioBlob.size < 1000) {
        showToast('Aufnahme zu kurz. Bitte erneut versuchen.', 'error');
        resetRecordingUI();
        return;
      }

      // Textarea auf "transkribiere..." setzen
      const textarea = document.getElementById('transkriptText');
      const hinweis = document.getElementById('diktatHint');
      if (hinweis) hinweis.textContent = 'Wird transkribiert...';

      try {
        showToast('Transkribiere mit OpenAI Whisper...', 'info');

        const formData = new FormData();
        formData.append('file', audioBlob, 'aufnahme.webm');

        const response = await fetch('/api/transkribieren', {
          method: 'POST',
          body: formData
        });

        const data = await response.json();

        if (data.text) {
          // Text anhaengen (nicht ersetzen)
          const existing = textarea.value;
          const separator = existing && !existing.endsWith(' ') ? ' ' : '';
          textarea.value = existing + separator + data.text;
          besuchFormState.transkript = textarea.value;
          showToast('Transkription erfolgreich!', 'success');
        } else if (data.fallback === 'browser_speech_api') {
          showToast('OpenAI Key nicht konfiguriert. Bitte OPENAI_API_KEY in Netlify setzen.', 'error');
        } else {
          showToast('Transkription fehlgeschlagen: ' + (data.error || 'Unbekannter Fehler'), 'error');
        }
      } catch (err) {
        console.error('Transkriptions-Fehler:', err);
        showToast('Transkription fehlgeschlagen. Netzwerkfehler.', 'error');
      }

      if (hinweis) hinweis.textContent = 'Antippen zum Aufnehmen';
      resetRecordingUI();
    };

    // State setzen
    besuchFormState.mediaRecorder = mediaRecorder;
    besuchFormState.transkript = document.getElementById('transkriptText').value;
    besuchFormState.isRecording = true;
    besuchFormState.isPaused = false;
    besuchFormState.recordingSeconds = 0;

    // Aufnahme starten
    mediaRecorder.start(1000); // Chunks alle 1 Sekunde

    // UI umschalten
    micStartArea.classList.add('hidden');
    recordingArea.classList.remove('hidden');

    // Timer starten
    startRecordingTimer();

  } catch (err) {
    console.error('Mikrofon-Fehler:', err);
    showToast('Kein Mikrofon-Zugriff. Bitte Berechtigung erteilen.', 'error');
  }
}

function stopRecording() {
  if (besuchFormState.mediaRecorder && besuchFormState.mediaRecorder.state !== 'inactive') {
    besuchFormState.mediaRecorder.stop();
  }

  besuchFormState.isRecording = false;
  besuchFormState.isPaused = false;

  // Timer stoppen
  stopRecordingTimer();
}

function togglePause() {
  if (!besuchFormState.isRecording || !besuchFormState.mediaRecorder) return;

  const pauseBtn = document.getElementById('pauseBtn');
  const recDot = document.getElementById('recDot');
  const recordingLabel = document.getElementById('recordingLabel');

  if (besuchFormState.isPaused) {
    // Fortsetzen
    besuchFormState.isPaused = false;
    besuchFormState.mediaRecorder.resume();

    pauseBtn.className = 'btn-pause';
    pauseBtn.innerHTML = '\u23F8\uFE0F Pause';
    recDot.classList.remove('paused');
    recordingLabel.textContent = 'Aufnahme laeuft...';

    // Timer fortsetzen
    startRecordingTimer();

  } else {
    // Pausieren
    besuchFormState.isPaused = true;
    besuchFormState.mediaRecorder.pause();

    pauseBtn.className = 'btn-resume';
    pauseBtn.innerHTML = '\u25B6\uFE0F Weiter';
    recDot.classList.add('paused');
    recordingLabel.textContent = 'Pausiert...';

    // Timer pausieren
    stopRecordingTimer();
  }
}

function resetRecordingUI() {
  besuchFormState.isRecording = false;
  besuchFormState.isPaused = false;

  const micStartArea = document.getElementById('micStartArea');
  const recordingArea = document.getElementById('recordingArea');

  if (micStartArea) micStartArea.classList.remove('hidden');
  if (recordingArea) recordingArea.classList.add('hidden');

  // Pause-Button zuruecksetzen
  const pauseBtn = document.getElementById('pauseBtn');
  if (pauseBtn) {
    pauseBtn.className = 'btn-pause';
    pauseBtn.innerHTML = '\u23F8\uFE0F Pause';
  }

  stopRecordingTimer();
}

// ---- Aufnahme-Timer ----

function startRecordingTimer() {
  stopRecordingTimer(); // Sicherheitshalber
  besuchFormState.recordingTimer = setInterval(() => {
    besuchFormState.recordingSeconds++;
    updateTimerDisplay();
  }, 1000);
}

function stopRecordingTimer() {
  if (besuchFormState.recordingTimer) {
    clearInterval(besuchFormState.recordingTimer);
    besuchFormState.recordingTimer = null;
  }
}

function updateTimerDisplay() {
  const timerEl = document.getElementById('recordingTimer');
  if (!timerEl) return;
  const mins = Math.floor(besuchFormState.recordingSeconds / 60);
  const secs = besuchFormState.recordingSeconds % 60;
  timerEl.textContent = `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
}

function nochmalDiktieren() {
  // Aktuellen Text speichern, dann nochmal aufnehmen (anhaengen)
  besuchFormState.transkript = document.getElementById('transkriptText').value;
  if (besuchFormState.transkript && !besuchFormState.transkript.endsWith(' ')) {
    besuchFormState.transkript += ' ';
    document.getElementById('transkriptText').value = besuchFormState.transkript;
  }
  startRecording();
}

// ---- KI-Analyse ----

async function kiAnalysieren() {
  const text = document.getElementById('transkriptText').value.trim();
  if (!text) {
    showToast('Bitte zuerst Text eingeben oder diktieren', 'error');
    return;
  }

  const btn = document.getElementById('analyseBtn');
  btn.disabled = true;
  btn.textContent = '\u231B Analysiere...';

  try {
    const result = await API.strukturieren(text);

    // Formularfelder befuellen
    if (result.firma) document.getElementById('feldFirma').value = result.firma;
    if (result.ort) document.getElementById('feldOrt').value = result.ort;
    if (result.kontaktperson) document.getElementById('feldKontakt').value = result.kontaktperson;
    if (result.segment) document.getElementById('feldSegment').value = result.segment;
    if (result.besuchstyp) document.getElementById('feldBesuchstyp').value = result.besuchstyp;
    if (result.themen) document.getElementById('feldThemen').value = result.themen;
    if (result.ergebnis) document.getElementById('feldErgebnis').value = result.ergebnis;
    if (result.naechsteSchritte) document.getElementById('feldSchritte').value = result.naechsteSchritte;
    if (result.stellungnahme) document.getElementById('feldStellungnahme').value = result.stellungnahme;
    if (result.projekt) document.getElementById('feldProjekt').value = result.projekt;

    // Produkte als Tags
    if (result.produkte) {
      besuchFormState.produkte = result.produkte.split(',').map(p => p.trim()).filter(Boolean);
      renderProduktTags();
    }

    showToast('KI-Analyse abgeschlossen! Bitte Daten pruefen.', 'success');
    updateStepIndicators();

    // Zum Formular scrollen
    document.getElementById('step2Section').scrollIntoView({ behavior: 'smooth' });

  } catch (error) {
    showToast('KI-Analyse fehlgeschlagen: ' + error.message, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '\u{1F52E} KI Analysieren & Ausfuellen';
  }
}

// ---- Produkt-Tags ----

function handleProduktKey(event) {
  if (event.key === 'Enter' || event.key === ',') {
    event.preventDefault();
    const input = event.target;
    const val = input.value.trim().replace(',', '');
    if (val && !besuchFormState.produkte.includes(val)) {
      besuchFormState.produkte.push(val);
      renderProduktTags();
    }
    input.value = '';
  }
}

function removeProdukt(el, name) {
  event.stopPropagation();
  besuchFormState.produkte = besuchFormState.produkte.filter(p => p !== name);
  renderProduktTags();
}

function renderProduktTags() {
  const container = document.getElementById('produkteContainer');
  const input = document.getElementById('produkteInput');
  // Alle Tags entfernen
  container.querySelectorAll('.tag').forEach(t => t.remove());
  // Neue Tags vor dem Input einfuegen
  besuchFormState.produkte.forEach(p => {
    const tag = document.createElement('span');
    tag.className = 'tag';
    tag.innerHTML = `${escapeHtml(p)}<span class="tag-remove" onclick="removeProdukt(this, '${escapeHtml(p)}')">&times;</span>`;
    container.insertBefore(tag, input);
  });
}

// ---- Wochentag Update ----

function updateWochentag() {
  const datum = document.getElementById('feldDatum').value;
  if (datum) {
    document.getElementById('feldWochentag').value = getWochentag(datum);
  }
}

// ---- Foto-Upload ----

function setupFotoDragDrop() {
  const dropzone = document.getElementById('fotosDropzone');
  if (!dropzone) return;

  ['dragenter', 'dragover'].forEach(evt => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(evt => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      uploadFotos(Array.from(files));
    }
  });
}

function handleFotoUpload(event) {
  const files = Array.from(event.target.files);
  if (files.length > 0) {
    uploadFotos(files);
  }
  event.target.value = ''; // Reset
}

async function uploadFotos(files) {
  for (const file of files) {
    if (!file.type.startsWith('image/')) continue;

    // Zu Base64 konvertieren
    const base64 = await fileToBase64(file);

    try {
      showToast('Foto wird hochgeladen...', 'info');
      const result = await API.fotoUpload(base64);
      besuchFormState.fotos.push({
        url: result.url,
        thumbnail: result.thumbnail || result.url,
        public_id: result.public_id
      });
      renderFotos();
      showToast('Foto hochgeladen!', 'success');
    } catch (error) {
      // Fallback: Bild lokal als Data-URL speichern
      besuchFormState.fotos.push({
        url: base64,
        thumbnail: base64,
        public_id: 'local_' + Date.now()
      });
      renderFotos();
      showToast('Cloudinary nicht verfuegbar, Foto lokal gespeichert', 'info');
    }
  }
  updateStepIndicators();
}

function fileToBase64(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.readAsDataURL(file);
  });
}

function renderFotos() {
  const grid = document.getElementById('fotosGrid');
  grid.innerHTML = besuchFormState.fotos.map((f, i) => {
    const thumbUrl = typeof f === 'string' ? f : (f.thumbnail || f.url);
    const fullUrl = typeof f === 'string' ? f : f.url;
    return `
      <div class="foto-thumb" data-index="${i}">
        <img src="${thumbUrl}" onclick="openLightbox('${fullUrl}')" alt="Foto ${i + 1}">
        <button class="foto-thumb-remove" onclick="removeFoto(${i})">&times;</button>
      </div>
    `;
  }).join('');
}

function removeFoto(index) {
  event.stopPropagation();
  besuchFormState.fotos.splice(index, 1);
  renderFotos();
  updateStepIndicators();
}

// ---- Besuch speichern ----

async function besuchSpeichern() {
  const firma = document.getElementById('feldFirma').value.trim();
  if (!firma) {
    showToast('Bitte mindestens die Firma angeben', 'error');
    document.getElementById('feldFirma').focus();
    return;
  }

  const btn = document.getElementById('speichernBtn');
  btn.disabled = true;
  btn.textContent = '\u231B Speichere...';

  const daten = {
    firma: firma,
    ort: document.getElementById('feldOrt').value.trim(),
    kontaktperson: document.getElementById('feldKontakt').value.trim(),
    segment: document.getElementById('feldSegment').value,
    datum: document.getElementById('feldDatum').value,
    tag: document.getElementById('feldWochentag').value,
    besuchstyp: document.getElementById('feldBesuchstyp').value,
    produkte: besuchFormState.produkte.join(', '),
    themen: document.getElementById('feldThemen').value.trim(),
    ergebnis: document.getElementById('feldErgebnis').value.trim(),
    naechsteSchritte: document.getElementById('feldSchritte').value.trim(),
    stellungnahme: document.getElementById('feldStellungnahme').value.trim(),
    projekt: document.getElementById('feldProjekt').value.trim(),
    fotos: JSON.stringify(besuchFormState.fotos.map(f => typeof f === 'string' ? f : f.url))
  };

  try {
    if (besuchFormState.editId) {
      daten.id = besuchFormState.editId;
      await API.besuche.bearbeiten(daten);
      showToast('Besuch erfolgreich aktualisiert!', 'success');
    } else {
      await API.besuche.erstellen(daten);
      showToast('Besuch erfolgreich gespeichert!', 'success');
    }

    // Zurueck zur Besuchsliste
    setTimeout(() => {
      window.location.hash = '#/besuche';
    }, 500);

  } catch (error) {
    showToast('Speichern fehlgeschlagen: ' + error.message, 'error');
    btn.disabled = false;
    btn.textContent = besuchFormState.editId ? '\u{1F4BE} Aenderungen speichern' : '\u2705 Besuch speichern';
  }
}
