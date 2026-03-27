// ========================================
// API-Schicht fuer D.P. Report
// ========================================

const API = {
  // Basis-URL fuer API-Aufrufe
  base: '/api',

  // Generischer Fetch-Wrapper
  async request(endpoint, options = {}) {
    const url = options.params
      ? `${this.base}/${endpoint}?${new URLSearchParams(options.params)}`
      : `${this.base}/${endpoint}`;

    const config = {
      method: options.method || 'GET',
      headers: { 'Content-Type': 'application/json' },
    };

    if (options.body) {
      config.body = JSON.stringify(options.body);
    }

    try {
      const resp = await fetch(url, config);
      const data = await resp.json();

      if (!resp.ok) {
        throw new Error(data.error || `Fehler ${resp.status}`);
      }

      return data;
    } catch (error) {
      console.error(`API Fehler (${endpoint}):`, error);
      throw error;
    }
  },

  // === Besuche ===
  besuche: {
    async laden(filter = {}) {
      return API.request('notion-besuche', { params: filter });
    },
    async einzeln(id) {
      return API.request('notion-besuche', { params: { id } });
    },
    async erstellen(daten) {
      return API.request('notion-besuche', { method: 'POST', body: daten });
    },
    async bearbeiten(daten) {
      return API.request('notion-besuche', { method: 'PUT', body: daten });
    },
    async loeschen(id) {
      return API.request('notion-besuche', { method: 'DELETE', body: { id } });
    }
  },

  // === Kunden ===
  kunden: {
    async laden(filter = {}) {
      return API.request('notion-kunden', { params: filter });
    },
    async einzeln(id) {
      return API.request('notion-kunden', { params: { id } });
    },
    async erstellen(daten) {
      return API.request('notion-kunden', { method: 'POST', body: daten });
    },
    async bearbeiten(daten) {
      return API.request('notion-kunden', { method: 'PUT', body: daten });
    }
  },

  // === Wochenbericht ===
  wochenbericht: {
    async laden(kw, jahr) {
      const params = {};
      if (kw) params.kw = kw;
      if (jahr) params.jahr = jahr;
      return API.request('notion-wochenbericht', { params });
    }
  },

  // === Strukturierung (Claude KI) ===
  async strukturieren(text) {
    return API.request('strukturieren', { method: 'POST', body: { text } });
  },

  // === Cloudinary Upload ===
  async fotoUpload(base64Image) {
    return API.request('cloudinary-upload', { method: 'POST', body: { image: base64Image } });
  }
};
