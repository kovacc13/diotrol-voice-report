// Notion Kunden-DB CRUD
// GET: Kunden laden / Einzelkunde mit Besuchen
// POST: Neuen Kunden anlegen
// PUT: Kunden bearbeiten

const KUNDEN_DB = '860bf8e12f454281aa30859bebc8e619';
const BESUCHE_DB = '5f1f25d10f1d45b4b4e45dabfc2c2c50';
const NOTION_VERSION = '2022-06-28';

function getHeaders() {
  return {
    'Authorization': `Bearer ${process.env.NOTION_API_KEY}`,
    'Content-Type': 'application/json',
    'Notion-Version': NOTION_VERSION
  };
}

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'GET, POST, PUT, OPTIONS',
  'Content-Type': 'application/json'
};

function getText(prop) {
  if (!prop) return '';
  if (prop.type === 'title') return prop.title?.map(t => t.plain_text).join('') || '';
  if (prop.type === 'rich_text') return prop.rich_text?.map(t => t.plain_text).join('') || '';
  if (prop.type === 'select') return prop.select?.name || '';
  if (prop.type === 'number') return prop.number;
  if (prop.type === 'date') return prop.date?.start || '';
  if (prop.type === 'relation') return prop.relation?.map(r => r.id) || [];
  if (prop.type === 'url') return prop.url || '';
  if (prop.type === 'email') return prop.email || '';
  if (prop.type === 'phone_number') return prop.phone_number || '';
  return '';
}

function kundeToObject(page) {
  const p = page.properties;
  return {
    id: page.id,
    firma: getText(p['Firma']),
    ort: getText(p['Ort']),
    hauptkontakt: getText(p['Hauptkontakt']),
    segment: getText(p['Segment']),
    telefon: getText(p['Telefon']),
    email: getText(p['Email']),
    webseite: getText(p['Webseite']),
    notizen: getText(p['Notizen']),
    letzteAktivitaet: getText(p['Letzte Aktivitaet']),
    kundeSeit: getText(p['Kunde seit']),
    status: getText(p['Status']),
    besucheIds: getText(p['Besuche']),
    created: page.created_time
  };
}

function besuchToObject(page) {
  const p = page.properties;
  return {
    id: page.id,
    firma: getText(p['Firma']),
    ort: getText(p['Ort']),
    kontaktperson: getText(p['Kontaktperson']),
    segment: getText(p['Segment']),
    tag: getText(p['Tag']),
    kw: getText(p['KW']),
    jahr: getText(p['Jahr']),
    typ: getText(p['Typ']),
    besuchstyp: getText(p['Besuchstyp']),
    produkte: getText(p['Produkte']),
    themen: getText(p['Themen']),
    ergebnis: getText(p['Ergebnis']),
    naechsteSchritte: getText(p['NaechsteSchritte']),
    stellungnahme: getText(p['Stellungnahme']),
    projekt: getText(p['Projekt']),
    datum: getText(p['Datum']),
    details: getText(p['Details']),
    fotos: getText(p['Fotos']),
    created: page.created_time
  };
}

export default async (req, context) => {
  if (req.method === 'OPTIONS') {
    return new Response('', { status: 204, headers: corsHeaders });
  }

  if (!process.env.NOTION_API_KEY) {
    return new Response(JSON.stringify({ error: 'NOTION_API_KEY nicht konfiguriert' }), { status: 503, headers: corsHeaders });
  }

  const url = new URL(req.url);

  try {
    // GET: Kunden laden
    if (req.method === 'GET') {
      const id = url.searchParams.get('id');
      const suche = url.searchParams.get('suche');

      // Einzelkunde mit allen Besuchen
      if (id) {
        // Kunden-Daten laden
        const kundeResp = await fetch(`https://api.notion.com/v1/pages/${id}`, {
          headers: getHeaders()
        });
        const kundePage = await kundeResp.json();

        if (kundePage.object === 'error') {
          return new Response(JSON.stringify({ error: kundePage.message }), { status: 404, headers: corsHeaders });
        }

        const kunde = kundeToObject(kundePage);

        // Besuche dieses Kunden laden
        const besucheResp = await fetch(`https://api.notion.com/v1/databases/${BESUCHE_DB}/query`, {
          method: 'POST',
          headers: getHeaders(),
          body: JSON.stringify({
            filter: {
              property: 'Kunde',
              relation: { contains: id }
            },
            sorts: [{ property: 'Datum', direction: 'descending' }]
          })
        });

        const besucheData = await besucheResp.json();
        kunde.besuche = (besucheData.results || []).map(besuchToObject);
        kunde.besucheAnzahl = kunde.besuche.length;

        return new Response(JSON.stringify(kunde), { status: 200, headers: corsHeaders });
      }

      // Alle Kunden laden
      const queryBody = {
        sorts: [{ property: 'Firma', direction: 'ascending' }],
        page_size: 100
      };

      if (suche) {
        queryBody.filter = {
          or: [
            { property: 'Firma', title: { contains: suche } },
            { property: 'Ort', rich_text: { contains: suche } }
          ]
        };
      }

      const resp = await fetch(`https://api.notion.com/v1/databases/${KUNDEN_DB}/query`, {
        method: 'POST',
        headers: getHeaders(),
        body: JSON.stringify(queryBody)
      });

      const data = await resp.json();

      if (data.object === 'error') {
        return new Response(JSON.stringify({ error: data.message }), { status: 400, headers: corsHeaders });
      }

      const kunden = (data.results || []).map(page => {
        const k = kundeToObject(page);
        k.besucheAnzahl = Array.isArray(k.besucheIds) ? k.besucheIds.length : 0;
        return k;
      });

      // Wenn Besuche-Relationen leer sind, Besuchsanzahl per Firmenname zaehlen
      const hatLeereRelationen = kunden.some(k => k.besucheAnzahl === 0);
      if (hatLeereRelationen) {
        try {
          // Alle Besuche laden um per Firmenname zu matchen
          const besucheResp = await fetch(`https://api.notion.com/v1/databases/${BESUCHE_DB}/query`, {
            method: 'POST',
            headers: getHeaders(),
            body: JSON.stringify({
              page_size: 100
            })
          });
          const besucheData = await besucheResp.json();
          const alleBesuche = (besucheData.results || []).map(besuchToObject);

          // Besuchsanzahl per Firmenname zaehlen
          const besuchsMap = {};
          alleBesuche.forEach(b => {
            if (b.firma) {
              const key = b.firma.trim().toLowerCase();
              besuchsMap[key] = (besuchsMap[key] || 0) + 1;
            }
          });

          // Besuchsanzahl zuweisen (hoechsten Wert nehmen)
          kunden.forEach(k => {
            const key = (k.firma || '').trim().toLowerCase();
            const matchedCount = besuchsMap[key] || 0;
            k.besucheAnzahl = Math.max(k.besucheAnzahl, matchedCount);
          });
        } catch (e) {
          console.error('Fehler beim Zaehlen der Besuche per Firmenname:', e);
        }
      }

      return new Response(JSON.stringify(kunden), { status: 200, headers: corsHeaders });
    }

    // POST: Neuen Kunden anlegen
    if (req.method === 'POST') {
      const data = await req.json();

      const properties = {
        'Firma': { title: [{ text: { content: data.firma || '' } }] }
      };

      if (data.ort) properties['Ort'] = { rich_text: [{ text: { content: data.ort } }] };
      if (data.hauptkontakt) properties['Hauptkontakt'] = { rich_text: [{ text: { content: data.hauptkontakt } }] };
      if (data.segment) properties['Segment'] = { select: { name: data.segment } };
      if (data.telefon) properties['Telefon'] = { phone_number: data.telefon };
      if (data.email) properties['Email'] = { email: data.email };
      if (data.webseite) properties['Webseite'] = { url: data.webseite };
      if (data.notizen) properties['Notizen'] = { rich_text: [{ text: { content: data.notizen } }] };
      if (data.status) properties['Status'] = { select: { name: data.status } };
      else properties['Status'] = { select: { name: 'Aktiv' } };

      const resp = await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: getHeaders(),
        body: JSON.stringify({
          parent: { database_id: KUNDEN_DB },
          properties
        })
      });

      const result = await resp.json();

      if (result.object === 'error') {
        return new Response(JSON.stringify({ error: result.message }), { status: 400, headers: corsHeaders });
      }

      return new Response(JSON.stringify({ success: true, id: result.id }), { status: 201, headers: corsHeaders });
    }

    // PUT: Kunden bearbeiten
    if (req.method === 'PUT') {
      const data = await req.json();
      const { id, ...fields } = data;

      if (!id) {
        return new Response(JSON.stringify({ error: 'Keine Kunden-ID angegeben' }), { status: 400, headers: corsHeaders });
      }

      const properties = {};
      if (fields.firma !== undefined) properties['Firma'] = { title: [{ text: { content: fields.firma } }] };
      if (fields.ort !== undefined) properties['Ort'] = { rich_text: [{ text: { content: fields.ort } }] };
      if (fields.hauptkontakt !== undefined) properties['Hauptkontakt'] = { rich_text: [{ text: { content: fields.hauptkontakt } }] };
      if (fields.segment !== undefined && fields.segment) properties['Segment'] = { select: { name: fields.segment } };
      if (fields.telefon !== undefined) properties['Telefon'] = { phone_number: fields.telefon || null };
      if (fields.email !== undefined) properties['Email'] = { email: fields.email || null };
      if (fields.webseite !== undefined) properties['Webseite'] = { url: fields.webseite || null };
      if (fields.notizen !== undefined) properties['Notizen'] = { rich_text: [{ text: { content: fields.notizen } }] };
      if (fields.status !== undefined && fields.status) properties['Status'] = { select: { name: fields.status } };

      const resp = await fetch(`https://api.notion.com/v1/pages/${id}`, {
        method: 'PATCH',
        headers: getHeaders(),
        body: JSON.stringify({ properties })
      });

      const result = await resp.json();

      if (result.object === 'error') {
        return new Response(JSON.stringify({ error: result.message }), { status: 400, headers: corsHeaders });
      }

      return new Response(JSON.stringify({ success: true, id: result.id }), { status: 200, headers: corsHeaders });
    }

    return new Response(JSON.stringify({ error: 'Methode nicht unterstuetzt' }), { status: 405, headers: corsHeaders });

  } catch (error) {
    console.error('Notion Kunden Fehler:', error);
    return new Response(JSON.stringify({ error: 'Serverfehler', details: error.message }), { status: 500, headers: corsHeaders });
  }
};

export const config = {
  path: "/api/notion-kunden"
};
