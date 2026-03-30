// Notion Besuche-DB CRUD
// GET: Besuche laden (Filter: kw, suche, kunde_id)
// POST: Neuen Besuch erstellen (inkl. Kunde anlegen/verknuepfen)
// PUT: Besuch bearbeiten
// DELETE: Besuch loeschen

const BESUCHE_DB = '5f1f25d10f1d45b4b4e45dabfc2c2c50';
const KUNDEN_DB = '860bf8e12f454281aa30859bebc8e619';
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
  'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
  'Content-Type': 'application/json'
};

// Hilfsfunktion: Kunde in Kunden-DB suchen
async function findKunde(firma) {
  const resp = await fetch(`https://api.notion.com/v1/databases/${KUNDEN_DB}/query`, {
    method: 'POST',
    headers: getHeaders(),
    body: JSON.stringify({
      filter: {
        property: 'Firma',
        title: { equals: firma }
      }
    })
  });
  const data = await resp.json();
  return data.results && data.results.length > 0 ? data.results[0] : null;
}

// Hilfsfunktion: Neuen Kunden anlegen
async function createKunde(besuch) {
  const properties = {
    'Firma': { title: [{ text: { content: besuch.firma || '' } }] },
  };
  if (besuch.ort) properties['Ort'] = { rich_text: [{ text: { content: besuch.ort } }] };
  if (besuch.kontaktperson) properties['Hauptkontakt'] = { rich_text: [{ text: { content: besuch.kontaktperson } }] };
  if (besuch.segment) properties['Segment'] = { select: { name: besuch.segment } };
  properties['Status'] = { select: { name: 'Aktiv' } };
  if (besuch.datum) properties['Letzte Aktivitaet'] = { date: { start: besuch.datum } };

  const resp = await fetch('https://api.notion.com/v1/pages', {
    method: 'POST',
    headers: getHeaders(),
    body: JSON.stringify({
      parent: { database_id: KUNDEN_DB },
      properties
    })
  });
  return await resp.json();
}

// Besuch-Properties bauen
function buildBesuchProperties(data, kundeId) {
  const props = {};

  if (data.firma !== undefined) props['Firma'] = { title: [{ text: { content: data.firma || '' } }] };
  if (data.ort !== undefined) props['Ort'] = { rich_text: [{ text: { content: data.ort || '' } }] };
  if (data.kontaktperson !== undefined) props['Kontaktperson'] = { rich_text: [{ text: { content: data.kontaktperson || '' } }] };
  if (data.segment !== undefined && data.segment) props['Segment'] = { select: { name: data.segment } };
  if (data.tag !== undefined && data.tag) props['Tag'] = { select: { name: data.tag } };
  if (data.kw !== undefined) props['KW'] = { number: parseInt(data.kw) || null };
  if (data.jahr !== undefined) props['Jahr'] = { number: parseInt(data.jahr) || new Date().getFullYear() };
  if (data.typ !== undefined && data.typ) props['Typ'] = { select: { name: data.typ } };
  if (data.besuchstyp !== undefined && data.besuchstyp) props['Besuchstyp'] = { select: { name: data.besuchstyp } };
  if (data.produkte !== undefined) props['Produkte'] = { rich_text: [{ text: { content: data.produkte || '' } }] };
  if (data.themen !== undefined) props['Themen'] = { rich_text: [{ text: { content: (data.themen || '').substring(0, 2000) } }] };
  if (data.ergebnis !== undefined) props['Ergebnis'] = { rich_text: [{ text: { content: (data.ergebnis || '').substring(0, 2000) } }] };
  if (data.naechsteSchritte !== undefined) props['NaechsteSchritte'] = { rich_text: [{ text: { content: (data.naechsteSchritte || '').substring(0, 2000) } }] };
  if (data.todosMarco !== undefined) props['TodosMarco'] = { rich_text: [{ text: { content: (data.todosMarco || '').substring(0, 2000) } }] };
  if (data.stellungnahme !== undefined) props['Stellungnahme'] = { rich_text: [{ text: { content: (data.stellungnahme || '').substring(0, 2000) } }] };
  if (data.projekt !== undefined) props['Projekt'] = { rich_text: [{ text: { content: data.projekt || '' } }] };
  if (data.datum !== undefined && data.datum) props['Datum'] = { date: { start: data.datum } };
  if (data.details !== undefined) props['Details'] = { rich_text: [{ text: { content: (data.details || '').substring(0, 2000) } }] };
  if (data.fotos !== undefined) props['Fotos'] = { rich_text: [{ text: { content: data.fotos || '[]' } }] };

  if (kundeId) {
    props['Kunde'] = { relation: [{ id: kundeId }] };
  }

  return props;
}

// Notion Page zu flachem Objekt konvertieren
function pageToObject(page) {
  const p = page.properties;
  const getText = (prop) => {
    if (!prop) return '';
    if (prop.type === 'title') return prop.title?.map(t => t.plain_text).join('') || '';
    if (prop.type === 'rich_text') return prop.rich_text?.map(t => t.plain_text).join('') || '';
    if (prop.type === 'select') return prop.select?.name || '';
    if (prop.type === 'number') return prop.number;
    if (prop.type === 'date') return prop.date?.start || '';
    if (prop.type === 'relation') return prop.relation?.map(r => r.id) || [];
    return '';
  };

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
    todosMarco: getText(p['TodosMarco']),
    stellungnahme: getText(p['Stellungnahme']),
    projekt: getText(p['Projekt']),
    datum: getText(p['Datum']),
    details: getText(p['Details']),
    fotos: getText(p['Fotos']),
    kundeId: getText(p['Kunde']),
    created: page.created_time,
    updated: page.last_edited_time
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
    // GET: Besuche laden
    if (req.method === 'GET') {
      const kw = url.searchParams.get('kw');
      const suche = url.searchParams.get('suche');
      const kundeId = url.searchParams.get('kunde_id');
      const id = url.searchParams.get('id');

      // Einzelnen Besuch laden
      if (id) {
        const resp = await fetch(`https://api.notion.com/v1/pages/${id}`, {
          headers: getHeaders()
        });
        const page = await resp.json();
        if (page.object === 'error') {
          return new Response(JSON.stringify({ error: page.message }), { status: 404, headers: corsHeaders });
        }
        return new Response(JSON.stringify(pageToObject(page)), { status: 200, headers: corsHeaders });
      }

      // Filter bauen
      const filters = [];

      if (kw) {
        filters.push({
          property: 'KW',
          number: { equals: parseInt(kw) }
        });
      }

      if (kundeId) {
        filters.push({
          property: 'Kunde',
          relation: { contains: kundeId }
        });
      }

      if (suche) {
        // Textsuche ueber mehrere Felder
        filters.push({
          or: [
            { property: 'Firma', title: { contains: suche } },
            { property: 'Produkte', rich_text: { contains: suche } },
            { property: 'Kontaktperson', rich_text: { contains: suche } },
            { property: 'Themen', rich_text: { contains: suche } }
          ]
        });
      }

      const queryBody = {
        sorts: [{ property: 'Datum', direction: 'descending' }],
        page_size: 100
      };

      if (filters.length > 0) {
        queryBody.filter = filters.length === 1 ? filters[0] : { and: filters };
      }

      const resp = await fetch(`https://api.notion.com/v1/databases/${BESUCHE_DB}/query`, {
        method: 'POST',
        headers: getHeaders(),
        body: JSON.stringify(queryBody)
      });

      const data = await resp.json();

      if (data.object === 'error') {
        return new Response(JSON.stringify({ error: data.message }), { status: 400, headers: corsHeaders });
      }

      const besuche = (data.results || []).map(pageToObject);
      return new Response(JSON.stringify(besuche), { status: 200, headers: corsHeaders });
    }

    // POST: Neuen Besuch erstellen
    if (req.method === 'POST') {
      const data = await req.json();

      // Kunde suchen oder anlegen
      let kundeId = null;
      if (data.firma) {
        const existingKunde = await findKunde(data.firma);
        if (existingKunde) {
          kundeId = existingKunde.id;
          // Letzte Aktivitaet aktualisieren
          if (data.datum) {
            await fetch(`https://api.notion.com/v1/pages/${kundeId}`, {
              method: 'PATCH',
              headers: getHeaders(),
              body: JSON.stringify({
                properties: {
                  'Letzte Aktivitaet': { date: { start: data.datum } }
                }
              })
            });
          }
        } else {
          const newKunde = await createKunde(data);
          kundeId = newKunde.id;
        }
      }

      // Wochentag aus Datum berechnen
      if (data.datum && !data.tag) {
        const wochentage = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
        const d = new Date(data.datum);
        data.tag = wochentage[d.getDay()];
      }

      // KW aus Datum berechnen falls nicht gesetzt
      if (data.datum && !data.kw) {
        const d = new Date(data.datum);
        const onejan = new Date(d.getFullYear(), 0, 1);
        const week = Math.ceil(((d - onejan) / 86400000 + onejan.getDay() + 1) / 7);
        data.kw = week;
      }

      if (!data.jahr && data.datum) {
        data.jahr = new Date(data.datum).getFullYear();
      }

      if (!data.typ) data.typ = 'Kundenbesuch';

      const properties = buildBesuchProperties(data, kundeId);

      const resp = await fetch('https://api.notion.com/v1/pages', {
        method: 'POST',
        headers: getHeaders(),
        body: JSON.stringify({
          parent: { database_id: BESUCHE_DB },
          properties
        })
      });

      const result = await resp.json();

      if (result.object === 'error') {
        return new Response(JSON.stringify({ error: result.message }), { status: 400, headers: corsHeaders });
      }

      return new Response(JSON.stringify({ success: true, id: result.id, kundeId }), { status: 201, headers: corsHeaders });
    }

    // PUT: Besuch bearbeiten
    if (req.method === 'PUT') {
      const data = await req.json();
      const { id, ...fields } = data;

      if (!id) {
        return new Response(JSON.stringify({ error: 'Keine Besuchs-ID angegeben' }), { status: 400, headers: corsHeaders });
      }

      // Wochentag neu berechnen
      if (fields.datum && !fields.tag) {
        const wochentage = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
        const d = new Date(fields.datum);
        fields.tag = wochentage[d.getDay()];
      }

      if (fields.datum && !fields.kw) {
        const d = new Date(fields.datum);
        const onejan = new Date(d.getFullYear(), 0, 1);
        fields.kw = Math.ceil(((d - onejan) / 86400000 + onejan.getDay() + 1) / 7);
      }

      // Kunden-Verknuepfung aktualisieren
      let kundeId = null;
      if (fields.firma) {
        const existingKunde = await findKunde(fields.firma);
        if (existingKunde) {
          kundeId = existingKunde.id;
        } else {
          const newKunde = await createKunde(fields);
          kundeId = newKunde.id;
        }
      }

      const properties = buildBesuchProperties(fields, kundeId);

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

    // DELETE: Besuch loeschen (archivieren)
    if (req.method === 'DELETE') {
      const { id } = await req.json();

      if (!id) {
        return new Response(JSON.stringify({ error: 'Keine Besuchs-ID angegeben' }), { status: 400, headers: corsHeaders });
      }

      const resp = await fetch(`https://api.notion.com/v1/pages/${id}`, {
        method: 'PATCH',
        headers: getHeaders(),
        body: JSON.stringify({ archived: true })
      });

      const result = await resp.json();

      if (result.object === 'error') {
        return new Response(JSON.stringify({ error: result.message }), { status: 400, headers: corsHeaders });
      }

      return new Response(JSON.stringify({ success: true }), { status: 200, headers: corsHeaders });
    }

    return new Response(JSON.stringify({ error: 'Methode nicht unterstuetzt' }), { status: 405, headers: corsHeaders });

  } catch (error) {
    console.error('Notion Besuche Fehler:', error);
    return new Response(JSON.stringify({ error: 'Serverfehler', details: error.message }), { status: 500, headers: corsHeaders });
  }
};

export const config = {
  path: "/api/notion-besuche"
};
