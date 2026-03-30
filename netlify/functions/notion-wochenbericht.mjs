// Wochenbericht: Alle Besuche einer KW + Statistiken
// GET ?kw=12&jahr=2026

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
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
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
  return '';
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
    todosMarco: getText(p['TodosMarco']),
    stellungnahme: getText(p['Stellungnahme']),
    projekt: getText(p['Projekt']),
    datum: getText(p['Datum']),
    details: getText(p['Details']),
    fotos: getText(p['Fotos'])
  };
}

export default async (req, context) => {
  if (req.method === 'OPTIONS') {
    return new Response('', { status: 204, headers: corsHeaders });
  }

  if (req.method !== 'GET') {
    return new Response(JSON.stringify({ error: 'Nur GET erlaubt' }), { status: 405, headers: corsHeaders });
  }

  if (!process.env.NOTION_API_KEY) {
    return new Response(JSON.stringify({ error: 'NOTION_API_KEY nicht konfiguriert' }), { status: 503, headers: corsHeaders });
  }

  const url = new URL(req.url);
  const kw = parseInt(url.searchParams.get('kw')) || getCurrentKW();
  const jahr = parseInt(url.searchParams.get('jahr')) || new Date().getFullYear();

  try {
    // Besuche dieser KW laden
    const filters = {
      and: [
        { property: 'KW', number: { equals: kw } },
        { property: 'Jahr', number: { equals: jahr } }
      ]
    };

    const resp = await fetch(`https://api.notion.com/v1/databases/${BESUCHE_DB}/query`, {
      method: 'POST',
      headers: getHeaders(),
      body: JSON.stringify({
        filter: filters,
        sorts: [{ property: 'Datum', direction: 'ascending' }],
        page_size: 100
      })
    });

    const data = await resp.json();

    if (data.object === 'error') {
      return new Response(JSON.stringify({ error: data.message }), { status: 400, headers: corsHeaders });
    }

    const besuche = (data.results || []).map(besuchToObject);

    // Statistiken berechnen
    const firmen = new Set(besuche.map(b => b.firma));
    const alleProdukte = new Set();
    besuche.forEach(b => {
      if (b.produkte) {
        b.produkte.split(',').forEach(p => {
          const trimmed = p.trim();
          if (trimmed) alleProdukte.add(trimmed);
        });
      }
    });

    // Erstbesuche = Neukunden (approximiert)
    const neukunden = besuche.filter(b => b.besuchstyp === 'Erstbesuch').length;
    const bestandskunden = besuche.length - neukunden;

    // Segmentverteilung berechnen (sortiert nach Anzahl absteigend)
    const segmentCount = {};
    besuche.forEach(b => {
      const seg = b.segment || 'Sonstiges';
      segmentCount[seg] = (segmentCount[seg] || 0) + 1;
    });
    const segmentRanking = Object.entries(segmentCount)
      .sort((a, b) => b[1] - a[1])
      .map(([segment, count]) => ({ segment, count, prozent: Math.round(count / besuche.length * 100) }));

    const stats = {
      totalBesuche: besuche.length,
      neukunden,
      bestandskunden,
      fokusProdukte: Array.from(alleProdukte).slice(0, 5).join(', '),
      firmenBesucht: firmen.size,
      segmentRanking
    };

    // Verfuegbare KWs ermitteln (fuer Dropdown)
    const kwResp = await fetch(`https://api.notion.com/v1/databases/${BESUCHE_DB}/query`, {
      method: 'POST',
      headers: getHeaders(),
      body: JSON.stringify({
        filter: { property: 'Jahr', number: { equals: jahr } },
        sorts: [{ property: 'KW', direction: 'descending' }],
        page_size: 100
      })
    });
    const kwData = await kwResp.json();
    const verfuegbareKWs = [...new Set((kwData.results || []).map(p => {
      const kwProp = p.properties['KW'];
      return kwProp?.number;
    }).filter(Boolean))].sort((a, b) => b - a);

    return new Response(JSON.stringify({
      kw,
      jahr,
      besuche,
      stats,
      verfuegbareKWs
    }), { status: 200, headers: corsHeaders });

  } catch (error) {
    console.error('Wochenbericht Fehler:', error);
    return new Response(JSON.stringify({ error: 'Serverfehler', details: error.message }), { status: 500, headers: corsHeaders });
  }
};

function getCurrentKW() {
  const now = new Date();
  const onejan = new Date(now.getFullYear(), 0, 1);
  return Math.ceil(((now - onejan) / 86400000 + onejan.getDay() + 1) / 7);
}

export const config = {
  path: "/api/notion-wochenbericht"
};
