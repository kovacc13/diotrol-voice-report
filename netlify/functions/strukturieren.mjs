// KI-Strukturierung via Claude API
// POST: Text -> strukturiertes JSON

export default async (req, context) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Content-Type': 'application/json'
  };

  if (req.method === 'OPTIONS') {
    return new Response('', { status: 204, headers });
  }

  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'Nur POST erlaubt' }), { status: 405, headers });
  }

  try {
    const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

    if (!ANTHROPIC_API_KEY) {
      return new Response(JSON.stringify({ error: 'ANTHROPIC_API_KEY nicht konfiguriert' }), { status: 503, headers });
    }

    const body = await req.json();
    const { text } = body;

    if (!text || text.trim().length === 0) {
      return new Response(JSON.stringify({ error: 'Kein Text zum Analysieren' }), { status: 400, headers });
    }

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 2000,
        messages: [{
          role: 'user',
          content: `Analysiere diesen diktierten Besuchsbericht eines Aussendienstmitarbeiters der Diotrol AG (Holzbeschichtungen/Holzschutz, Schweiz) und extrahiere die Informationen als JSON.

Der Text:
"""
${text}
"""

Antworte NUR mit einem JSON-Objekt (kein Markdown, keine Erklaerung):
{
  "firma": "Firmenname",
  "ort": "Ort/Stadt",
  "kontaktperson": "Name der Kontaktperson",
  "segment": "eines von: Schreiner, Maler, Zimmerei, Fensterbau, Fassadenbau, Architekt, Generalunternehmer, Handel, Sonstiges",
  "besuchstyp": "eines von: Erstbesuch, Folgebesuch, Kaltbesuch, Beratung, Reklamation, Schulung, Messe, Telefonat",
  "produkte": "Produkt1, Produkt2, Produkt3",
  "themen": "\u2022 Thema 1\\n\u2022 Thema 2\\n\u2022 Thema 3",
  "ergebnis": "Zusammenfassung des Ergebnisses",
  "naechsteSchritte": "\u2022 Schritt 1\\n\u2022 Schritt 2",
  "stellungnahme": "optional, leer lassen wenn nicht erwaehnt",
  "projekt": "optional, Projektname wenn erwaehnt"
}`
        }]
      })
    });

    if (!response.ok) {
      const errText = await response.text();
      console.error('Claude API Fehler:', errText);
      return new Response(JSON.stringify({ error: 'Claude API Fehler', details: errText }), { status: 502, headers });
    }

    const result = await response.json();
    let rawText = result.content[0].text;

    // Entferne moegliche Markdown-Bloecke (Claude gibt manchmal ```json ... ``` zurueck)
    rawText = rawText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();

    // JSON parsen
    let structured;
    try {
      structured = JSON.parse(rawText);
    } catch {
      // Letzter Versuch: erstes { bis letztes }
      const start = rawText.indexOf('{');
      const end = rawText.lastIndexOf('}');
      if (start !== -1 && end !== -1) {
        structured = JSON.parse(rawText.substring(start, end + 1));
      } else {
        throw new Error('Konnte JSON nicht aus Antwort extrahieren');
      }
    }

    return new Response(JSON.stringify(structured), { status: 200, headers });

  } catch (error) {
    console.error('Strukturierung Fehler:', error);
    return new Response(JSON.stringify({ error: 'Fehler bei KI-Analyse', details: error.message }), { status: 500, headers });
  }
};

export const config = {
  path: "/api/strukturieren"
};
