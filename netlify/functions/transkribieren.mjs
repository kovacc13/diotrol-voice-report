// Transkribierung via OpenAI Whisper API
// POST: Audio-Blob -> Text
// Verwendet OpenAI Whisper (beste Qualitaet fuer Deutsch)

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
    const OPENAI_API_KEY = process.env.OPENAI_API_KEY;

    if (!OPENAI_API_KEY) {
      return new Response(JSON.stringify({
        error: 'OPENAI_API_KEY nicht konfiguriert',
        fallback: 'browser_speech_api'
      }), { status: 503, headers });
    }

    // Audio-Daten aus Request lesen
    const formData = await req.formData();
    const audioFile = formData.get('file');

    if (!audioFile) {
      return new Response(JSON.stringify({ error: 'Keine Audio-Datei gesendet' }), { status: 400, headers });
    }

    // An OpenAI Whisper API senden
    const whisperForm = new FormData();
    whisperForm.append('file', audioFile, 'audio.webm');
    whisperForm.append('model', 'whisper-1');
    whisperForm.append('language', 'de');
    whisperForm.append('response_format', 'json');

    const whisperResponse = await fetch('https://api.openai.com/v1/audio/transcriptions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`
      },
      body: whisperForm
    });

    if (!whisperResponse.ok) {
      const errText = await whisperResponse.text();
      console.error('OpenAI Whisper API Fehler:', errText);
      return new Response(JSON.stringify({
        error: 'Whisper API Fehler',
        details: errText,
        fallback: 'browser_speech_api'
      }), { status: 502, headers });
    }

    const result = await whisperResponse.json();
    return new Response(JSON.stringify({ text: result.text || '' }), { status: 200, headers });

  } catch (error) {
    console.error('Transkribierung Fehler:', error);
    return new Response(JSON.stringify({
      error: 'Serverfehler bei Transkribierung',
      fallback: 'browser_speech_api'
    }), { status: 500, headers });
  }
};

export const config = {
  path: "/api/transkribieren"
};
