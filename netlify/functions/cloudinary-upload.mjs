// Cloudinary Bild-Upload
// POST: Base64 oder Multipart -> Cloudinary URL

export default async (req, context) => {
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Content-Type': 'application/json'
  };

  if (req.method === 'OPTIONS') {
    return new Response('', { status: 204, headers: corsHeaders });
  }

  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'Nur POST erlaubt' }), { status: 405, headers: corsHeaders });
  }

  const { CLOUDINARY_CLOUD_NAME, CLOUDINARY_API_KEY, CLOUDINARY_API_SECRET } = process.env;

  if (!CLOUDINARY_CLOUD_NAME || !CLOUDINARY_API_KEY || !CLOUDINARY_API_SECRET) {
    return new Response(JSON.stringify({ error: 'Cloudinary nicht konfiguriert' }), { status: 503, headers: corsHeaders });
  }

  try {
    const body = await req.json();
    const { image } = body; // Base64 String mit data:image/... Prefix

    if (!image) {
      return new Response(JSON.stringify({ error: 'Kein Bild gesendet' }), { status: 400, headers: corsHeaders });
    }

    // Timestamp und Signatur erstellen
    const timestamp = Math.round(Date.now() / 1000);
    const params = `folder=diotrol-besuche&timestamp=${timestamp}${CLOUDINARY_API_SECRET}`;

    // SHA-1 Signatur berechnen
    const encoder = new TextEncoder();
    const data = encoder.encode(params);
    const hashBuffer = await crypto.subtle.digest('SHA-1', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    const signature = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');

    // Upload zu Cloudinary
    const formData = new URLSearchParams();
    formData.append('file', image);
    formData.append('api_key', CLOUDINARY_API_KEY);
    formData.append('timestamp', timestamp.toString());
    formData.append('signature', signature);
    formData.append('folder', 'diotrol-besuche');

    const uploadResp = await fetch(
      `https://api.cloudinary.com/v1_1/${CLOUDINARY_CLOUD_NAME}/image/upload`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: formData.toString()
      }
    );

    const result = await uploadResp.json();

    if (result.error) {
      return new Response(JSON.stringify({ error: result.error.message }), { status: 400, headers: corsHeaders });
    }

    return new Response(JSON.stringify({
      url: result.secure_url,
      public_id: result.public_id,
      width: result.width,
      height: result.height,
      thumbnail: result.secure_url.replace('/upload/', '/upload/c_thumb,w_200,h_150/')
    }), { status: 200, headers: corsHeaders });

  } catch (error) {
    console.error('Cloudinary Upload Fehler:', error);
    return new Response(JSON.stringify({ error: 'Upload fehlgeschlagen', details: error.message }), { status: 500, headers: corsHeaders });
  }
};

export const config = {
  path: "/api/cloudinary-upload"
};
