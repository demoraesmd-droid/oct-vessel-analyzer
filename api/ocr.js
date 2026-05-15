// /api/ocr.js
// Vercel serverless function. Receives panel crops (pie chart + parameters box + header) as base64 PNGs.
// Calls Anthropic API (Claude vision) to extract structured RNFL clock-hour values, disc parameters,
// patient info, and eye laterality. Returns JSON.

export const config = {
  api: {
    bodyParser: { sizeLimit: '5mb' }
  }
};

const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';
const MODEL = 'claude-sonnet-4-6'; // vision-capable; cheap enough per call

const SYSTEM_PROMPT = `You are reading values from an ophthalmic OCT report screenshot. You will be shown three image regions from the same report:

1. The clock-hour pie chart showing RNFL thickness values in µm, with 12 numbers arranged clockwise around a colored circle.
2. The optic disc parameters box showing labeled values: RA (mm²), DA (mm²), LCDR, VCDR, CV (mm³), @RPH (µm).
3. The top header strip showing patient ID, name, DOB, exam date, and image quality.

Return ONLY a valid JSON object — no markdown, no commentary, no code fences. Structure:

{
  "eye": "OD" or "OS",
  "clockHours": { "1": <int>, "2": <int>, "3": <int>, "4": <int>, "5": <int>, "6": <int>, "7": <int>, "8": <int>, "9": <int>, "10": <int>, "11": <int>, "12": <int> },
  "discParams": { "RA": <float|null>, "DA": <float|null>, "LCDR": <float|null>, "VCDR": <float|null>, "CV": <float|null>, "RPH": <int|null> },
  "patient": { "name": <string|null>, "dob": <string|null>, "examDate": <string|null>, "imageQuality": <int|null> }
}

Rules:
- All clock-hour values are integers 20–250 µm. If a value is unreadable, use null.
- Eye laterality: the pie chart appears the same regardless of eye — derive eye from the header strip ("OD(R)" or "OS(L)" prefix) if present. If header strip absent or unclear, return "OD" as default.
- Patient name: as written. DOB and exam date: as written (preserve format).
- Image quality: integer if shown, else null.
- Disc parameters: null if not visible.
- Output JSON only. No preamble, no markdown, no commentary.`;

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY environment variable not set on server' });
  }

  try {
    const { pieB64, paramsB64, headerB64 } = req.body || {};
    if (!pieB64) {
      return res.status(400).json({ error: 'pieB64 (pie chart image as base64) is required' });
    }

    // Build vision message — order matches the SYSTEM_PROMPT description
    const content = [];
    content.push({ type: 'text', text: 'Image 1 — clock-hour pie chart:' });
    content.push({
      type: 'image',
      source: { type: 'base64', media_type: 'image/png', data: stripDataUrlPrefix(pieB64) }
    });
    if (paramsB64) {
      content.push({ type: 'text', text: 'Image 2 — optic disc parameters box:' });
      content.push({
        type: 'image',
        source: { type: 'base64', media_type: 'image/png', data: stripDataUrlPrefix(paramsB64) }
      });
    }
    if (headerB64) {
      content.push({ type: 'text', text: 'Image 3 — top header strip (patient info):' });
      content.push({
        type: 'image',
        source: { type: 'base64', media_type: 'image/png', data: stripDataUrlPrefix(headerB64) }
      });
    }
    content.push({ type: 'text', text: 'Return the JSON now.' });

    const anthropicRes = await fetch(ANTHROPIC_API_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: 1500,
        system: SYSTEM_PROMPT,
        messages: [{ role: 'user', content }]
      })
    });

    if (!anthropicRes.ok) {
      const errText = await anthropicRes.text();
      return res.status(502).json({ error: `Anthropic API error ${anthropicRes.status}: ${errText.slice(0, 500)}` });
    }

    const data = await anthropicRes.json();
    const textBlock = (data.content || []).find(b => b.type === 'text');
    if (!textBlock) {
      return res.status(502).json({ error: 'No text content returned from Claude' });
    }

    // Strip any markdown code fences just in case the model added them
    let raw = textBlock.text.trim();
    raw = raw.replace(/^```json\s*/, '').replace(/^```\s*/, '').replace(/\s*```$/, '');

    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (e) {
      return res.status(502).json({ error: `Could not parse Claude response as JSON: ${e.message}`, raw: raw.slice(0, 500) });
    }

    return res.status(200).json(parsed);
  } catch (err) {
    return res.status(500).json({ error: `Server error: ${err.message || String(err)}` });
  }
}

function stripDataUrlPrefix(s) {
  if (!s) return s;
  const m = s.match(/^data:image\/[a-z]+;base64,(.+)$/);
  return m ? m[1] : s;
}
