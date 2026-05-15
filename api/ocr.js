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

const SYSTEM_PROMPT = `You are an OCR extraction tool. Your ONLY job is to output valid JSON that exactly matches the schema. Never write explanations, markdown, or commentary.

You will be shown three image regions from an ophthalmic OCT report:
1. The clock-hour pie chart — 12 RNFL thickness numbers (µm) arranged clockwise around a colored circle.
2. The optic disc parameters box — labeled values: RA (mm²), DA (mm²), LCDR, VCDR, CV (mm³), @RPH (µm).
3. The top header strip — patient ID, name, DOB, exam date, image quality, and the eye marker "OD(R)" or "OS(L)".

Output schema (exactly this structure):
{
  "eye": "OD" or "OS",
  "clockHours": { "1": <int>, "2": <int>, "3": <int>, "4": <int>, "5": <int>, "6": <int>, "7": <int>, "8": <int>, "9": <int>, "10": <int>, "11": <int>, "12": <int> },
  "discParams": { "RA": <float|null>, "DA": <float|null>, "LCDR": <float|null>, "VCDR": <float|null>, "CV": <float|null>, "RPH": <int|null> },
  "patient": { "name": <string|null>, "dob": <string|null>, "examDate": <string|null>, "imageQuality": <int|null> }
}

Rules:
- Clock-hour values are integers 20–250 µm. If unreadable, use null.
- Eye: read from the header strip ("OD" or "OS"). Default "OD" if absent.
- Disc parameters: null if missing.
- Image quality: integer (e.g. 61) or null.

ABSOLUTE RULE: Output ONLY the JSON object. Start your response with "{" and end with "}". No prose, no markdown fences, no explanations. Any characters outside the JSON will cause the system to fail.`;

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

    // Strip any markdown code fences just in case
    let raw = textBlock.text.trim();
    raw = raw.replace(/^```json\s*/, '').replace(/^```\s*/, '').replace(/\s*```$/, '');

    // Truncate after the matching closing brace to discard any trailing prose
    raw = extractJSONObject(raw);

    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (e) {
      // Last-resort salvage: parse markdown-formatted output by extracting key:value patterns
      const salvaged = salvageJSON(textBlock.text);
      if (salvaged) {
        return res.status(200).json(salvaged);
      }
      return res.status(502).json({
        error: `Could not parse Claude response as JSON: ${e.message}`,
        raw: raw.slice(0, 500)
      });
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

// Extract just the first balanced JSON object from a string — discard anything after closing brace
function extractJSONObject(s) {
  let depth = 0, inStr = false, esc = false;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (esc) { esc = false; continue; }
    if (c === '\\' && inStr) { esc = true; continue; }
    if (c === '"') { inStr = !inStr; continue; }
    if (inStr) continue;
    if (c === '{') depth++;
    else if (c === '}') {
      depth--;
      if (depth === 0) return s.slice(0, i + 1);
    }
  }
  return s;
}

// Last-ditch salvage: parse markdown-formatted output by extracting key:value patterns
function salvageJSON(text) {
  try {
    const result = { eye: 'OD', clockHours: {}, discParams: {}, patient: {} };

    // Clock-hour values: look for patterns like "12: 138" or "- 12: 138" or "12. 138"
    for (let c = 1; c <= 12; c++) {
      const re = new RegExp(`(?:^|\\s|-|\\*)\\s*${c}\\s*[:.\\)]\\s*(\\d{2,3})`, 'm');
      const m = text.match(re);
      if (m) {
        const v = parseInt(m[1], 10);
        if (v >= 20 && v <= 250) result.clockHours[c] = v;
      }
    }
    if (Object.keys(result.clockHours).length < 8) return null;  // not enough data salvaged

    // Disc params
    const findNum = (label, re) => {
      const m = text.match(re);
      return m ? parseFloat(m[1]) : null;
    };
    result.discParams.RA = findNum('RA', /RA[^:\d]{0,10}([\d.]+)/i);
    result.discParams.DA = findNum('DA', /DA[^:\d]{0,10}([\d.]+)/i);
    result.discParams.LCDR = findNum('LCDR', /LCDR[^:\d]{0,10}([\d.]+)/i);
    result.discParams.VCDR = findNum('VCDR', /VCDR[^:\d]{0,10}([\d.]+)/i);
    result.discParams.CV = findNum('CV', /CV[^:\d]{0,10}([\d.]+)/i);
    result.discParams.RPH = findNum('RPH', /RPH[^:\d]{0,10}([\d.]+)/i);

    // Eye
    if (/\bOS\b|\bOS\(L\)/.test(text)) result.eye = 'OS';
    else if (/\bOD\b|\bOD\(R\)/.test(text)) result.eye = 'OD';

    return result;
  } catch (e) {
    return null;
  }
}
