// /api/docx.js
// Vercel serverless function. Receives the structured analysis result + annotated images as base64,
// builds a DOCX clinical report using the `docx` package, returns it as binary.

import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, PageOrientation, HeadingLevel, BorderStyle, WidthType, ShadingType, VerticalAlign
} from 'docx';

export const config = {
  api: {
    bodyParser: { sizeLimit: '15mb' }
  }
};

const SCAN_RADIUS_UM = 1730;
const ARC_PER_QUADRANT = (2 * Math.PI * SCAN_RADIUS_UM) / 4;
const Q_COLORS = { S: 'FFE4B0', N: 'C7DBF5', I: 'F5C7DB', T: 'C7E5C7' };
const Q_NAMES = { S: 'Superior', N: 'Nasal', I: 'Inferior', T: 'Temporal' };

const border = { style: BorderStyle.SINGLE, size: 4, color: '888888' };
const borders = { top: border, bottom: border, left: border, right: border };

const cell = (text, width, opts = {}) => new TableCell({
  borders,
  width: { size: width, type: WidthType.DXA },
  margins: { top: 80, bottom: 80, left: 120, right: 120 },
  shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
  verticalAlign: VerticalAlign.CENTER,
  children: [new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({
      text: String(text),
      bold: opts.bold || false,
      size: opts.size || 18,
      color: opts.color || '000000',
      italics: opts.italics || false
    })]
  })]
});
const p = (text, opts = {}) => new Paragraph({ children: [new TextRun({ text, ...(opts.run || {}) })] });
const h1 = text => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text, bold: true })] });
const h2 = text => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text, bold: true })] });

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const r = req.body;
    if (!r || !r.eye || !r.rnfl || !r.vessels || !r.aggregation) {
      return res.status(400).json({ error: 'Missing required fields in request body' });
    }

    const rfBuf = r.annRfPngB64 ? Buffer.from(stripDataUrl(r.annRfPngB64), 'base64') : null;
    const bsBuf = r.annBsPngB64 ? Buffer.from(stripDataUrl(r.annBsPngB64), 'base64') : null;

    const children = buildReportChildren(r, rfBuf, bsBuf);

    const doc = new Document({
      styles: {
        default: { document: { run: { font: 'Arial', size: 20 } } },
        paragraphStyles: [
          {
            id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { size: 30, bold: true, font: 'Arial', color: '1F3864' },
            paragraph: { spacing: { before: 240, after: 180 }, outlineLevel: 0 }
          },
          {
            id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { size: 24, bold: true, font: 'Arial', color: '2E5497' },
            paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 }
          }
        ]
      },
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: 15840, orientation: PageOrientation.LANDSCAPE },
            margin: { top: 720, right: 720, bottom: 720, left: 720 }
          }
        },
        children
      }]
    });

    const buf = await Packer.toBuffer(doc);
    const fname = `OCT_${r.eye}_${(r.header?.patientName || 'patient').replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30)}_${(r.header?.examDate || '').replace(/\//g, '-')}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
    res.setHeader('Content-Length', buf.length);
    return res.status(200).send(buf);
  } catch (err) {
    return res.status(500).json({ error: `DOCX generation failed: ${err.message || String(err)}` });
  }
}

function stripDataUrl(s) {
  if (!s) return s;
  const m = s.match(/^data:image\/[a-z]+;base64,(.+)$/);
  return m ? m[1] : s;
}

function buildReportChildren(r, rfBuf, bsBuf) {
  const children = [];

  // Title and header line
  children.push(h1(`OCT Peripapillary Vessel Analysis — ${r.eye === 'OD' ? 'Right' : 'Left'} Eye (${r.eye})`));
  const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const patientLine = [
    r.header?.patientName ? `Patient: ${r.header.patientName}` : 'Patient: —',
    r.header?.dob ? `DOB: ${r.header.dob}` : '',
    r.header?.examDate ? `Exam: ${r.header.examDate}` : '',
    `Generated: ${today}`
  ].filter(Boolean).join('   |   ');
  children.push(p(patientLine, { run: { italics: true, color: '555555', size: 18 } }));
  children.push(p(`Analysis: OCT peripapillary scan (1730 µm radius). Image quality: ${r.header?.iq || '—'}.`,
    { run: { italics: true, color: '555555', size: 18 } }));

  // Caveat box
  if (r.mode === 'redfree') {
    children.push(new Paragraph({
      border: { left: { style: BorderStyle.SINGLE, size: 24, color: 'CC4400', space: 12 } },
      spacing: { before: 120, after: 120 },
      indent: { left: 200, right: 200 },
      children: [
        new TextRun({ text: '⚠ Methodology caveat — this scan only ', bold: true, color: 'AA3300', size: 20 }),
        new TextRun({ text: 'B-scan vessel-shadow detection didn\'t meet the ≥15% contrast threshold. Vessel positions and diameters use the red-free image as a fallback. Diameter estimates from red-free systematically run larger than B-scan FWHM, so absolute %RNFL values are not directly comparable to B-scan-derived reports.', size: 18 })
      ]
    }));
  } else if (r.mode === 'hybrid') {
    const supp = r.vessels.filter(v => v.source === 'redfree-supplement');
    const suppQuads = [...new Set(supp.map(v => v.quadrant))].sort().join(', ');
    children.push(new Paragraph({
      border: { left: { style: BorderStyle.SINGLE, size: 24, color: 'CC4400', space: 12 } },
      spacing: { before: 120, after: 120 },
      indent: { left: 200, right: 200 },
      children: [
        new TextRun({ text: '⚠ Mixed-method detection — ', bold: true, color: 'AA3300', size: 20 }),
        new TextRun({ text: `${supp.length} vessel(s) were added from red-free in regions where B-scan shadow contrast fell below 15% — typically due to local RNFL thinning (sectors: ${suppQuads}). These vessels are tagged "redfree-suppl." in the table; their diameters use red-free FWHM (which runs systematically larger than B-scan FWHM), while B-scan ground-truth vessels keep their B-scan FWHM measurements. The two methods are not directly comparable in absolute terms, but combining them captures vessels that pure B-scan would miss in atrophic sectors.`, size: 18 })
      ]
    }));
  }

  // Disc parameters
  const params = r.params || {};
  children.push(h2('Optic Disc Parameters'));
  children.push(p(
    `Rim Area = ${params.RA?.toFixed?.(2) ?? '—'} mm²   |   Disc Area = ${params.DA?.toFixed?.(2) ?? '—'} mm²   |   LCDR = ${params.LCDR ?? '—'}   |   VCDR = ${params.VCDR ?? '—'}   |   Cup Volume = ${params.CV ?? '—'} mm³   |   @RPH = ${params.RPH ?? '—'} µm`,
    { run: { size: 18 } }
  ));

  // Clock-hour table
  children.push(h2('RNFL Thickness — Clock-Hour Map'));
  const cwidth = 683;
  const labelW = 9300 - 12 * cwidth;
  const cwArr = [labelW, ...new Array(12).fill(cwidth)];
  const order = [12, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const quadOf = { 12: 'S', 1: 'S', 2: 'N', 3: 'N', 4: 'N', 5: 'I', 6: 'I', 7: 'I', 8: 'T', 9: 'T', 10: 'T', 11: 'S' };
  const clockTable = new Table({
    width: { size: 9300, type: WidthType.DXA },
    columnWidths: cwArr,
    rows: [
      new TableRow({
        children: [
          cell('Clock hour', cwArr[0], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
          ...order.map((c, i) => cell(c, cwArr[i + 1], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }))
        ]
      }),
      new TableRow({
        children: [
          cell('Quadrant', cwArr[0], { fill: 'EEEEEE', bold: true, align: AlignmentType.CENTER, size: 16 }),
          ...order.map((c, i) => cell(quadOf[c], cwArr[i + 1], { fill: Q_COLORS[quadOf[c]], align: AlignmentType.CENTER, size: 16 }))
        ]
      }),
      new TableRow({
        children: [
          cell('RNFL (µm)', cwArr[0], { bold: true, align: AlignmentType.CENTER, fill: 'EEEEEE' }),
          ...order.map((c, i) => {
            const v = r.rnfl.clock[c];
            let fill = 'B0E5B0';
            if (v < 60) fill = 'F5B0B0';
            else if (v < 75) fill = 'F5E5B0';
            return cell(v, cwArr[i + 1], { align: AlignmentType.CENTER, fill, bold: v < 75 });
          })
        ]
      })
    ]
  });
  children.push(clockTable);
  children.push(p(
    `Quadrant averages: S = ${r.rnfl.S.toFixed(0)} µm  |  N = ${r.rnfl.N.toFixed(0)} µm  |  I = ${r.rnfl.I.toFixed(0)} µm  |  T = ${r.rnfl.T.toFixed(0)} µm  |  Global = ${r.rnfl.G.toFixed(0)} µm.`,
    { run: { size: 18, bold: true } }
  ));

  // Vessel table
  children.push(h2('Vessel Crossings'));
  const usableCount = r.vessels.filter(v => v.dUm !== null).length;
  const excludedCount = r.vessels.filter(v => v.dUm === null).length;
  const methodDesc = r.mode === 'bscan' ? 'B-scan ground truth'
                    : r.mode === 'hybrid' ? 'B-scan ground truth + red-free supplement (in B-scan-weak sectors)'
                    : 'red-free fallback';
  children.push(p(`Detection method: ${methodDesc}. ${usableCount} vessels measured + ${excludedCount} excluded.`,
    { run: { size: 18 } }));

  const vw = [500, 1200, 1100, 900, 1300, 1300, 1300, 1700];
  const vesselRows = [new TableRow({
    tableHeader: true,
    children: [
      cell('#', vw[0], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('Sector', vw[1], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('Method', vw[2], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('Angle (°)', vw[3], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('Diameter (µm)', vw[4], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('Range', vw[5], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('Area (µm²)', vw[6], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('Status', vw[7], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER })
    ]
  })];
  let vi = 0;
  for (const v of r.vessels) {
    vi++;
    const ok = v.dUm !== null;
    const methodLabel = v.source === 'bscan' ? 'B-scan'
                      : v.source === 'redfree-supplement' ? 'red-free suppl.'
                      : 'red-free';
    const methodColor = v.source === 'redfree-supplement' ? 'AA3300' : '000000';
    vesselRows.push(new TableRow({
      children: [
        cell(vi, vw[0], { fill: Q_COLORS[v.quadrant], align: AlignmentType.CENTER, bold: true }),
        cell(`${Q_NAMES[v.quadrant]} (${v.quadrant})`, vw[1], { fill: Q_COLORS[v.quadrant] }),
        cell(methodLabel, vw[2], { align: AlignmentType.CENTER, color: methodColor, italics: v.source === 'redfree-supplement' }),
        cell(v.angle.toFixed(1), vw[3], { align: AlignmentType.CENTER }),
        cell(ok ? v.dMid.toFixed(0) : '—', vw[4], { align: AlignmentType.CENTER, bold: true, color: ok ? '000000' : 'AA4400', italics: !ok }),
        cell(ok ? `${v.dMin.toFixed(0)} – ${v.dMax.toFixed(0)}` : '—', vw[5], { align: AlignmentType.CENTER, italics: !ok }),
        cell(ok ? `${v.aMid.toFixed(0)} (${v.aMin.toFixed(0)} – ${v.aMax.toFixed(0)})` : '—', vw[6], { align: AlignmentType.CENTER, italics: !ok }),
        cell(ok ? `${(v.contrast * 100).toFixed(1)}% contrast` : `EXCLUDED — ${v.flag || ''}`, vw[7], { align: AlignmentType.CENTER, color: ok ? '000000' : 'AA4400', italics: !ok })
      ]
    }));
  }
  children.push(new Table({ width: { size: 9300, type: WidthType.DXA }, columnWidths: vw, rows: vesselRows }));

  // Quadrant summary
  children.push(h2('Per-Quadrant Vessel / RNFL Analysis'));
  children.push(p('Cross-sectional area: vessel area A = π(d/2)²; RNFL area = arc length per quadrant (905.8 µm × 3 = 2717.5 µm) × quadrant-mean RNFL thickness.',
    { run: { size: 18 } }));

  const qw = [1500, 1100, 1100, 1500, 1700, 1700, 700];
  const quadRows = [new TableRow({
    tableHeader: true,
    children: [
      cell('Quadrant', qw[0], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('Vessels', qw[1], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('RNFL (µm)', qw[2], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('RNFL Area (µm²)', qw[3], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('Vessel Area (µm²)', qw[4], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('Range', qw[5], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('% RNFL', qw[6], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER })
    ]
  })];
  for (const q of ['S', 'N', 'I', 'T']) {
    const s = r.aggregation.summary[q];
    const vL = s.nExcluded > 0 ? `${s.n} (+${s.nExcluded} excl.)` : `${s.n}`;
    quadRows.push(new TableRow({
      children: [
        cell(`${Q_NAMES[q]} (${q})`, qw[0], { fill: Q_COLORS[q], bold: true }),
        cell(vL, qw[1], { align: AlignmentType.CENTER }),
        cell(s.rnflThick.toFixed(0), qw[2], { align: AlignmentType.CENTER }),
        cell(s.rnflA.toFixed(0), qw[3], { align: AlignmentType.CENTER }),
        cell(s.sumAMid.toFixed(0), qw[4], { align: AlignmentType.CENTER, bold: true }),
        cell(`${s.sumAMin.toFixed(0)} – ${s.sumAMax.toFixed(0)}`, qw[5], { align: AlignmentType.CENTER }),
        cell(`${s.pctMid.toFixed(1)}%`, qw[6], { align: AlignmentType.CENTER, bold: true })
      ]
    }));
  }
  const totV = ['S', 'N', 'I', 'T'].reduce((sum, q) => sum + r.aggregation.summary[q].n, 0);
  const totEx = ['S', 'N', 'I', 'T'].reduce((sum, q) => sum + r.aggregation.summary[q].nExcluded, 0);
  quadRows.push(new TableRow({
    children: [
      cell('Total / Global', qw[0], { fill: 'D5D5D5', bold: true }),
      cell(totEx > 0 ? `${totV} (+${totEx} excl.)` : `${totV}`, qw[1], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell(r.rnfl.G.toFixed(0), qw[2], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell(r.aggregation.totRA.toFixed(0), qw[3], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell(r.aggregation.totA.toFixed(0), qw[4], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER }),
      cell('', qw[5], { fill: 'D5D5D5' }),
      cell(`${r.aggregation.globalPct.toFixed(1)}%`, qw[6], { fill: 'D5D5D5', bold: true, align: AlignmentType.CENTER })
    ]
  }));
  children.push(new Table({ width: { size: 9300, type: WidthType.DXA }, columnWidths: qw, rows: quadRows }));

  // Annotated images
  if (rfBuf) {
    children.push(h2('Annotated Red-Free Image'));
    const rfDim = r.annRfDim || { width: 640, height: 385 };
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new ImageRun({
        data: rfBuf,
        transformation: { width: 480, height: Math.round(480 * rfDim.height / rfDim.width) },
        type: 'png'
      })]
    }));
  }
  if (bsBuf) {
    children.push(h2('Annotated B-scan'));
    const bsDim = r.annBsDim || { width: 1150, height: 270 };
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new ImageRun({
        data: bsBuf,
        transformation: { width: 720, height: Math.round(720 * bsDim.height / bsDim.width) },
        type: 'png'
      })]
    }));
  }

  // Methodology notes
  children.push(h2('Methodology Notes'));
  children.push(p('• Circle centre fitted from green dashed-circle pixel positions (least-squares).', { run: { size: 16 } }));
  children.push(p(`• Scan radius 1730 µm; circle radius ${r.circle?.r?.toFixed?.(1) ?? '—'} px; scale ${r.circle?.scale?.toFixed?.(2) ?? '—'} µm/px.`,
    { run: { size: 16 } }));
  const vesselSourceLine = r.mode === 'bscan' ? 'B-scan deep-band shadow detection (≥15% contrast)'
                          : r.mode === 'hybrid' ? 'B-scan deep-band shadows where contrast ≥15%, supplemented with red-free brightness-profile detections in regions B-scan missed'
                          : 'red-free brightness profile around the fitted circle';
  children.push(p(`• Vessel positions from ${vesselSourceLine}.`, { run: { size: 16 } }));
  children.push(p('• Diameter measured from shadow FWHM with conservative background reference. ±20% range provided to bracket measurement uncertainty.', { run: { size: 16 } }));
  children.push(p(`• ${r.eye} orientation: Temporal = ${r.eye === 'OD' ? 'LEFT' : 'RIGHT'} side of disc, Nasal = ${r.eye === 'OD' ? 'RIGHT' : 'LEFT'} side.`, { run: { size: 16 } }));
  children.push(p('• Generated by OCT Vessel Analyzer. For research and clinical decision support; not a diagnostic device.',
    { run: { size: 16, italics: true, color: '888888' } }));

  return children;
}
