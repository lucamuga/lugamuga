// Gemeinsame Bausteine für die Word-Dokumente (docx-js).
const {
  Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle,
} = require('docx');

const FONT = 'Calibri';
const ACC = '9E3B24';
const W = 9638; // Inhaltsbreite A4 mit 2 cm Rändern

const runs = (x) => (Array.isArray(x) ? x : [x]).map(r => (typeof r === 'string' ? t(r) : r));
const t = (text, o = {}) => new TextRun({ text, font: FONT, size: 21, ...o });
const b = (text, o = {}) => t(text, { bold: true, ...o });
const P = (x, o = {}) => new Paragraph({ spacing: { after: 120, line: 276 }, ...o, children: runs(x) });
const H1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: false, spacing: { before: 360, after: 160 }, children: [new TextRun({ text, font: FONT, size: 32, bold: true, color: '1C2430' })] });
const H2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 240, after: 100 }, children: [new TextRun({ text, font: FONT, size: 25, bold: true, color: ACC })] });
const H3 = (text) => new Paragraph({ keepNext: true, spacing: { before: 160, after: 60 }, children: [new TextRun({ text, font: FONT, size: 22, bold: true, color: '1C2430' })] });
const BL = (x, lvl = 0) => new Paragraph({ numbering: { reference: 'bul', level: lvl }, spacing: { after: 60, line: 264 }, children: runs(x) });
const NL = (x, ref = 'num') => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 60, line: 264 }, children: runs(x) });
const gap = () => new Paragraph({ spacing: { after: 80 }, children: [] });

const border = { style: BorderStyle.SINGLE, size: 4, color: 'D5D2CC' };
const borders = { top: border, bottom: border, left: border, right: border };

function table(widths, rows, opts = {}) {
  const sum = widths.reduce((a, c) => a + c, 0);
  const ws = widths.map(w => Math.round(w * W / sum));
  ws[ws.length - 1] += W - ws.reduce((a, c) => a + c, 0);
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: ws,
    rows: rows.map((r, ri) => new TableRow({
      tableHeader: ri === 0,
      cantSplit: true,
      children: r.map((cell, ci) => {
        const head = ri === 0;
        const fill = head ? ACC : (opts.zebra && ri % 2 === 0 ? 'F6F4F1' : undefined);
        const txt = Array.isArray(cell) ? cell : [cell];
        return new TableCell({
          borders, width: { size: ws[ci], type: WidthType.DXA },
          shading: fill ? { fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
          margins: { top: 60, bottom: 60, left: 90, right: 90 },
          children: [new Paragraph({
            alignment: (opts.right && opts.right.includes(ci) && !head) ? AlignmentType.RIGHT : AlignmentType.LEFT,
            children: txt.map(x => (typeof x === 'string' ? t(x, { size: 18, bold: head || (opts.boldFirst && ci === 0), color: head ? 'FFFFFF' : undefined }) : x)),
          })],
        });
      }),
    })),
  });
}

// Farbige Kästen: Hinweis (rot), Tipp (grün), Info (grau)
const STYLES = {
  warn: { fill: 'F8E1DF', bar: 'B3362F' },
  tip: { fill: 'E3F1E8', bar: '2E7A4B' },
  info: { fill: 'F6F4F1', bar: ACC },
  key: { fill: 'F5E6E0', bar: ACC },
};
function box(title, lines, kind = 'key') {
  const s = STYLES[kind];
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      borders: { top: border, bottom: border, right: border, left: { style: BorderStyle.SINGLE, size: 24, color: s.bar } },
      shading: { fill: s.fill, type: ShadingType.CLEAR, color: 'auto' }, margins: { top: 100, bottom: 100, left: 160, right: 160 },
      children: [new Paragraph({ spacing: { after: 60 }, children: [b(title, { size: 21, color: s.bar })] }),
        ...lines.map(l => new Paragraph({ spacing: { after: 50 }, children: runs(l) }))],
    })] })],
  });
}

const numbering = { config: [
  { reference: 'bul', levels: [
    { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } },
    { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 260 } } } },
  ] },
  ...['num', 'num2', 'num3', 'num4'].map(ref => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { run: { bold: true, color: ACC }, paragraph: { indent: { left: 360, hanging: 300 } } } }] })),
] };

const styles = {
  default: { document: { run: { font: FONT, size: 21 } } },
  paragraphStyles: [
    { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 32, bold: true }, paragraph: { outlineLevel: 0 } },
    { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 25, bold: true }, paragraph: { outlineLevel: 1 } },
  ],
};

module.exports = { FONT, ACC, W, t, b, P, H1, H2, H3, BL, NL, gap, table, box, numbering, styles };
