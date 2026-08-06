// Shared RTL (Hebrew) DOCX building helpers.
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageBreak, LevelFormat, convertInchesToTwip,
} = require('docx');

const FONT = 'Arial';
const BODY = 22;   // 11pt (half-points)
const SMALL = 19;

// Palette
const INK = '1A1A1A';
const MUTED = '5C5C5C';
const GOLD = '8A6A12';
const GREEN = '2E5E3A';
const RED = 'A32020';
const BLUE = '1F4E79';

const run = (text, o = {}) => new TextRun({
  text, rightToLeft: true, font: FONT,
  size: o.size || BODY, bold: !!o.bold, italics: !!o.italics,
  color: o.color || INK, strike: !!o.strike, underline: o.underline ? {} : undefined,
});

// Split on \n into separate runs joined by real line breaks — Word collapses
// a literal \n inside <w:t>, which would flatten verse and rewritten passages.
const runsOf = (text, o = {}) => String(text).split('\n').map((seg, i) => new TextRun({
  text: seg, rightToLeft: true, font: FONT,
  size: o.size || BODY, bold: !!o.bold, italics: !!o.italics,
  color: o.color || INK, break: i > 0 ? 1 : undefined,
}));

// Rich paragraph: pass a string, or array of [text, opts] pairs.
const P = (content, o = {}) => {
  const kids = typeof content === 'string'
    ? runsOf(content, o)
    : content.flatMap(c => (typeof c === 'string' ? runsOf(c, o) : runsOf(c[0], { ...o, ...c[1] })));
  return new Paragraph({
    children: kids,
    bidirectional: true,
    alignment: o.align || AlignmentType.RIGHT,
    spacing: { after: o.after === undefined ? 120 : o.after, before: o.before || 0, line: o.line || 300 },
    indent: o.indent,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
    border: o.leftBar
      ? { right: { style: BorderStyle.SINGLE, size: 18, color: o.leftBar, space: 8 } }
      : (o.box ? {
          top: { style: BorderStyle.SINGLE, size: 4, color: o.box, space: 6 },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: o.box, space: 6 },
          left: { style: BorderStyle.SINGLE, size: 4, color: o.box, space: 6 },
          right: { style: BorderStyle.SINGLE, size: 4, color: o.box, space: 6 },
        } : undefined),
    keepNext: !!o.keepNext,
    pageBreakBefore: !!o.pageBreakBefore,
  });
};

const H = (text, level, o = {}) => new Paragraph({
  children: [run(text, { bold: true, size: o.size || (level === 1 ? 34 : level === 2 ? 27 : 23), color: o.color || (level === 1 ? INK : level === 2 ? BLUE : INK) })],
  heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
  bidirectional: true,
  alignment: AlignmentType.RIGHT,
  spacing: { before: o.before === undefined ? (level === 1 ? 400 : 320) : o.before, after: o.after === undefined ? 160 : o.after },
  pageBreakBefore: !!o.pageBreakBefore,
  keepNext: true,
});

const H1 = (t, o) => H(t, 1, o);
const H2 = (t, o) => H(t, 2, o);
const H3 = (t, o) => H(t, 3, o);

const BULLET = (content, o = {}) => {
  const kids = typeof content === 'string'
    ? runsOf(content, o)
    : content.flatMap(c => (typeof c === 'string' ? runsOf(c, o) : runsOf(c[0], { ...o, ...c[1] })));
  return new Paragraph({
    children: kids,
    bidirectional: true,
    alignment: AlignmentType.RIGHT,
    numbering: { reference: o.numbered ? 'num-list' : 'bullet-list', level: o.level || 0 },
    spacing: { after: 80, line: 290 },
  });
};

// Manuscript quotation — indented, distinct.
const QUOTE = (text, o = {}) => P(text, {
  italics: true, color: o.color || MUTED, size: SMALL,
  indent: { right: 360, left: 360 }, after: 60, line: 280,
  leftBar: o.bar || 'CCCCCC', ...o,
});

// Editorial note callout.
const NOTE = (label, text, color = GOLD, fill = 'FBF6E7') =>
  P([[label + '  ', { bold: true, color }], [text, {}]], { fill, indent: { right: 120, left: 120 }, after: 140, size: SMALL, leftBar: color });

const RULE = () => new Paragraph({
  children: [run('')], bidirectional: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'D0D0D0', space: 1 } },
  spacing: { before: 120, after: 200 },
});

const SPACER = (n = 1) => Array.from({ length: n }, () => P('', { after: 0 }));
const BREAK = () => new Paragraph({ children: [new PageBreak()], bidirectional: true });

// RTL table. cols = array of header strings; widths in DXA; rows = array of arrays.
// Cell content may be a string or an array of [text, opts] pairs.
const TABLE = (cols, widths, rows, o = {}) => {
  const total = widths.reduce((a, b) => a + b, 0);
  const cell = (content, w, opts = {}) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: 'auto' } : undefined,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: (Array.isArray(content) && Array.isArray(content[0]) ? content : [content])
      .map(c => (Array.isArray(c) && Array.isArray(c[0]))
        ? P(c, { after: 40, size: SMALL, ...opts })
        : P(c, { after: 40, size: SMALL, ...opts })),
  });
  return new Table({
    columnWidths: widths,
    width: { size: total, type: WidthType.DXA },
    visuallyRightToLeft: true,
    rows: [
      new TableRow({
        tableHeader: true,
        children: cols.map((c, i) => cell(c, widths[i], { fill: o.headerFill || 'EDEFF2', bold: true })),
      }),
      ...rows.map((r, ri) => new TableRow({
        children: r.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? 'FAFAFA' : undefined })),
      })),
    ],
  });
};

const numbering = {
  config: [
    {
      reference: 'bullet-list',
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.RIGHT,
          style: { paragraph: { indent: { right: 360, hanging: 220 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '◦', alignment: AlignmentType.RIGHT,
          style: { paragraph: { indent: { right: 720, hanging: 220 } } } },
      ],
    },
    {
      reference: 'num-list',
      levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.RIGHT,
          style: { paragraph: { indent: { right: 400, hanging: 260 } } } },
      ],
    },
  ],
};

async function build(children, outPath, title) {
  const doc = new Document({
    numbering,
    styles: {
      default: {
        document: { run: { font: FONT, size: BODY, color: INK }, paragraph: { bidirectional: true, alignment: AlignmentType.RIGHT } },
      },
    },
    sections: [{
      properties: {
        page: {
          margin: { top: convertInchesToTwip(0.9), bottom: convertInchesToTwip(0.9), left: convertInchesToTwip(0.9), right: convertInchesToTwip(0.9) },
        },
      },
      children,
    }],
  });
  const buf = await Packer.toBuffer(doc);
  require('fs').writeFileSync(outPath, buf);
  console.log('wrote', outPath, (buf.length / 1024).toFixed(1) + ' KB');
}

module.exports = { P, H1, H2, H3, BULLET, QUOTE, NOTE, RULE, SPACER, BREAK, TABLE, build, run,
  INK, MUTED, GOLD, GREEN, RED, BLUE, BODY, SMALL, AlignmentType };
