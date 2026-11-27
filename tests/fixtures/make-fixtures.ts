/**
 * Synthetic .docx fixtures (execution context §6, phase 2.5).
 *
 * Built with JSZip + raw WordprocessingML so we control every run exactly.
 * No personal content: these files are safe for CI. Output is deterministic
 * (fixed zip dates, seeded random text), and written to tests/fixtures/generated/.
 *
 * Run by the Vitest globalSetup, or by hand with `npm run fixtures`.
 * This file runs directly under Node's type stripping: keep it free of
 * enums, namespaces, parameter properties and path aliases.
 */
import { mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const THIS_FILE = fileURLToPath(import.meta.url);
export const GENERATED_DIR = join(dirname(THIS_FILE), 'generated');

export const FIXTURES = {
  headingsList: 'fx-a-headings-list.docx',
  table: 'fx-b-table.docx',
  inline: 'fx-c-inline.docx',
  otherColour: 'fx-d-yellow.docx',
  large: 'fx-e-large.docx',
} as const;

export type FixtureName = (typeof FIXTURES)[keyof typeof FIXTURES];

/* ------------------------------------------------------------------ */
/* WordprocessingML builders                                           */
/* ------------------------------------------------------------------ */

const FIXED_DATE = new Date('2026-01-01T00:00:00Z');
const FONT = 'Times New Roman';
const BODY_SIZE = 32; // half-points: 16 pt

type Highlight = 'darkYellow' | 'yellow';

type RunOptions = {
  highlight?: Highlight;
  bold?: boolean;
  italic?: boolean;
  size?: number; // half-points
  style?: string; // character style id (w:rStyle)
};

function escapeXml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

/** Run properties in schema order, the way WPS writes them (highlight before lang). */
function rPr(o: RunOptions = {}): string {
  const size = o.size ?? BODY_SIZE;
  return (
    '<w:rPr>' +
    (o.style ? `<w:rStyle w:val="${o.style}"/>` : '') +
    `<w:rFonts w:hint="default" w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>` +
    (o.bold ? '<w:b/><w:bCs/>' : '') +
    (o.italic ? '<w:i/><w:iCs/>' : '') +
    `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>` +
    (o.highlight ? `<w:highlight w:val="${o.highlight}"/>` : '') +
    '<w:lang w:val="en-US"/>' +
    '</w:rPr>'
  );
}

function t(text: string): string {
  const preserve = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
  return `<w:t${preserve}>${escapeXml(text)}</w:t>`;
}

/** One text run. */
function r(text: string, o: RunOptions = {}): string {
  return `<w:r>${rPr(o)}${t(text)}</w:r>`;
}

/** A run holding one special element: <w:tab/>, <w:br/>, <w:noBreakHyphen/>, <w:softHyphen/>. */
function special(element: string, o: RunOptions = {}): string {
  return `<w:r>${rPr(o)}<${element}/></w:r>`;
}

function p(content: string, pPr = ''): string {
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content}</w:p>`;
}

const pStyle = (id: string) => `<w:pStyle w:val="${id}"/>`;
const numPr = (numId: number, ilvl = 0) =>
  `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;

/** Paragraph from a simple list of [text, highlighted?] pieces. */
function para(pieces: [string, Highlight?][], pPr = '', o: RunOptions = {}): string {
  return p(pieces.map(([text, highlight]) => r(text, { ...o, highlight })).join(''), pPr);
}

type Cell = { content: string; width: number; gridSpan?: number; vMerge?: 'restart' | 'continue' };

function tc(cell: Cell): string {
  const props =
    `<w:tcW w:w="${cell.width}" w:type="dxa"/>` +
    (cell.gridSpan ? `<w:gridSpan w:val="${cell.gridSpan}"/>` : '') +
    (cell.vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : '') +
    (cell.vMerge === 'continue' ? '<w:vMerge/>' : '');
  return `<w:tc><w:tcPr>${props}</w:tcPr>${cell.content}</w:tc>`;
}

function tbl(gridCols: number[], rows: Cell[][]): string {
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
    .join('');
  return (
    '<w:tbl>' +
    `<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid>${gridCols.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
    rows.map((row) => `<w:tr>${row.map(tc).join('')}</w:tr>`).join('') +
    '</w:tbl>'
  );
}

/* ------------------------------------------------------------------ */
/* Package parts                                                       */
/* ------------------------------------------------------------------ */

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** A4 page with WPS default margins. */
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>' +
  '</w:sectPr>';

function documentXml(body: string): string {
  return (
    XML_HEAD +
    `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${body}${SECT_PR}</w:body></w:document>`
  );
}

const CONTENT_TYPES =
  XML_HEAD +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
  '</Types>';

const ROOT_RELS =
  XML_HEAD +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>` +
  '</Relationships>';

const HYPERLINK_REL_ID = 'rIdLink1';

function documentRels(withHyperlink: boolean): string {
  return (
    XML_HEAD +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/>` +
    (withHyperlink
      ? `<Relationship Id="${HYPERLINK_REL_ID}" Type="${REL}/hyperlink" Target="https://example.com/" TargetMode="External"/>`
      : '') +
    '</Relationships>'
  );
}

const STYLES =
  XML_HEAD +
  `<w:styles xmlns:w="${NS_W}">` +
  '<w:docDefaults><w:rPrDefault><w:rPr>' +
  `<w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>` +
  `<w:sz w:val="${BODY_SIZE}"/><w:szCs w:val="${BODY_SIZE}"/><w:lang w:val="en-US"/>` +
  '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
  '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>' +
  '<w:rPr><w:b/><w:bCs/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
  '<w:pPr><w:keepNext/><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr>' +
  '<w:rPr><w:b/><w:bCs/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>' +
  '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/>' +
  '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
  '</w:styles>';

const NUMBERING =
  XML_HEAD +
  `<w:numbering xmlns:w="${NS_W}">` +
  '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>' +
  '<w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/>' +
  `<w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}"/></w:rPr></w:lvl></w:abstractNum>` +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
  '</w:numbering>';

const NUMBERED = 1;
const BULLETED = 2;

async function packageDocx(body: string, withHyperlink = false): Promise<Uint8Array> {
  const zip = new JSZip();
  // createFolders: false — no "word/" or "_rels/" directory entries (WPS/Word don't write
  // them, and JSZip would stamp them with the current time, breaking determinism).
  const add = (path: string, content: string) =>
    zip.file(path, content, { date: FIXED_DATE, createFolders: false });
  add('[Content_Types].xml', CONTENT_TYPES);
  add('_rels/.rels', ROOT_RELS);
  add('word/document.xml', documentXml(body));
  add('word/_rels/document.xml.rels', documentRels(withHyperlink));
  add('word/styles.xml', STYLES);
  add('word/numbering.xml', NUMBERING);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/* ------------------------------------------------------------------ */
/* The fixtures                                                        */
/* ------------------------------------------------------------------ */

/** (a) Heading 1/2 + a numbered list (+ a bulleted list). */
function headingsListBody(): string {
  return [
    para([['Profile '], ['summary', 'darkYellow']], pStyle('Heading1'), { size: 44, bold: true }),
    para(
      [
        ['I build '],
        ['full-stack platform', 'darkYellow'],
        [' tools for small teams, with a focus on '],
        ['RAG', 'darkYellow'],
        [' search.'],
      ],
      '',
    ),
    para([['Main '], ['stack', 'darkYellow']], pStyle('Heading2'), { size: 36, bold: true }),
    para([['TypeScript', 'darkYellow'], [' and React on the front end.']], numPr(NUMBERED)),
    para([['Node.js', 'darkYellow'], [' and PostgreSQL on the back end.']], numPr(NUMBERED)),
    para([['Deployed on Vercel since '], ['2026', 'darkYellow'], ['.']], numPr(NUMBERED)),
    para([['A bullet with a '], ['marked', 'darkYellow'], [' word.']], numPr(BULLETED)),
  ].join('');
}

/** (b) A table with gridSpan + vMerge + a nested table. */
function tableBody(): string {
  const W = 2600;
  const nested = tbl(
    [1300, 1300],
    [
      [
        { width: 1300, content: para([['Inner', 'darkYellow'], [' one']]) },
        { width: 1300, content: para([['Inner two']]) },
      ],
      [
        { width: 1300, content: para([['Inner three']]) },
        { width: 1300, content: para([['nested', 'darkYellow'], [' cell']]) },
      ],
    ],
  );
  const table = tbl(
    [W, W, W],
    [
      [
        {
          width: W * 2,
          gridSpan: 2,
          content: para([['Merged', 'darkYellow'], [' across two columns']]),
        },
        { width: W, content: para([['Top right']]) },
      ],
      [
        {
          width: W,
          vMerge: 'restart',
          content: para([['Merged '], ['down', 'darkYellow'], [' two rows']]),
        },
        { width: W, content: para([['Middle '], ['cell', 'darkYellow']]) },
        { width: W, content: para([['Plain cell']]) },
      ],
      [
        { width: W, vMerge: 'continue', content: p('') },
        // A cell holding a table must still end with a paragraph.
        { width: W, content: nested + para([['After the inner table']]) },
        { width: W, content: para([['Last', 'darkYellow'], [' cell']]) },
      ],
    ],
  );
  return [
    para([['Text '], ['before', 'darkYellow'], [' the table.']]),
    table,
    para([['Text after the '], ['table', 'darkYellow'], ['.']]),
  ].join('');
}

/** (c) One paragraph with a hyperlink, tab, line break, no-break hyphen, soft hyphen and a split word. */
function inlineBody(): string {
  const dy: RunOptions = { highlight: 'darkYellow' };
  const content = [
    r('Visit '),
    `<w:hyperlink r:id="${HYPERLINK_REL_ID}" w:history="1">`,
    r('our', { style: 'Hyperlink', highlight: 'darkYellow' }),
    r(' website', { style: 'Hyperlink' }),
    '</w:hyperlink>',
    r(' now.'),
    special('w:tab'),
    r('Name:'),
    special('w:br'),
    r('a well'),
    special('w:noBreakHyphen'),
    r('known '),
    // One word split over two runs, as WPS often writes it.
    r('Engin', dy),
    r('eering', dy),
    r(' and high'),
    special('w:softHyphen'),
    r('light test.'),
  ].join('');
  return p(content);
}

/** (d) A yellow word next to darkYellow words. */
function otherColourBody(): string {
  return [
    para([
      ['The '],
      ['urgent', 'yellow'],
      [' '],
      ['deadline matters', 'darkYellow'],
      [' for everyone.'],
    ]),
    para([['Both '], ['yellow', 'yellow'], ['darkYellow', 'darkYellow'], [' touch here.']]),
  ].join('');
}

/** Small deterministic PRNG (mulberry32). */
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let x = a;
    x = Math.imul(x ^ (x >>> 15), x | 1);
    x ^= x + Math.imul(x ^ (x >>> 7), x | 61);
    return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
  };
}

const LOREM = (
  'software engineering project platform design review report summary answer interview data model ' +
  'user interface test release version client server browser phone document paragraph highlight ' +
  'word table heading list team plan risk cost value quality system process result method study ' +
  'research analysis feature module service deploy build cloud storage network security'
).split(' ');

export const LARGE_PARAGRAPHS = 1400;

/** (e) ~100 pages: 1400 paragraphs of 12–28 words (~28 000 words at 16 pt). */
function largeBody(): string {
  const next = rng(13);
  const pick = () => LOREM[Math.floor(next() * LOREM.length)] ?? 'word';
  const out: string[] = [];
  for (let i = 0; i < LARGE_PARAGRAPHS; i++) {
    if (i % 50 === 0) {
      out.push(para([[`Section ${i / 50 + 1}`]], pStyle('Heading1'), { size: 44, bold: true }));
    }
    const count = 12 + Math.floor(next() * 17);
    const pieces: [string, Highlight?][] = [];
    let plain = '';
    for (let w = 0; w < count; w++) {
      const word = pick();
      const sep = w === count - 1 ? '.' : ' ';
      const roll = next();
      if (roll < 0.12 && w < count - 2) {
        // A highlighted pair joined by its gap, like "real estate".
        if (plain) pieces.push([plain]);
        const second = pick();
        pieces.push([`${word} ${second}`, 'darkYellow']);
        plain = ' ';
        w++;
      } else if (roll < 0.25) {
        if (plain) pieces.push([plain]);
        pieces.push([word, 'darkYellow']);
        plain = sep;
      } else {
        plain += word + sep;
      }
    }
    if (plain) pieces.push([plain]);
    out.push(para(pieces, i % 7 === 3 ? numPr(BULLETED) : ''));
  }
  return out.join('');
}

/** Builds every fixture in memory (deterministic bytes). */
export async function buildFixtures(): Promise<Record<FixtureName, Uint8Array>> {
  return {
    [FIXTURES.headingsList]: await packageDocx(headingsListBody()),
    [FIXTURES.table]: await packageDocx(tableBody()),
    [FIXTURES.inline]: await packageDocx(inlineBody(), true),
    [FIXTURES.otherColour]: await packageDocx(otherColourBody()),
    [FIXTURES.large]: await packageDocx(largeBody()),
  };
}

/** Writes every fixture to `outDir`, skipping files whose bytes are already identical. */
export async function makeFixtures(outDir: string = GENERATED_DIR): Promise<string[]> {
  mkdirSync(outDir, { recursive: true });
  const built = await buildFixtures();
  const written: string[] = [];
  for (const [name, bytes] of Object.entries(built)) {
    const path = join(outDir, name);
    if (existsSync(path) && Buffer.compare(readFileSync(path), bytes) === 0) continue;
    writeFileSync(path, bytes);
    written.push(name);
  }
  return written;
}

// `npm run fixtures` / `node tests/fixtures/make-fixtures.ts`
if (process.argv[1] && resolve(process.argv[1]) === THIS_FILE) {
  const written = await makeFixtures();
  console.log(
    written.length
      ? `Wrote ${written.join(', ')} to ${GENERATED_DIR}`
      : `Fixtures already up to date in ${GENERATED_DIR}`,
  );
}
