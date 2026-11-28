// @vitest-environment jsdom
// Phase 5.3: round trip read → export → read on every fixture, the generated ones
// (always run) and the owner's real WPS files (local only, skipped in CI).
//
// The expected values are checked against the raw XML text and the raw zip bytes,
// not through the reader and writer under test.
import JSZip from 'jszip';
import { beforeAll, describe, expect, it } from 'vitest';
import { rangesFromUnits } from '@/core/marks';
import { joinableFromGaps, normalizeMarks } from '@/core/units';
import { type DocumentMarks, exportDocx } from '@/docx/export';
import { openDocx } from '@/docx/open';
import { readDocumentMarks } from '@/docx/read-marks';
import type { DocModel } from '@/docx/types';
import { FIXTURES } from './fixtures/make-fixtures.ts';
import {
  describeWithLocalFixture,
  readGeneratedFixture,
  readLocalFixture,
} from './helpers/fixtures.ts';
import { physicalOrder, rawRecords, reversePhysicalOrder } from './helpers/zip.ts';

const DOCUMENT = 'word/document.xml';
const DY = 'darkYellow';

/* ---------------------------- raw XML reading ----------------------------- */

const RUN = /<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/g;
const CONTENT = /<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:(tab|br|cr|noBreakHyphen)(?: [^>]*)?\/>/g;
const SPECIAL: Record<string, string> = { tab: '\t', br: '\n', cr: '\n', noBreakHyphen: '-' };

const unescape = (text: string) =>
  text
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&amp;', '&');

type RawChar = {
  char: string;
  highlight: string | null;
  /** The raw bytes of the run's rPr without its w:highlight. */
  rPr: string;
};

/** One entry per visible character of the document, read from the XML text. */
function rawChars(xml: string): RawChar[] {
  const out: RawChar[] = [];
  for (const [run] of xml.matchAll(RUN)) {
    const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(run)?.[0] ?? '';
    const highlight = /<w:highlight w:val="(\w+)"\/>/.exec(rPr)?.[1] ?? null;
    let rest = rPr.replace(/<w:highlight w:val="\w+"\/>/, '');
    if (rest === '<w:rPr></w:rPr>' || rest === '<w:rPr/>') rest = '';
    for (const match of run.slice(run.indexOf(rPr) + rPr.length).matchAll(CONTENT)) {
      const text = match[2] ? SPECIAL[match[2]]! : unescape(match[1]!);
      for (const char of text.split('')) out.push({ char, highlight, rPr: rest });
    }
  }
  return out;
}

const textOf = (chars: RawChar[]) => chars.map((c) => c.char).join('');
const highlightsOf = (chars: RawChar[]) => chars.map((c) => c.highlight ?? '.').join(' ');
const rPrOf = (chars: RawChar[]) => chars.map((c) => c.rPr).join('\n');
/** The document without its runs: paragraphs, properties, tables, links, section… */
const skeleton = (xml: string) => xml.replace(RUN, '');

/**
 * word/document.xml as text. The serializer writes the empty element
 * `<w:p></w:p>` of fixture (b) as `<w:p/>` (same XML), so both are read that way.
 */
const documentText = async (bytes: Uint8Array) => {
  const xml = await (await JSZip.loadAsync(bytes)).file(DOCUMENT)!.async('string');
  return xml.replaceAll('<w:p></w:p>', '<w:p/>');
};

/* --------------------------------- marks ---------------------------------- */

const marksOf = (doc: DocModel) =>
  new Map([...readDocumentMarks(doc)].map(([index, m]) => [index, m.units]));

/**
 * New marks made from the file's marks: every third word is toggled, in every
 * second paragraph all possible gaps are joined, and the first marked paragraph
 * loses its first marked word (so every file has a removal).
 */
function editedMarks(doc: DocModel): Map<number, number[]> {
  const before = marksOf(doc);
  const out = new Map<number, number[]>();
  const [first] = before.keys();
  for (const { index, words, gaps } of doc.paragraphs) {
    const units = new Set(before.get(index) ?? []);
    for (let word = index % 3; word < words.length; word += 3) {
      if (!units.delete(2 * word)) units.add(2 * word);
    }
    if (index % 2 === 0) gaps.forEach((_, i) => units.add(2 * i + 1));
    if (index === first) units.delete(before.get(index)![0]!);
    const valid = normalizeMarks([...units], words.length, joinableFromGaps(gaps));
    if (valid.length > 0) out.set(index, valid);
  }
  return out;
}

/** Which characters of the whole document `marks` want in darkYellow. */
function wantedChars(doc: DocModel, marks: DocumentMarks): boolean[] {
  return doc.paragraphs.flatMap(({ index, text, words, gaps }) => {
    const wanted = new Array<boolean>(text.length).fill(false);
    for (const [start, end] of rangesFromUnits(marks.get(index) ?? [], words, gaps)) {
      wanted.fill(true, start, end);
    }
    return wanted;
  });
}

const exportBytes = async (doc: DocModel, marks: DocumentMarks) =>
  new Uint8Array(await (await exportDocx(doc, marks)).arrayBuffer());

/* ------------------------------- the suite -------------------------------- */

type Options = {
  /** False for a file with a partial highlight, which the export rounds up (2026-11-27). */
  canonical?: boolean;
};

function roundTrip(name: string, read: () => Uint8Array, { canonical = true }: Options = {}) {
  let source: Uint8Array;
  let doc: DocModel;
  let sourceXml: string;
  let unchanged: Uint8Array;
  let edited: Uint8Array;
  let marks: Map<number, number[]>;

  beforeAll(async () => {
    source = new Uint8Array(read());
    doc = await openDocx(source, name);
    sourceXml = await documentText(source);
    marks = editedMarks(doc);
    unchanged = await exportBytes(doc, marksOf(doc));
    edited = await exportBytes(doc, marks);
  }, 120_000);

  it('the raw reading used by these tests sees the same text as the reader', () => {
    expect(textOf(rawChars(sourceXml))).toBe(doc.paragraphs.map((p) => p.text).join(''));
  });

  it.runIf(canonical)(
    '(a) no change: identical text, highlight per character and run properties',
    async () => {
      const xml = await documentText(unchanged);
      const [before, after] = [rawChars(sourceXml), rawChars(xml)];
      expect(textOf(after)).toBe(textOf(before));
      expect(highlightsOf(after)).toBe(highlightsOf(before));
      expect(rPrOf(after)).toBe(rPrOf(before));
      // Stronger than asked: nothing at all changes in word/document.xml.
      expect(xml).toBe(sourceXml);
    },
    60_000,
  );

  it('(b) added and removed marks: expected highlight per character', async () => {
    const before = rawChars(sourceXml);
    const after = rawChars(await documentText(edited));
    const wanted = wantedChars(doc, marks);
    const expected = before.map((c, i) => ({
      ...c,
      highlight: wanted[i] ? DY : c.highlight === DY ? null : c.highlight,
    }));
    // The edit really adds and removes highlights.
    expect(before.some((c, i) => c.highlight !== DY && expected[i]!.highlight === DY)).toBe(true);
    expect(before.some((c, i) => c.highlight === DY && expected[i]!.highlight !== DY)).toBe(true);

    expect(textOf(after)).toBe(textOf(before));
    expect(highlightsOf(after)).toBe(highlightsOf(expected));
  }, 60_000);

  it('(b) every other rPr is byte identical, and nothing outside the runs changes', async () => {
    const xml = await documentText(edited);
    expect(rPrOf(rawChars(xml))).toBe(rPrOf(rawChars(sourceXml)));
    expect(skeleton(xml)).toBe(skeleton(sourceXml));
  }, 60_000);

  it('(b) the exported file reads back with the new marks', async () => {
    const reopened = await openDocx(edited, name);
    expect(marksOf(reopened)).toEqual(marks);
    expect(reopened.paragraphs.map((p) => p.text)).toEqual(doc.paragraphs.map((p) => p.text));
  }, 60_000);

  it('(c) every other zip entry is identical, in the same order', async () => {
    const original = rawRecords(source);
    expect(original[0]?.name).toBe('[Content_Types].xml');
    const originalZip = await JSZip.loadAsync(source);

    for (const bytes of [unchanged, edited]) {
      const records = rawRecords(bytes);
      expect(records.map((r) => r.name)).toEqual(original.map((r) => r.name));
      expect(physicalOrder(bytes)).toEqual(physicalOrder(source));
      const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
      for (const [i, record] of original.entries()) {
        if (record.name === DOCUMENT) continue;
        // Raw records (header + compressed data, and directory record).
        expect(records[i]!.local).toEqual(record.local);
        expect(records[i]!.central).toEqual(record.central);
        const entry = originalZip.files[record.name]!;
        if (!entry.dir) {
          expect(await zip.files[record.name]!.async('uint8array')).toEqual(
            await entry.async('uint8array'),
          );
        }
      }
    }
  }, 60_000);
}

describe.each(Object.values(FIXTURES))('round trip: generated %s', (name) => {
  // (d) holds "yellowdarkYellow", one word partly darkYellow: see the test below.
  roundTrip(name, () => readGeneratedFixture(name), {
    canonical: name !== FIXTURES.otherColour,
  });
});

describe('round trip: generated (d), the partial highlight', () => {
  it('(a) changes only that word, which becomes darkYellow as a whole', async () => {
    const source = readGeneratedFixture(FIXTURES.otherColour);
    const doc = await openDocx(source, FIXTURES.otherColour);
    const before = rawChars(await documentText(source));
    const after = rawChars(await documentText(await exportBytes(doc, marksOf(doc))));

    expect(textOf(after)).toBe(textOf(before));
    expect(rPrOf(after)).toBe(rPrOf(before));
    const changed = after.flatMap((c, i) => (c.highlight === before[i]!.highlight ? [] : [i]));
    const word = textOf(before).indexOf('yellowdarkYellow');
    expect(changed).toEqual([...'yellow'].map((_, i) => word + i));
    expect(changed.map((i) => after[i]!.highlight)).toEqual(new Array(6).fill(DY));
    // "urgent" stays yellow.
    expect(after[textOf(after).indexOf('urgent')]!.highlight).toBe('yellow');
  });
});

// The generated files have no folder entries: give them the shape of a WPS package.
describe('round trip: a package with folder entries stored in another order', () => {
  let shaped: Uint8Array;
  beforeAll(async () => {
    const from = await JSZip.loadAsync(readGeneratedFixture(FIXTURES.inline));
    const zip = new JSZip();
    const date = new Date('2026-01-01T00:00:00Z');
    for (const [path, entry] of Object.entries(from.files)) {
      const folder = path.slice(0, path.lastIndexOf('/') + 1);
      if (folder && !zip.files[folder]) {
        zip.file(folder, null, { dir: true, date, createFolders: false });
      }
      zip.file(path, await entry.async('uint8array'), { date, createFolders: false });
    }
    shaped = reversePhysicalOrder(
      await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
    );
    expect(rawRecords(shaped).map((r) => r.name)).toContain('word/');
  });

  roundTrip('shaped.docx', () => shaped);
});

describeWithLocalFixture('test.docx', 'round trip on the real WPS file', () => {
  roundTrip('test.docx', () => readLocalFixture('test.docx'));

  it('5.4 example: "Cyprus International" splits its run and nothing else', async () => {
    const source = readLocalFixture('test.docx');
    const doc = await openDocx(source, 'test.docx');
    const marks = marksOf(doc);
    // P1 words 6 and 7 and the gap between them.
    marks.set(1, [...(marks.get(1) ?? []), 12, 13, 14]);
    const bytes = await exportBytes(doc, marks);
    const [before, after] = [
      rawChars(await documentText(source)),
      rawChars(await documentText(bytes)),
    ];

    const start = textOf(before).indexOf('Cyprus International');
    const changed = after.flatMap((c, i) => (c.highlight === before[i]!.highlight ? [] : [i]));
    expect(changed).toEqual([...'Cyprus International'].map((_, i) => start + i));
    expect(changed.every((i) => after[i]!.highlight === DY)).toBe(true);
    expect(await documentText(bytes)).toContain('<w:t>Cyprus International</w:t>');
    expect(rPrOf(after)).toBe(rPrOf(before));
  });
});

describeWithLocalFixture('fx-a-headings-list.docx', 'round trip on the real headings file', () => {
  roundTrip('fx-a-headings-list.docx', () => readLocalFixture('fx-a-headings-list.docx'));

  it('has the WPS folder entries, and keeps them', async () => {
    const source = readLocalFixture('fx-a-headings-list.docx');
    const doc = await openDocx(source, 'fx-a-headings-list.docx');
    const folders = ['_rels/', 'docProps/', 'word/', 'word/_rels/', 'word/theme/'];
    const names = rawRecords(await exportBytes(doc, new Map())).map((r) => r.name);
    expect(names.filter((n) => n.endsWith('/'))).toEqual(folders);
  });
});
