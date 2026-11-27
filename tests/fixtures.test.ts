// @vitest-environment jsdom
// (jsdom gives us DOMParser to check that every document.xml is well-formed.)
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { buildFixtures, FIXTURES, LARGE_PARAGRAPHS } from './fixtures/make-fixtures.ts';
import {
  describeWithLocalFixture,
  readGeneratedFixture,
  readLocalFixture,
} from './helpers/fixtures.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function loadDocument(bytes: Uint8Array): Promise<XMLDocument> {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file('word/document.xml')?.async('string');
  if (xml === undefined) throw new Error('word/document.xml is missing');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error('word/document.xml is not well-formed');
  }
  return doc;
}

const all = (root: Document | Element, local: string) =>
  Array.from(root.getElementsByTagNameNS(W, local));

/** Highlight colour of each w:r that has one (paragraph-mark rPr excluded). */
function runHighlights(doc: XMLDocument): string[] {
  return all(doc, 'r').flatMap((run) => {
    const rPr = Array.from(run.children).find((c) => c.localName === 'rPr');
    const hl = rPr && Array.from(rPr.children).find((c) => c.localName === 'highlight');
    const val = hl?.getAttributeNS(W, 'val');
    return val ? [val] : [];
  });
}

/** Texts of each w:r, in order. */
const runTexts = (doc: XMLDocument) =>
  all(doc, 'r').map((run) =>
    all(run, 't')
      .map((t) => t.textContent)
      .join(''),
  );

describe('generated fixtures', () => {
  const names = Object.values(FIXTURES);

  it.each(names)('%s is a valid package in Times New Roman 16 with darkYellow', async (name) => {
    const bytes = readGeneratedFixture(name);
    const zip = await JSZip.loadAsync(bytes);
    for (const part of [
      '[Content_Types].xml',
      '_rels/.rels',
      'word/document.xml',
      'word/_rels/document.xml.rels',
      'word/styles.xml',
      'word/numbering.xml',
    ]) {
      expect(zip.file(part), part).not.toBeNull();
    }
    const doc = await loadDocument(bytes);
    const fonts = all(doc, 'rFonts').map((f) => f.getAttributeNS(W, 'ascii'));
    expect(new Set(fonts)).toEqual(new Set(['Times New Roman']));
    expect(all(doc, 'sz').some((s) => s.getAttributeNS(W, 'val') === '32')).toBe(true);
    expect(runHighlights(doc)).toContain('darkYellow');
  });

  it('is deterministic (same bytes every time)', async () => {
    const [first, second] = await Promise.all([buildFixtures(), buildFixtures()]);
    for (const name of names) {
      expect(Buffer.compare(first[name], second[name]), name).toBe(0);
    }
  });

  it('(a) has Heading 1, Heading 2 and a numbered list', async () => {
    const doc = await loadDocument(readGeneratedFixture(FIXTURES.headingsList));
    const styles = all(doc, 'pStyle').map((s) => s.getAttributeNS(W, 'val'));
    expect(styles).toEqual(expect.arrayContaining(['Heading1', 'Heading2']));
    const numIds = all(doc, 'numId').map((n) => n.getAttributeNS(W, 'val'));
    expect(numIds.filter((id) => id === '1')).toHaveLength(3);
    expect(runTexts(doc)).toContain('full-stack platform');
  });

  it('(b) has gridSpan, vMerge restart/continue and a nested table', async () => {
    const doc = await loadDocument(readGeneratedFixture(FIXTURES.table));
    expect(all(doc, 'gridSpan')).toHaveLength(1);
    const vMerges = all(doc, 'vMerge').map((v) => v.getAttributeNS(W, 'val'));
    // A bare <w:vMerge/> (no w:val) means "continue".
    expect(vMerges).toEqual(['restart', null]);
    const tables = all(doc, 'tbl');
    expect(tables).toHaveLength(2);
    expect(tables.some((t) => t.parentElement?.closest('tbl') !== null)).toBe(true);
  });

  it('(c) has a hyperlink, tab, br, noBreakHyphen, softHyphen and a word split over two runs', async () => {
    const doc = await loadDocument(readGeneratedFixture(FIXTURES.inline));
    expect(all(doc, 'p')).toHaveLength(1);
    for (const local of ['hyperlink', 'tab', 'br', 'noBreakHyphen', 'softHyphen']) {
      expect(all(doc, local), local).toHaveLength(1);
    }
    const texts = runTexts(doc);
    const i = texts.indexOf('Engin');
    expect(texts[i + 1]).toBe('eering');
  });

  it('(d) has a yellow word next to darkYellow words', async () => {
    const doc = await loadDocument(readGeneratedFixture(FIXTURES.otherColour));
    const hl = runHighlights(doc);
    expect(hl).toContain('yellow');
    expect(hl).toContain('darkYellow');
  });

  it('(e) is about 100 pages long', async () => {
    const doc = await loadDocument(readGeneratedFixture(FIXTURES.large));
    expect(all(doc, 'p').length).toBeGreaterThanOrEqual(LARGE_PARAGRAPHS);
    const words = (doc.documentElement.textContent ?? '').split(/\s+/).filter(Boolean).length;
    // ~280 words per A4 page at 16 pt.
    expect(words).toBeGreaterThan(25_000);
  });
});

describeWithLocalFixture('test.docx', "owner's real WPS file", () => {
  it('opens and has 34 darkYellow runs', async () => {
    const doc = await loadDocument(readLocalFixture('test.docx'));
    expect(runHighlights(doc).filter((h) => h === 'darkYellow')).toHaveLength(34);
  });
});

describeWithLocalFixture('fx-a-headings-list.docx', "owner's real WPS headings file", () => {
  it('opens and has 2 darkYellow runs', async () => {
    const doc = await loadDocument(readLocalFixture('fx-a-headings-list.docx'));
    expect(runHighlights(doc).filter((h) => h === 'darkYellow')).toHaveLength(2);
  });
});
