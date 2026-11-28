import { describe, expect, it } from 'vitest';
import { flattenParagraph } from './paragraphs';
import { firstParagraph } from './testing';
import { applyMarks } from './write';
import { W, wAttr, wChild } from './xml';

const DY = '<w:highlight w:val="darkYellow"/>';
const run = (text: string, rPr = '', attrs = '') =>
  `<w:r${attrs}>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

/** Applies `ranges` to a one-paragraph body. */
function apply(content: string, ranges: [number, number][]): Element {
  const p = firstParagraph(`<w:p>${content}</w:p>`);
  applyMarks(p, ranges);
  return p;
}

const runsOf = (p: Element) => Array.from(p.getElementsByTagNameNS(W, 'r'));
/** [text, highlight] of each run. */
const view = (p: Element) =>
  runsOf(p).map((r) => [
    r.textContent,
    wAttr(wChild(wChild(r, 'rPr'), 'highlight'), 'val') ?? null,
  ]);
const rPrNames = (r: Element) => Array.from(wChild(r, 'rPr')?.children ?? [], (c) => c.nodeName);
const serialize = (node: Node) => new XMLSerializer().serializeToString(node);

describe('applyMarks: splitting runs', () => {
  it('splits a run at the range boundaries', () => {
    const p = apply(run(' from Cyprus International University, '), [[6, 26]]);
    expect(view(p)).toEqual([
      [' from ', null],
      ['Cyprus International', 'darkYellow'],
      [' University, ', null],
    ]);
  });

  it('gives every piece a deep copy of the rPr and the attributes of the run', () => {
    const rPr =
      '<w:rFonts w:ascii="Times New Roman"/><w:sz w:val="32"/><w:lang w:val="en-US"/>' +
      '<w14:textFill xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:noFill/></w14:textFill>';
    const p = firstParagraph(`<w:p>${run('one two three', rPr, ' w:rsidR="00AA"')}</w:p>`);
    const original = serialize(wChild(runsOf(p)[0], 'rPr')!);
    applyMarks(p, [[4, 7]]);

    const runs = runsOf(p);
    expect(runs).toHaveLength(3);
    for (const r of runs) {
      expect(wAttr(r, 'rsidR')).toBe('00AA');
      const copy = wChild(r, 'rPr')!.cloneNode(true) as Element;
      wChild(copy, 'highlight')?.remove();
      expect(serialize(copy)).toBe(original);
    }
  });

  it('keeps the text and adds xml:space="preserve" only where a piece needs it', () => {
    const p = apply(run('one two three'), [[4, 7]]);
    expect(flattenParagraph(p).text).toBe('one two three');
    const texts = Array.from(p.getElementsByTagNameNS(W, 't'));
    expect(texts.map((t) => [t.textContent, t.getAttribute('xml:space')])).toEqual([
      ['one ', 'preserve'],
      ['two', null],
      [' three', 'preserve'],
    ]);
  });

  it('merges adjacent content with the same state into one run', () => {
    const p = apply('<w:r><w:t>ab </w:t><w:t>cd</w:t><w:tab/><w:t>ef</w:t></w:r>', [[0, 5]]);
    expect(view(p)).toEqual([
      ['ab cd', 'darkYellow'],
      ['ef', null],
    ]);
    expect(Array.from(runsOf(p)[1]!.children, (c) => c.localName)).toEqual(['tab', 't']);
  });

  it('splits around tab, br and noBreakHyphen according to their own state', () => {
    const content = '<w:r><w:t>a</w:t><w:noBreakHyphen/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r>';
    const p = apply(content, [[0, 3]]);
    expect(runsOf(p).map((r) => Array.from(r.children, (c) => c.localName))).toEqual([
      ['rPr', 't', 'noBreakHyphen', 't'],
      ['br', 't'],
    ]);
  });

  it('keeps content without characters with the piece before it, or with the first one', () => {
    const content =
      '<w:r><w:lastRenderedPageBreak/><w:t>high</w:t><w:softHyphen/><w:t>light on</w:t></w:r>';
    const p = apply(content, [[0, 9]]);
    expect(runsOf(p).map((r) => Array.from(r.children, (c) => c.localName))).toEqual([
      ['rPr', 'lastRenderedPageBreak', 't', 'softHyphen', 't'],
      ['t'],
    ]);
    expect(view(p)).toEqual([
      ['highlight', 'darkYellow'],
      [' on', null],
    ]);
  });

  it('splits runs inside a hyperlink in place', () => {
    const p = apply(`<w:hyperlink>${run('our website')}</w:hyperlink>${run(' now')}`, [[0, 3]]);
    const link = wChild(p, 'hyperlink')!;
    expect(runsOf(link).map((r) => r.textContent)).toEqual(['our', ' website']);
    expect(flattenParagraph(p).text).toBe('our website now');
  });
});

describe('applyMarks: highlight values', () => {
  it('removes darkYellow, and the rPr that only held it', () => {
    const p = apply(run('word', DY), []);
    expect(view(p)).toEqual([['word', null]]);
    expect(wChild(runsOf(p)[0], 'rPr')).toBeUndefined();
  });

  it('keeps another colour on unwanted pieces and replaces it on wanted ones', () => {
    const p = apply(run('very urgent task', '<w:highlight w:val="yellow"/>'), [[5, 11]]);
    expect(view(p)).toEqual([
      ['very ', 'yellow'],
      ['urgent', 'darkYellow'],
      [' task', 'yellow'],
    ]);
  });

  it('inserts w:highlight in schema order: after sz/szCs, before u, shd, lang and w14:*', () => {
    const rPr =
      '<w:color w:val="000000"/><w:sz w:val="32"/><w:szCs w:val="32"/><w:u w:val="single"/>' +
      '<w:shd w:val="clear"/><w:vertAlign w:val="baseline"/><w:lang w:val="en-US"/>';
    const p = apply(run('word', rPr), [[0, 4]]);
    expect(rPrNames(runsOf(p)[0]!)).toEqual([
      'w:color',
      'w:sz',
      'w:szCs',
      'w:highlight',
      'w:u',
      'w:shd',
      'w:vertAlign',
      'w:lang',
    ]);
  });

  it.each([
    ['<w:lang w:val="en-US"/>', ['w:highlight', 'w:lang']],
    ['<w:eastAsianLayout w:id="1"/>', ['w:highlight', 'w:eastAsianLayout']],
    [
      '<w14:ligatures xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/>',
      ['w:highlight', 'w14:ligatures'],
    ],
    ['<w:b/>', ['w:b', 'w:highlight']],
  ])('inserts w:highlight correctly next to %s', (rPr, expected) => {
    const p = apply(run('word', rPr), [[0, 4]]);
    expect(rPrNames(runsOf(p)[0]!)).toEqual(expected);
  });

  it('creates the rPr as the first child when the run has none', () => {
    const p = apply(run('word'), [[0, 4]]);
    const r = runsOf(p)[0]!;
    expect(r.firstElementChild?.localName).toBe('rPr');
    expect(view(p)).toEqual([['word', 'darkYellow']]);
  });

  it('only touches the direct rPr, not a tracked rPrChange', () => {
    const rPr = `<w:b/>${DY}<w:rPrChange w:id="1"><w:rPr>${DY}</w:rPr></w:rPrChange>`;
    const p = apply(run('word', rPr), []);
    expect(rPrNames(runsOf(p)[0]!)).toEqual(['w:b', 'w:rPrChange']);
    expect(p.getElementsByTagNameNS(W, 'highlight')).toHaveLength(1);
  });

  it('leaves a run that is already right, and a run without text, byte identical', () => {
    const content =
      run('marked', DY) +
      run(' plain ', '<w:b/>') +
      `<w:r><w:rPr>${DY}</w:rPr><w:fldChar w:fldCharType="begin"/></w:r>`;
    const p = firstParagraph(`<w:p><w:pPr><w:rPr>${DY}</w:rPr></w:pPr>${content}</w:p>`);
    const before = serialize(p);
    applyMarks(p, [[0, 6]]);
    expect(serialize(p)).toBe(before);
  });

  it('uses the prefix of the document', () => {
    const xml = new DOMParser().parseFromString(
      `<x:document xmlns:x="${W}"><x:body><x:p><x:r><x:t>one two</x:t></x:r></x:p></x:body></x:document>`,
      'application/xml',
    );
    const p = xml.getElementsByTagNameNS(W, 'p')[0]!;
    applyMarks(p, [[0, 3]]);
    expect(serialize(p)).toContain('<x:rPr><x:highlight x:val="darkYellow"/></x:rPr>');
  });
});
