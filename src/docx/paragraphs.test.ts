import { describe, expect, it } from 'vitest';
import { collectParagraphs, flattenParagraph, paragraphRuns } from './paragraphs';
import { firstParagraph, parseBody } from './testing';

const r = (inner: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${inner}</w:r>`;
const t = (text: string) => `<w:t xml:space="preserve">${text}</w:t>`;

describe('collectParagraphs', () => {
  it('returns every w:p in document order, at every depth, text boxes excluded', () => {
    const textBox =
      '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:txbx><w:txbxContent>' +
      `<w:p>${r(t('in a text box'))}</w:p>` +
      '</w:txbxContent></wps:txbx></w:drawing></mc:Choice></mc:AlternateContent>';
    const xml = parseBody(
      `<w:p>${r(t('zero'))}</w:p>` +
        `<w:tbl><w:tr><w:tc><w:p>${r(t('one'))}</w:p>` +
        `<w:tbl><w:tr><w:tc><w:p>${r(t('two'))}</w:p></w:tc></w:tr></w:tbl>` +
        `<w:p>${r(textBox)}${r(t('three'))}</w:p></w:tc></w:tr></w:tbl>` +
        '<w:p/>',
    );
    const texts = collectParagraphs(xml).map((p) => flattenParagraph(p).text);
    expect(texts).toEqual(['zero', 'one', 'two', 'three', '']);
  });
});

describe('flattenParagraph', () => {
  it('reads runs inside hyperlink, ins, smartTag, sdtContent and fldSimple', () => {
    const p = firstParagraph(
      '<w:p>' +
        r(t('a ')) +
        `<w:hyperlink r:id="rId1">${r(t('b '))}</w:hyperlink>` +
        `<w:ins w:id="1" w:author="x">${r(t('c '))}</w:ins>` +
        `<w:smartTag w:uri="u" w:element="e">${r(t('d '))}</w:smartTag>` +
        `<w:sdt><w:sdtPr/><w:sdtContent>${r(t('e '))}</w:sdtContent></w:sdt>` +
        `<w:fldSimple w:instr="PAGE">${r(t('f'))}</w:fldSimple>` +
        '</w:p>',
    );
    expect(flattenParagraph(p).text).toBe('a b c d e f');
  });

  it('skips delText, instrText, softHyphen, moved-away text and mc:Fallback copies', () => {
    const p = firstParagraph(
      '<w:p>' +
        r(t('keep')) +
        `<w:del w:id="1" w:author="x">${r('<w:delText>gone</w:delText>')}</w:del>` +
        r('<w:fldChar w:fldCharType="begin"/>') +
        r('<w:instrText> PAGE </w:instrText>') +
        r('<w:fldChar w:fldCharType="separate"/>') +
        r(t('1')) +
        r('<w:fldChar w:fldCharType="end"/>') +
        r(t('high') + '<w:softHyphen/>' + t('light')) +
        `<w:moveFrom w:id="2" w:author="x">${r(t('moved'))}</w:moveFrom>` +
        `<mc:AlternateContent><mc:Choice Requires="wps">${r(t('X'))}</mc:Choice>` +
        `<mc:Fallback>${r(t('X'))}</mc:Fallback></mc:AlternateContent>` +
        '</w:p>',
    );
    expect(flattenParagraph(p).text).toBe('keep1highlightX');
  });

  it('turns tab / br / cr into hard breaks and noBreakHyphen into "-"', () => {
    const p = firstParagraph(
      '<w:p>' +
        r(t('a') + '<w:tab/>' + t('b') + '<w:br/>' + t('c') + '<w:cr/>' + t('well')) +
        r('<w:noBreakHyphen/>') +
        r(t('known')) +
        '</w:p>',
    );
    const flat = flattenParagraph(p);
    expect(flat.text).toBe('a\tb\nc\nwell-known');
    expect(flat.hardBreaks).toEqual([1, 3, 5]);
  });

  it('maps every character to its run, node and offset', () => {
    const p = firstParagraph('<w:p>' + r(t('ab')) + r('<w:tab/>' + t('c')) + '</w:p>');
    const runs = paragraphRuns(p);
    const flat = flattenParagraph(p);
    expect(flat.runs).toEqual(runs);
    expect(flat.charMap).toHaveLength(flat.text.length);
    expect(flat.charMap.map((c) => [runs.indexOf(c.run), c.node.localName, c.offset])).toEqual([
      [0, 't', 0],
      [0, 't', 1],
      [1, 'tab', null],
      [1, 't', 0],
    ]);
  });

  it('keeps a word split over two runs (WPS style) as one piece of text', () => {
    const p = firstParagraph('<w:p>' + r(t('Engin')) + r(t('eering')) + r(t(' rocks')) + '</w:p>');
    const flat = flattenParagraph(p);
    expect(flat.text).toBe('Engineering rocks');
    expect(flat.charMap[4]!.run).not.toBe(flat.charMap[5]!.run);
  });

  it('counts UTF-16 units, like the tokenizer', () => {
    const p = firstParagraph('<w:p>' + r(t('a 😀 b')) + '</w:p>');
    const flat = flattenParagraph(p);
    expect(flat.text).toBe('a 😀 b');
    expect(flat.charMap).toHaveLength(6);
  });

  it('ignores runs of a paragraph nested in a text box', () => {
    const p = firstParagraph(
      '<w:p>' +
        r(t('outer ')) +
        r(
          '<w:drawing><wps:txbx><w:txbxContent>' +
            `<w:p>${r(t('inner'))}</w:p>` +
            '</w:txbxContent></wps:txbx></w:drawing>',
        ) +
        r(t('text')) +
        '</w:p>',
    );
    expect(flattenParagraph(p).text).toBe('outer text');
  });

  it('gives an empty result for an empty paragraph', () => {
    expect(flattenParagraph(firstParagraph('<w:p/>'))).toEqual({
      text: '',
      charMap: [],
      hardBreaks: [],
      runs: [],
    });
  });
});
