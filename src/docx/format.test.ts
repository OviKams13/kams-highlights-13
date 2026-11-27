import { describe, expect, it } from 'vitest';
import {
  buildRunViews,
  createListCounter,
  formatNumber,
  parseNumbering,
  parseStyles,
  resolveParagraphProps,
  resolveRunFormat,
  styleChain,
} from './format';
import { flattenParagraph } from './paragraphs';
import { firstParagraph, parseBody, partXml } from './testing';
import { A, W } from './xml';

const STYLES = partXml(
  'styles',
  '<w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="21"/>' +
    '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    // WPS-style numeric ids.
    '<w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/><w:basedOn w:val="1"/>' +
    '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="44"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="1"/>' +
    '<w:pPr><w:jc w:val="center"/><w:ind w:left="400" w:firstLine="200"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="32"/><w:color w:val="FF0000"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Outline"><w:name w:val="My outline"/><w:basedOn w:val="1"/>' +
    '<w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body text"/><w:basedOn w:val="1"/>' +
    '<w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="1"/>' +
    '<w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Loop1"><w:name w:val="Loop 1"/><w:basedOn w:val="Loop2"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Loop2"><w:name w:val="Loop 2"/><w:basedOn w:val="Loop1"/></w:style>' +
    '<w:style w:type="character" w:default="1" w:styleId="11"><w:name w:val="Default Paragraph Font"/></w:style>' +
    '<w:style w:type="character" w:styleId="Link"><w:name w:val="Hyperlink"/><w:basedOn w:val="11"/>' +
    '<w:rPr><w:i/><w:color w:val="0000ff"/><w:u w:val="single"/></w:rPr></w:style>',
);

const THEME = new DOMParser().parseFromString(
  `<a:theme xmlns:a="${A}"><a:themeElements><a:fontScheme name="Office">` +
    '<a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont>' +
    '</a:fontScheme></a:themeElements></a:theme>',
  'application/xml',
);

const styles = parseStyles(STYLES, THEME);

const run = (rPr: string) =>
  firstParagraph(`<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>x</w:t></w:r></w:p>`).getElementsByTagNameNS(
    W,
    'r',
  )[0]!;

describe('parseStyles', () => {
  it('reads styles, defaults and theme fonts', () => {
    expect(styles.defaultParagraphStyleId).toBe('1');
    expect(styles.defaultCharacterStyleId).toBe('11');
    expect(styles.themeFonts).toEqual({ major: 'Calibri Light', minor: 'Calibri' });
    expect(styles.styles.get('2')?.name).toBe('heading 1');
  });

  it('walks basedOn chains base first, and survives a cycle', () => {
    expect(styleChain(styles, 'Body').map((s) => s.id)).toEqual(['1', 'Body']);
    expect(styleChain(styles, 'Loop1').map((s) => s.id)).toEqual(['Loop2', 'Loop1']);
    expect(styleChain(styles, 'missing')).toEqual([]);
  });

  it('copes with a missing styles part', () => {
    expect(parseStyles(undefined).styles.size).toBe(0);
  });
});

describe('resolveRunFormat (direct → character style → paragraph style → docDefaults)', () => {
  it('falls back to docDefaults, with theme fonts', () => {
    expect(resolveRunFormat(styles, '1', run(''))).toEqual({ fontFamily: 'Calibri', sizePt: 10.5 });
  });

  it('applies the paragraph style chain', () => {
    expect(resolveRunFormat(styles, 'Body', run(''))).toEqual({
      fontFamily: 'Times New Roman',
      sizePt: 16,
      color: '#FF0000',
    });
    expect(resolveRunFormat(styles, '2', run(''))).toMatchObject({ bold: true, sizePt: 22 });
  });

  it('applies the character style over the paragraph style', () => {
    expect(resolveRunFormat(styles, 'Body', run('<w:rStyle w:val="Link"/>'))).toEqual({
      fontFamily: 'Times New Roman',
      sizePt: 16,
      italic: true,
      underline: true,
      color: '#0000FF',
    });
  });

  it('applies direct formatting last, including explicit resets', () => {
    const format = resolveRunFormat(
      styles,
      'Body',
      run(
        '<w:rStyle w:val="Link"/><w:rFonts w:ascii="Arial"/><w:i w:val="0"/><w:b/>' +
          '<w:u w:val="none"/><w:color w:val="auto"/><w:sz w:val="40"/><w:highlight w:val="yellow"/>',
      ),
    );
    expect(format).toEqual({ fontFamily: 'Arial', sizePt: 20, bold: true, highlight: 'yellow' });
  });

  it('lets a theme font win over an explicit name on the same layer only', () => {
    expect(
      resolveRunFormat(styles, 'Body', run('<w:rFonts w:ascii="Arial" w:asciiTheme="majorHAnsi"/>'))
        .fontFamily,
    ).toBe('Calibri Light');
    // rFonts with only w:cs (like a WPS run) keeps the inherited ascii font.
    expect(
      resolveRunFormat(styles, 'Body', run('<w:rFonts w:hint="default" w:cs="Arial"/>')).fontFamily,
    ).toBe('Times New Roman');
  });

  it('treats w:highlight none as no highlight', () => {
    expect(resolveRunFormat(styles, '1', run('<w:highlight w:val="none"/>')).highlight).toBe(
      undefined,
    );
  });
});

describe('buildRunViews', () => {
  it('splits the text per run and keeps only non-darkYellow highlights', () => {
    const p = firstParagraph(
      '<w:p>' +
        '<w:r><w:t xml:space="preserve">plain </w:t></w:r>' +
        '<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>yel</w:t></w:r>' +
        '<w:r><w:rPr><w:highlight w:val="darkYellow"/><w:b/></w:rPr><w:t xml:space="preserve"> dark</w:t></w:r>' +
        '</w:p>',
    );
    const { charMap } = flattenParagraph(p);
    expect(buildRunViews(styles, '1', charMap)).toEqual([
      { start: 0, end: 6, fontFamily: 'Calibri', sizePt: 10.5 },
      { start: 6, end: 9, fontFamily: 'Calibri', sizePt: 10.5, originalHighlight: 'yellow' },
      { start: 9, end: 14, fontFamily: 'Calibri', sizePt: 10.5, bold: true },
    ]);
  });
});

const NUMBERING = parseNumbering(
  partXml(
    'numbering',
    '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
      '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
      '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/></w:lvl>' +
      '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="(%3)"/></w:lvl>' +
      '</w:abstractNum>' +
      '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/>' +
      '<w:lvlText w:val=""/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
      '<w:num w:numId="3"><w:abstractNumId w:val="0"/>' +
      '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>',
  ),
);

describe('list markers', () => {
  it('formats numbers', () => {
    expect([1, 4, 9, 14, 40].map((n) => formatNumber(n, 'upperRoman'))).toEqual([
      'I',
      'IV',
      'IX',
      'XIV',
      'XL',
    ]);
    expect([1, 26, 27, 28].map((n) => formatNumber(n, 'lowerLetter'))).toEqual([
      'a',
      'z',
      'aa',
      'bb',
    ]);
    expect(formatNumber(3, 'decimalZero')).toBe('03');
    expect(formatNumber(3, 'none')).toBe('');
    expect(formatNumber(3, 'chineseCounting')).toBe('3');
  });

  it('counts per numId and ilvl in document order, restarting deeper levels', () => {
    const counter = createListCounter(NUMBERING);
    const seq: [string, number][] = [
      ['1', 0],
      ['1', 1],
      ['1', 1],
      ['1', 2],
      ['1', 0],
      ['1', 1],
      ['2', 0],
      ['3', 0],
      ['1', 0],
    ];
    expect(seq.map(([numId, ilvl]) => counter.next(numId, ilvl)?.text)).toEqual([
      '1.',
      '1.a)',
      '1.b)',
      '(i)',
      '2.',
      '2.a)',
      '•',
      '5.',
      '3.',
    ]);
    expect(counter.next('99', 0)).toBeUndefined();
  });
});

describe('resolveParagraphProps', () => {
  const propsOf = (body: string) => {
    const xml = parseBody(body);
    const counter = createListCounter(NUMBERING);
    return Array.from(xml.getElementsByTagNameNS(W, 'p')).map((p) =>
      resolveParagraphProps(styles, NUMBERING, counter, p),
    );
  };

  it('finds headings by style name (numeric ids) or outline level', () => {
    const [h1, outline, bodyText, direct, normal] = propsOf(
      '<w:p><w:pPr><w:pStyle w:val="2"/></w:pPr></w:p>' +
        '<w:p><w:pPr><w:pStyle w:val="Outline"/></w:pPr></w:p>' +
        '<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr></w:p>' +
        '<w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:p>' +
        '<w:p/>',
    );
    expect(h1).toMatchObject({ styleId: '2', headingLevel: 1 });
    expect(outline?.headingLevel).toBe(3);
    expect(bodyText?.headingLevel).toBeUndefined();
    expect(direct?.headingLevel).toBe(2);
    expect(normal).toEqual({ styleId: '1', alignment: 'justify' });
  });

  it('resolves alignment and indentation (twips → pt), direct over style', () => {
    const [style, direct] = propsOf(
      '<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr></w:p>' +
        '<w:p><w:pPr><w:pStyle w:val="Body"/><w:jc w:val="right"/><w:ind w:hanging="100"/></w:pPr></w:p>',
    );
    expect(style).toMatchObject({ alignment: 'center', indent: { leftPt: 20, firstLinePt: 10 } });
    expect(direct).toMatchObject({ alignment: 'right', indent: { leftPt: 20, hangingPt: 5 } });
  });

  it('gives list markers, with the numbering indentation between style and direct', () => {
    const [first, second, styled, none] = propsOf(
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>' +
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="1440"/></w:pPr></w:p>' +
        '<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr></w:p>' +
        '<w:p><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:p>',
    );
    expect(first).toMatchObject({
      list: { numId: '1', ilvl: 0, kind: 'number', text: '1.' },
      indent: { leftPt: 36, hangingPt: 18 },
    });
    expect(second).toMatchObject({
      list: { text: '2.' },
      indent: { leftPt: 72, hangingPt: 18 },
    });
    expect(styled?.list).toEqual({ numId: '2', ilvl: 0, kind: 'bullet', text: '•' });
    expect(none?.list).toBeUndefined();
  });
});
