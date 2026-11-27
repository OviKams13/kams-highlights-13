import { describe, expect, it } from 'vitest';
import { tokenize } from '@/core/tokenize';
import { flattenParagraph } from './paragraphs';
import {
  darkYellowRanges,
  isDarkYellowRun,
  readDocumentMarks,
  readParagraphMarks,
} from './read-marks';
import { firstParagraph } from './testing';

const plain = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const hl = (text: string, colour = 'darkYellow') =>
  `<w:r><w:rPr><w:sz w:val="32"/><w:highlight w:val="${colour}"/><w:lang w:val="en-US"/></w:rPr>` +
  `<w:t xml:space="preserve">${text}</w:t></w:r>`;

/** Flattens + tokenizes one paragraph built from run XML. */
function paragraph(runs: string) {
  const flat = flattenParagraph(firstParagraph(`<w:p>${runs}</w:p>`));
  return { ...flat, ...tokenize(flat.text, flat.hardBreaks) };
}

describe('darkYellow → char ranges', () => {
  it('reads only the direct darkYellow highlight of each run', () => {
    const p = paragraph(plain('a ') + hl('bc') + hl('d', 'yellow') + hl(' e'));
    expect(darkYellowRanges(p.charMap)).toEqual([
      [2, 4],
      [5, 7],
    ]);
    expect(isDarkYellowRun(p.charMap[0]!.run)).toBe(false);
  });
});

describe('readParagraphMarks', () => {
  it('reads whole words and joined gaps, like the prototype', () => {
    const p = paragraph(hl("I'm") + plain(' a ') + hl('Software Engineering') + plain(' graduate'));
    expect(readParagraphMarks(p)).toEqual({
      units: [0, 4, 5, 6],
      ranges: [
        [0, 3],
        [6, 26],
      ],
    });
  });

  it('reads a word split over two highlighted runs as one mark', () => {
    const p = paragraph(plain('the ') + hl('Engin') + hl('eering') + plain(' team'));
    expect(readParagraphMarks(p).units).toEqual([2]);
  });

  it('rounds a partial highlight up to the whole word (decision 2026-11-27)', () => {
    const p = paragraph(plain('from ') + hl('Cyprus') + plain("'s campus"));
    expect(readParagraphMarks(p)).toEqual({ units: [2], ranges: [[5, 13]] });
  });

  it('drops a WPS trailing-space highlight', () => {
    const p = paragraph(hl('word ') + plain('next'));
    expect(readParagraphMarks(p)).toEqual({ units: [0], ranges: [[0, 4]] });
  });

  it('never joins over a tab or a line break', () => {
    const p = paragraph(
      hl('one') + '<w:r><w:rPr><w:highlight w:val="darkYellow"/></w:rPr><w:tab/></w:r>' + hl('two'),
    );
    expect(p.text).toBe('one\ttwo');
    expect(readParagraphMarks(p).units).toEqual([0, 2]);
  });

  it('ignores other highlight colours', () => {
    const p = paragraph(hl('urgent', 'yellow') + plain(' ') + hl('deadline matters'));
    expect(readParagraphMarks(p).units).toEqual([2, 3, 4]);
  });
});

describe('readDocumentMarks', () => {
  it('keeps only the paragraphs that have marks, keyed by paragraphIndex', () => {
    const paragraphs = [
      paragraph(plain('none here')),
      paragraph(''),
      paragraph(hl('yes') + plain(' no')),
    ].map((p, index) => ({ ...p, index }));
    const marks = readDocumentMarks({ paragraphs });
    expect([...marks.keys()]).toEqual([2]);
    expect(marks.get(2)).toEqual({ units: [0], ranges: [[0, 3]] });
  });
});
