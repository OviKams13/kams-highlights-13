// @vitest-environment jsdom
// Phase 4.6: the .docx reader on the owner's real WPS files (local only, skipped in CI)
// and on the generated fixtures (always run). jsdom provides DOMParser.
import { describe, expect, it } from 'vitest';
import { openDocx } from '@/docx/open';
import { readDocumentMarks, readParagraphMarks } from '@/docx/read-marks';
import type { DocModel, ParagraphModel } from '@/docx/types';
import { FIXTURES, LARGE_PARAGRAPHS } from './fixtures/make-fixtures.ts';
import {
  describeWithLocalFixture,
  readGeneratedFixture,
  readLocalFixture,
} from './helpers/fixtures.ts';

const wordsOf = (p: ParagraphModel) => p.words.map((w) => p.text.slice(w.start, w.end));
const marksOf = (p: ParagraphModel) => readParagraphMarks(p).units;
const para = (doc: DocModel, index: number) => {
  const p = doc.paragraphs[index];
  if (!p) throw new Error(`no paragraph ${index}`);
  return p;
};

/**
 * Expected values computed by Claude Web with the Python prototype
 * (a.prototype/docx_roundtrip.py). Pn = paragraphIndex, empty paragraphs included.
 */
const TEST_DOCX: Record<number, { words: string[]; marks: number[] }> = {
  1: {
    words: [
      "I'm",
      'a',
      'Software',
      'Engineering',
      'graduate',
      'from',
      'Cyprus',
      'International',
      'University',
      'in',
      'January',
      '2026',
    ],
    marks: [0, 4, 5, 6, 8],
  },
  2: {
    words: [
      'My',
      'main',
      'stack',
      'is',
      'JavaScript',
      'and',
      'TypeScript',
      'with',
      'React',
      'and',
      'Node.js',
    ],
    marks: [0, 2, 4, 6, 8, 12, 14, 16, 20],
  },
  3: {
    words: [
      'On',
      'my',
      'own',
      'time',
      'I',
      'like',
      'to',
      'build',
      'projects',
      'to',
      'really',
      'understand',
      'how',
      'things',
      'work',
      'like',
      'a',
      'full-stack',
      'real',
      'estate',
      'platform',
      'and',
      'a',
      'finance',
      'tracker',
    ],
    marks: [
      0, 1, 2, 4, 5, 6, 8, 9, 10, 14, 16, 22, 24, 26, 27, 28, 30, 34, 36, 37, 38, 40, 46, 47, 48,
    ],
  },
  4: {
    words: [
      'and',
      'right',
      'now',
      "I'm",
      'going',
      'deeper',
      'on',
      'RAG',
      'and',
      'system',
      'design',
    ],
    marks: [4, 6, 8, 9, 10, 12, 14, 18, 19, 20],
  },
  6: {
    words: ['I', 'completed', 'two', 'internships', 'at', 'Near', 'East', 'University'],
    marks: [0, 1, 2, 4, 5, 6, 10, 11, 12, 13, 14],
  },
};

describeWithLocalFixture('test.docx', 'reader on the real WPS file', () => {
  const load = () => openDocx(readLocalFixture('test.docx'), 'test.docx');

  it('has 8 w:p and 5 non-empty paragraphs (1, 2, 3, 4, 6)', async () => {
    const doc = await load();
    expect(doc.paragraphs).toHaveLength(8);
    const nonEmpty = doc.paragraphs.filter((p) => p.words.length > 0).map((p) => p.index);
    expect(nonEmpty).toEqual([1, 2, 3, 4, 6]);
    expect(doc.blocks).toEqual(
      doc.paragraphs.map((p) => ({ kind: 'paragraph', paragraphIndex: p.index })),
    );
  });

  it.each(Object.entries(TEST_DOCX))(
    'P%s: words and marks = the prototype',
    async (index, expected) => {
      const p = para(await load(), Number(index));
      expect(wordsOf(p)).toEqual(expected.words);
      expect(marksOf(p)).toEqual(expected.marks);
    },
  );

  it('reads joined marks such as "Software Engineering" and "real estate"', async () => {
    const doc = await load();
    const marked = (p: ParagraphModel) =>
      readParagraphMarks(p).ranges.map(([start, end]) => p.text.slice(start, end));
    expect(marked(para(doc, 1))).toEqual(["I'm", 'Software Engineering', 'graduate']);
    // Gap 39 (between estate and platform) is not highlighted in the file.
    expect(marked(para(doc, 3))).toEqual(
      expect.arrayContaining(['full-stack', 'real estate', 'platform']),
    );
  });

  it('gives initial marks for the 5 non-empty paragraphs only', async () => {
    const marks = readDocumentMarks(await load());
    expect([...marks.keys()]).toEqual([1, 2, 3, 4, 6]);
  });

  it('resolves Times New Roman 16 and bullet markers', async () => {
    const doc = await load();
    for (const index of [1, 2, 3, 4, 6]) {
      const p = para(doc, index);
      expect(p.props.list).toMatchObject({ kind: 'bullet', text: '•' });
      expect(p.props.indent).toEqual({ leftPt: 21, hangingPt: 21 });
      for (const run of p.runs) expect(run.sizePt).toBe(16);
      expect(p.runs[0]?.fontFamily).toBe('Times New Roman');
      expect(p.runs.every((r) => r.originalHighlight === undefined)).toBe(true);
    }
    expect(para(doc, 5).props.list).toBeUndefined();
  });
});

describeWithLocalFixture('fx-a-headings-list.docx', 'reader on the real WPS headings file', () => {
  const load = () =>
    openDocx(readLocalFixture('fx-a-headings-list.docx'), 'fx-a-headings-list.docx');

  it('reads headings, list items, words and marks', async () => {
    const doc = await load();
    expect(para(doc, 0)).toMatchObject({ text: 'Interview answers', props: { headingLevel: 1 } });
    expect(para(doc, 1)).toMatchObject({ text: 'Question 1', props: { headingLevel: 2 } });
    expect(marksOf(para(doc, 0))).toEqual([]);
    expect(marksOf(para(doc, 1))).toEqual([]);

    expect(wordsOf(para(doc, 2))).toEqual([
      'I',
      'built',
      'a',
      'full-stack',
      'platform',
      'with',
      'React',
      'and',
      'Node.js',
    ]);
    expect(marksOf(para(doc, 2))).toEqual([6, 7, 8]);
    expect(wordsOf(para(doc, 3))).toEqual([
      "I'm",
      'going',
      'deeper',
      'on',
      'RAG',
      'and',
      'system',
      'design',
    ]);
    expect(marksOf(para(doc, 3))).toEqual([8]);
    expect(wordsOf(para(doc, 4))).toEqual(['I', 'completed', 'two', 'internships', 'in', '2025']);
    expect(marksOf(para(doc, 4))).toEqual([]);
  });

  it('puts the three list items in one numbered list: 1. 2. 3.', async () => {
    const doc = await load();
    const lists = [2, 3, 4].map((i) => para(doc, i).props.list);
    expect(new Set(lists.map((l) => l?.numId)).size).toBe(1);
    expect(lists.map((l) => [l?.kind, l?.text])).toEqual([
      ['number', '1.'],
      ['number', '2.'],
      ['number', '3.'],
    ]);
    expect(para(doc, 0).props.list).toBeUndefined();
  });
});

describe('reader on the generated fixtures', () => {
  const load = (name: (typeof FIXTURES)[keyof typeof FIXTURES]) =>
    openDocx(readGeneratedFixture(name), name);

  it('(a) headings, numbered and bulleted lists', async () => {
    const doc = await load(FIXTURES.headingsList);
    expect(doc.paragraphs.map((p) => p.props.headingLevel)).toEqual([
      1,
      undefined,
      2,
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(doc.paragraphs.map((p) => p.props.list?.text)).toEqual([
      undefined,
      undefined,
      undefined,
      '1.',
      '2.',
      '3.',
      '•',
    ]);
    expect(marksOf(para(doc, 1))).toEqual([4, 5, 6, 24]);
    expect(para(doc, 0).runs[0]).toMatchObject({
      fontFamily: 'Times New Roman',
      sizePt: 22,
      bold: true,
    });
  });

  it('(b) table with gridSpan, vMerge and a nested table', async () => {
    const doc = await load(FIXTURES.table);
    expect(doc.paragraphs.map((p) => p.text)).toEqual([
      'Text before the table.',
      'Merged across two columns',
      'Top right',
      'Merged down two rows',
      'Middle cell',
      'Plain cell',
      '',
      'Inner one',
      'Inner two',
      'Inner three',
      'nested cell',
      'After the inner table',
      'Last cell',
      'Text after the table.',
    ]);
    const [before, table, after] = doc.blocks;
    expect(before).toEqual({ kind: 'paragraph', paragraphIndex: 0 });
    expect(after).toEqual({ kind: 'paragraph', paragraphIndex: 13 });
    if (table?.kind !== 'table') throw new Error('expected a table');
    expect(table.rows.map((row) => row.cells.map((c) => [c.gridSpan, c.vMerge]))).toEqual([
      [
        [2, undefined],
        [1, undefined],
      ],
      [
        [1, 'restart'],
        [1, undefined],
        [1, undefined],
      ],
      [
        [1, 'continue'],
        [1, undefined],
        [1, undefined],
      ],
    ]);
    const nestedCell = table.rows[2]?.cells[1]?.blocks;
    expect(nestedCell?.[0]?.kind).toBe('table');
    expect(nestedCell?.[1]).toEqual({ kind: 'paragraph', paragraphIndex: 11 });
    // Marks inside cells, including the nested table.
    expect([...readDocumentMarks(doc).keys()]).toEqual([0, 1, 3, 4, 7, 10, 12, 13]);
  });

  it('(c) hyperlink, tab, br, noBreakHyphen, softHyphen and a word split over two runs', async () => {
    const doc = await load(FIXTURES.inline);
    const p = para(doc, 0);
    expect(p.text).toBe(
      'Visit our website now.\tName:\na well-known Engineering and highlight test.',
    );
    expect(p.hardBreaks).toEqual([22, 28]);
    expect(wordsOf(p)).toEqual([
      'Visit',
      'our',
      'website',
      'now',
      'Name',
      'a',
      'well-known',
      'Engineering',
      'and',
      'highlight',
      'test',
    ]);
    expect(p.gaps.map((g) => g.joinable)).toEqual([
      true,
      true,
      true,
      false,
      false,
      true,
      true,
      true,
      true,
      true,
    ]);
    // "our" (in the hyperlink) and "Engineering" (two runs) are marked.
    expect(marksOf(p)).toEqual([2, 14]);
    const link = p.runs.find((r) => r.start === 6);
    expect(link).toMatchObject({ end: 9, underline: true, color: '#0563C1' });
  });

  it('(d) a yellow word is not a mark, but is shown as originalHighlight', async () => {
    const doc = await load(FIXTURES.otherColour);
    const [p0, p1] = [para(doc, 0), para(doc, 1)];
    expect(marksOf(p0)).toEqual([4, 5, 6]);
    expect(
      p0.runs
        .filter((r) => r.originalHighlight)
        .map((r) => [p0.text.slice(r.start, r.end), r.originalHighlight]),
    ).toEqual([['urgent', 'yellow']]);
    // "yellowdarkYellow" is one word, partly darkYellow: rounded up to the whole word.
    expect(wordsOf(p1)[1]).toBe('yellowdarkYellow');
    expect(marksOf(p1)).toEqual([2]);
  });

  it('(e) ~100 pages open, with every paragraph indexed', { timeout: 60_000 }, async () => {
    const doc = await load(FIXTURES.large);
    // 1400 paragraphs + one Heading 1 every 50.
    expect(doc.paragraphs).toHaveLength(LARGE_PARAGRAPHS + LARGE_PARAGRAPHS / 50);
    expect(doc.blocks).toHaveLength(doc.paragraphs.length);
    const words = doc.paragraphs.reduce((n, p) => n + p.words.length, 0);
    expect(words).toBeGreaterThan(25_000);
    expect(readDocumentMarks(doc).size).toBeGreaterThan(1000);
  });
});
