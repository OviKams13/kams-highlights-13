import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { mergeRanges, rangesFromUnits, unitsFromRanges } from './marks';
import { tokenize } from './tokenize';
import { joinableFromGaps, normalizeMarks } from './units';

describe('mergeRanges', () => {
  it('sorts, merges overlapping and touching ranges', () => {
    expect(
      mergeRanges([
        [10, 12],
        [0, 3],
        [3, 5],
        [2, 4],
        [11, 15],
      ]),
    ).toEqual([
      [0, 5],
      [10, 15],
    ]);
  });

  it('drops empty and invalid ranges', () => {
    expect(
      mergeRanges([
        [3, 3],
        [5, 2],
        [-1, 4],
        [1.5, 3],
        [7, 9],
      ]),
    ).toEqual([[7, 9]]);
    expect(mergeRanges([])).toEqual([]);
  });

  it('does not change its input', () => {
    const input: [number, number][] = [
      [4, 6],
      [0, 5],
    ];
    mergeRanges(input);
    expect(input).toEqual([
      [4, 6],
      [0, 5],
    ]);
  });
});

// "Extremely pleased and excited." → Extremely [0,9) ␣ pleased [10,17) ␣ and [18,21) ␣ excited [22,29)
const TEXT = 'Extremely pleased and excited.';
const { words, gaps } = tokenize(TEXT);

describe('rangesFromUnits', () => {
  it('gives one range per word, and one range for words joined by their gap', () => {
    expect(rangesFromUnits([2, 4], words, gaps)).toEqual([
      [10, 17],
      [18, 21],
    ]);
    expect(rangesFromUnits([2, 3, 4], words, gaps)).toEqual([[10, 21]]);
  });

  it('never highlights the trailing punctuation', () => {
    expect(rangesFromUnits([6], words, gaps)).toEqual([[22, 29]]);
  });

  it('normalizes first: a lone gap or a unit out of range gives nothing', () => {
    expect(rangesFromUnits([3, 99, -1], words, gaps)).toEqual([]);
  });

  it('never highlights a gap that is not joinable', () => {
    const text = 'Name:\tKams';
    const tokens = tokenize(text, [5]);
    expect(rangesFromUnits([0, 1, 2], tokens.words, tokens.gaps)).toEqual([
      [0, 4],
      [6, 10],
    ]);
  });

  it('highlights the whole gap, punctuation included ("a — b")', () => {
    const tokens = tokenize('a — b');
    expect(rangesFromUnits([0, 1, 2], tokens.words, tokens.gaps)).toEqual([[0, 5]]);
  });
});

describe('unitsFromRanges', () => {
  it('marks a word when its whole span is covered', () => {
    expect(unitsFromRanges([[10, 17]], words, gaps)).toEqual([2]);
    expect(unitsFromRanges([[9, 18]], words, gaps)).toEqual([2]);
  });

  it('rounds a partial highlight up to the whole word (decision 2026-11-27)', () => {
    expect(unitsFromRanges([[10, 16]], words, gaps)).toEqual([2]);
    expect(unitsFromRanges([[12, 13]], words, gaps)).toEqual([2]);
    expect(unitsFromRanges([[16, 17]], words, gaps)).toEqual([2]);
    // Touching only the gap next to a word does not mark it.
    expect(unitsFromRanges([[9, 10]], words, gaps)).toEqual([]);
    expect(unitsFromRanges([[17, 18]], words, gaps)).toEqual([]);
  });

  it('a highlight on "Cyprus" only of "Cyprus\'s" marks the whole word', () => {
    const text = "from Cyprus's university";
    const tokens = tokenize(text);
    expect(tokens.words.map(({ start, end }) => text.slice(start, end))).toEqual([
      'from',
      "Cyprus's",
      'university',
    ]);
    const units = unitsFromRanges([[5, 11]], tokens.words, tokens.gaps);
    expect(units).toEqual([2]);
    // On export, the whole word is highlighted.
    expect(rangesFromUnits(units, tokens.words, tokens.gaps)).toEqual([[5, 13]]);
  });

  it('"word " highlighted with its trailing space, then an unmarked word, gives the word only', () => {
    // WPS double-click highlights the word plus the space after it.
    const text = 'word next';
    const tokens = tokenize(text);
    const units = unitsFromRanges([[0, 5]], tokens.words, tokens.gaps);
    expect(units).toEqual([0]);
    // On export, the highlighted trailing space is dropped.
    expect(rangesFromUnits(units, tokens.words, tokens.gaps)).toEqual([[0, 4]]);
  });

  it('does not mark a gap that is only partly covered, even with both words marked', () => {
    const tokens = tokenize('a — b');
    expect(
      unitsFromRanges(
        [
          [0, 2],
          [4, 5],
        ],
        tokens.words,
        tokens.gaps,
      ),
    ).toEqual([0, 2]);
    expect(unitsFromRanges([[0, 5]], tokens.words, tokens.gaps)).toEqual([0, 1, 2]);
  });

  it('accepts a word covered by several touching ranges', () => {
    expect(
      unitsFromRanges(
        [
          [10, 13],
          [13, 17],
        ],
        words,
        gaps,
      ),
    ).toEqual([2]);
  });

  it('marks a gap when both words are marked and the whole gap is covered', () => {
    expect(unitsFromRanges([[10, 21]], words, gaps)).toEqual([2, 3, 4]);
    expect(
      unitsFromRanges(
        [
          [10, 17],
          [18, 21],
        ],
        words,
        gaps,
      ),
    ).toEqual([2, 4]);
  });

  it('does not mark a covered gap when one of its words is not touched', () => {
    // [17, 21) = the gap after "pleased" + "and": "pleased" has no highlighted character.
    expect(unitsFromRanges([[17, 21]], words, gaps)).toEqual([4]);
  });

  it('a range starting inside a word and covering the gap joins both words', () => {
    expect(unitsFromRanges([[12, 21]], words, gaps)).toEqual([2, 3, 4]);
  });

  it('does not mark a covered gap that is not joinable', () => {
    const text = 'Name:\tKams';
    const tokens = tokenize(text, [5]);
    expect(unitsFromRanges([[0, 10]], tokens.words, tokens.gaps)).toEqual([0, 2]);
  });

  it('a highlight covering the trailing punctuation still marks the word', () => {
    expect(unitsFromRanges([[22, 30]], words, gaps)).toEqual([6]);
  });

  it('returns [] for a paragraph without words', () => {
    expect(unitsFromRanges([[0, 3]], [], [])).toEqual([]);
  });
});

describe('ranges ↔ units (property, ERD rule)', () => {
  const char = fc.constantFrom('a', 'b', '7', 'é', ' ', ' ', '\t', '\n', ' ', '.', ',', '—');
  const paragraph = fc.array(char, { maxLength: 40 }).map((chars) => {
    const text = chars.join('');
    const hard = [...text.matchAll(/[\t\n]/g)].map((match) => match.index);
    return { text, ...tokenize(text, hard) };
  });
  const scenario = paragraph.chain((p) =>
    fc.tuple(
      fc.constant(p),
      fc.array(fc.integer({ min: -2, max: 2 * p.words.length + 1 }), { maxLength: 30 }),
    ),
  );

  it('unitsFromRanges(rangesFromUnits(u)) = normalizeMarks(u)', () => {
    fc.assert(
      fc.property(scenario, ([{ words, gaps }, units]) => {
        const expected = normalizeMarks(units, words.length, joinableFromGaps(gaps));
        expect(unitsFromRanges(rangesFromUnits(units, words, gaps), words, gaps)).toEqual(expected);
      }),
      { numRuns: 500 },
    );
  });

  it('any range marks exactly the words it touches, and only fully covered gaps', () => {
    const withRanges = paragraph.chain((p) =>
      fc.tuple(
        fc.constant(p),
        fc.array(
          fc
            .tuple(
              fc.integer({ min: 0, max: p.text.length }),
              fc.integer({ min: 0, max: p.text.length }),
            )
            .map(([a, b]): [number, number] => [Math.min(a, b), Math.max(a, b)]),
          { maxLength: 6 },
        ),
      ),
    );
    fc.assert(
      fc.property(withRanges, ([{ words, gaps }, ranges]) => {
        const hit = (k: number) => ranges.some(([start, end]) => k >= start && k < end);
        const units = new Set(unitsFromRanges(ranges, words, gaps));
        words.forEach((word, i) => {
          let touched = false;
          for (let k = word.start; k < word.end; k += 1) touched ||= hit(k);
          expect(units.has(2 * i)).toBe(touched);
        });
        gaps.forEach((gap, i) => {
          let full = true;
          for (let k = gap.start; k < gap.end; k += 1) full &&= hit(k);
          const expected = gap.joinable && full && units.has(2 * i) && units.has(2 * i + 2);
          expect(units.has(2 * i + 1)).toBe(expected);
        });
      }),
      { numRuns: 500 },
    );
  });

  it('rangesFromUnits gives sorted, merged ranges that never touch', () => {
    fc.assert(
      fc.property(scenario, ([{ words, gaps }, units]) => {
        const ranges = rangesFromUnits(units, words, gaps);
        expect(mergeRanges(ranges)).toEqual(ranges);
        ranges.forEach(([start, end], i) => {
          expect(end).toBeGreaterThan(start);
          const next = ranges[i + 1];
          if (next) expect(next[0]).toBeGreaterThan(end);
        });
      }),
    );
  });

  it('mergeRanges covers exactly the same characters as its input', () => {
    const range = fc
      .tuple(fc.integer({ min: 0, max: 30 }), fc.integer({ min: 0, max: 30 }))
      .map(([a, b]): [number, number] => [Math.min(a, b), Math.max(a, b)]);
    fc.assert(
      fc.property(fc.array(range, { maxLength: 10 }), (ranges) => {
        const chars = (list: readonly [number, number][]) => {
          const set = new Set<number>();
          for (const [start, end] of list) for (let k = start; k < end; k += 1) set.add(k);
          return [...set].sort((x, y) => x - y);
        };
        expect(chars(mergeRanges(ranges))).toEqual(chars(ranges));
      }),
    );
  });
});
