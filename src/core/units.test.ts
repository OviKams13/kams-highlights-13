import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { tokenize } from './tokenize';
import { canMarkSpace, joinableFromGaps, normalizeMarks, toggleMark } from './units';

// The example of the hand-off §2: "Extremely pleased and excited."
// Extremely=0, ␣=1, pleased=2, ␣=3, and=4, ␣=5, excited=6.
const WORDS = 4;

/** Plays taps in order, from no marks. */
function play(taps: number[], wordCount = WORDS, joinable?: (unit: number) => boolean) {
  return taps.reduce<number[]>((marks, unit) => toggleMark(marks, unit, wordCount, joinable), []);
}

describe('tap rules (hand-off §4, 15.1 scenarios)', () => {
  it('a tap on a word highlights it; a second tap removes it', () => {
    expect(play([2])).toEqual([2]);
    expect(play([2, 2])).toEqual([]);
  });

  it('never joins two highlighted words by itself', () => {
    expect(play([2, 4])).toEqual([2, 4]);
  });

  it('a gap is tappable only when both of its words are highlighted', () => {
    expect(play([3])).toEqual([]);
    expect(play([2, 3])).toEqual([2]);
    expect(play([4, 3])).toEqual([4]);
    expect(canMarkSpace(new Set([2]), 3)).toBe(false);
    expect(canMarkSpace(new Set([2, 4]), 3)).toBe(true);
  });

  it('a tap on a tappable gap joins the two words ("pleased and" = {2,3,4})', () => {
    expect(play([2, 4, 3])).toEqual([2, 3, 4]);
  });

  it('a re-tap on a joined gap splits them, both words stay highlighted', () => {
    expect(play([2, 4, 3, 3])).toEqual([2, 4]);
  });

  it('un-highlighting a word also removes the gaps on both sides of it', () => {
    const joined = play([0, 2, 4, 1, 3]);
    expect(joined).toEqual([0, 1, 2, 3, 4]);
    expect(toggleMark(joined, 2, WORDS)).toEqual([0, 4]);
  });

  it('a gap next to a word just removed cannot be tapped any more', () => {
    const afterRemoval = play([0, 2, 4, 1, 3, 2]);
    expect(afterRemoval).toEqual([0, 4]);
    expect(toggleMark(afterRemoval, 1, WORDS)).toEqual([0, 4]);
    expect(toggleMark(afterRemoval, 3, WORDS)).toEqual([0, 4]);
  });

  it('removing the first or last word of a joined run keeps the rest joined', () => {
    const run = play([0, 2, 4, 1, 3]);
    expect(toggleMark(run, 0, WORDS)).toEqual([2, 3, 4]);
    expect(toggleMark(run, 4, WORDS)).toEqual([0, 1, 2]);
  });

  it('ignores units outside the paragraph (leading/trailing text has no unit)', () => {
    expect(play([-1, -2, 7, 8, 2.5])).toEqual([]);
  });
});

describe('normalizeMarks (verbatim from the hand-off)', () => {
  it('keeps whole numbers in range, removes repeats and sorts', () => {
    expect(normalizeMarks([6, 0, 2, 2, -2, 7, 1.5, Number.NaN], WORDS)).toEqual([0, 2, 6]);
  });

  it('drops a gap that lacks one of its words', () => {
    expect(normalizeMarks([1, 2, 3, 4], WORDS)).toEqual([2, 3, 4]);
    expect(normalizeMarks([0, 1], WORDS)).toEqual([0]);
  });

  it('returns [] when there are no words', () => {
    expect(normalizeMarks([0, 1, 2], 0)).toEqual([]);
  });

  it('without a predicate, behaves exactly like the hand-off (every gap joinable)', () => {
    expect(normalizeMarks([0, 1, 2], 2)).toEqual([0, 1, 2]);
  });
});

describe('joinable predicate (phase 3.1, D-05)', () => {
  // "Name:\tKams Highlights": the gap ":\t" holds a tab and cannot be joined.
  const text = 'Name:\tKams Highlights';
  const { words, gaps } = tokenize(text, [5]);
  const joinable = joinableFromGaps(gaps);

  it('reads the predicate from the gaps', () => {
    expect(words).toHaveLength(3);
    expect([1, 3].map(joinable)).toEqual([false, true]);
    expect([0, 2, 5, -1].map(joinable)).toEqual([false, false, false, false]);
  });

  it('a gap with a hard break is never tappable, even with both words highlighted', () => {
    const marks = play([0, 2], words.length, joinable);
    expect(marks).toEqual([0, 2]);
    expect(canMarkSpace(new Set(marks), 1, joinable)).toBe(false);
    expect(toggleMark(marks, 1, words.length, joinable)).toEqual([0, 2]);
  });

  it('the other gaps still join', () => {
    expect(play([2, 4, 3], words.length, joinable)).toEqual([2, 3, 4]);
  });

  it('normalizeMarks drops a non-joinable gap', () => {
    expect(normalizeMarks([0, 1, 2], words.length, joinable)).toEqual([0, 2]);
  });

  it('words are never affected by the predicate', () => {
    expect(play([0, 2, 4], words.length, () => false)).toEqual([0, 2, 4]);
  });
});

describe('tap rules (properties)', () => {
  const scenario = fc
    .integer({ min: 0, max: 8 })
    .chain((wordCount) =>
      fc.tuple(
        fc.constant(wordCount),
        fc.array(fc.boolean(), { minLength: wordCount, maxLength: wordCount }),
        fc.array(fc.integer({ min: -2, max: 2 * wordCount + 1 }), { maxLength: 30 }),
      ),
    );

  it('every tap gives normalized marks, and only changes the tapped unit or its gaps', () => {
    fc.assert(
      fc.property(scenario, ([wordCount, joinFlags, taps]) => {
        const joinable = (unit: number) => joinFlags[(unit - 1) / 2] === true;
        let marks: number[] = [];
        for (const unit of taps) {
          const next = toggleMark(marks, unit, wordCount, joinable);
          expect(normalizeMarks(next, wordCount, joinable)).toEqual(next);
          const changed = [
            ...next.filter((u) => !marks.includes(u)),
            ...marks.filter((u) => !next.includes(u)),
          ];
          for (const u of changed) expect(Math.abs(u - unit)).toBeLessThanOrEqual(1);
          // Never joins by itself: a word tap never adds a gap.
          if (unit % 2 === 0) {
            expect(next.filter((u) => u % 2 === 1 && !marks.includes(u))).toEqual([]);
          }
          marks = next;
        }
      }),
    );
  });
});
