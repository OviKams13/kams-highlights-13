/**
 * Tap rules of the Vokamsbulary 13 pencil (15.1), per paragraph (D-04, D-06).
 *
 * `normalizeMarks`, `canMarkSpace` and `toggleMark` are copied verbatim from
 * a.handoff-highlight.md §1.2 (`src/lib/review.ts` lines 76–151 at 838bd7c).
 * The only change (phase 3.1) is an optional `joinable(unit)` predicate: a gap
 * that contains a tab or a line break (D-05) can never be marked. Without the
 * predicate every gap is joinable, which is exactly the hand-off behaviour.
 *
 * Numbering: word i is unit 2i, the gap after it is unit 2i + 1.
 */
import type { Gap } from './tokenize';

/** Says whether gap unit `unit` (an odd number) may be marked. */
export type Joinable = (unit: number) => boolean;

const ALWAYS_JOINABLE: Joinable = () => true;

/** The `joinable` predicate of a tokenized paragraph: gap i is unit 2i + 1. */
export function joinableFromGaps(gaps: readonly Gap[]): Joinable {
  return (unit) => unit % 2 === 1 && (gaps[(unit - 1) / 2]?.joinable ?? false);
}

/**
 * Valid, sorted highlights for a definition of `wordCount` words: whole
 * numbers inside the definition, no repeats, and a space only when both of
 * its words are highlighted. Used on the device and again on the server.
 */
export function normalizeMarks(
  marks: readonly number[],
  wordCount: number,
  joinable: Joinable = ALWAYS_JOINABLE,
): number[] {
  const last = 2 * wordCount - 2;
  const set = new Set(marks.filter((unit) => Number.isInteger(unit) && unit >= 0 && unit <= last));
  for (const unit of set) {
    if (unit % 2 === 1 && (!joinable(unit) || !set.has(unit - 1) || !set.has(unit + 1)))
      set.delete(unit);
  }
  return [...set].sort((a, b) => a - b);
}

/** A space can be highlighted only when both words beside it are. */
export function canMarkSpace(
  marks: ReadonlySet<number>,
  unit: number,
  joinable: Joinable = ALWAYS_JOINABLE,
): boolean {
  return joinable(unit) && marks.has(unit - 1) && marks.has(unit + 1);
}

/**
 * One tap on a unit. A word switches on or off; switching it off also removes
 * the spaces beside it. A space switches only when both its words are on.
 * Words are never joined by themselves.
 */
export function toggleMark(
  marks: readonly number[],
  unit: number,
  wordCount: number,
  joinable: Joinable = ALWAYS_JOINABLE,
): number[] {
  const set = new Set(marks);
  if (set.has(unit)) {
    set.delete(unit);
    if (unit % 2 === 0) {
      set.delete(unit - 1);
      set.delete(unit + 1);
    }
  } else if (unit % 2 === 0 || canMarkSpace(set, unit, joinable)) {
    set.add(unit);
  }
  return normalizeMarks([...set], wordCount, joinable);
}
