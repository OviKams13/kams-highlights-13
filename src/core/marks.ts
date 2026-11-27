/**
 * Ranges ↔ units (D-07, ERD "Ranges ↔ units").
 *
 * Highlights are stored as character ranges per paragraph, the same form as
 * the highlights already in a .docx. The tap rules work on units (word i = 2i,
 * gap i = 2i + 1). These pure functions convert between the two.
 */
import type { Gap, Span } from './tokenize';
import { joinableFromGaps, normalizeMarks } from './units';

/** A highlighted character range [start, end) in the flattened paragraph text. */
export type CharRange = [start: number, end: number];

/**
 * Sorted, merged, non-overlapping ranges. Ranges that overlap or touch
 * (`[0, 3]` and `[3, 5]`) become one; empty or invalid ranges are dropped.
 */
export function mergeRanges(ranges: readonly (readonly [number, number])[]): CharRange[] {
  const sorted = ranges
    .filter(
      ([start, end]) =>
        Number.isInteger(start) && Number.isInteger(end) && start >= 0 && end > start,
    )
    .map(([start, end]): CharRange => [start, end])
    .sort((a, b) => a[0] - b[0] || a[1] - b[1]);

  const merged: CharRange[] = [];
  for (const range of sorted) {
    const last = merged.at(-1);
    if (last && range[0] <= last[1]) last[1] = Math.max(last[1], range[1]);
    else merged.push(range);
  }
  return merged;
}

/** True when one merged range covers the whole span (merged ranges never touch). */
function covered(merged: readonly CharRange[], span: Span): boolean {
  // Binary search for the last range that starts at or before span.start.
  let low = 0;
  let high = merged.length - 1;
  let found = -1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    if (merged[mid]![0] <= span.start) {
      found = mid;
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }
  const range = merged[found];
  return range !== undefined && range[1] >= span.end;
}

/** True when at least one character of the span is in a merged range. */
function touched(merged: readonly CharRange[], span: Span): boolean {
  // Binary search for the first range that ends after span.start.
  let low = 0;
  let high = merged.length;
  while (low < high) {
    const mid = (low + high) >> 1;
    if (merged[mid]![1] <= span.start) low = mid + 1;
    else high = mid;
  }
  const range = merged[low];
  return range !== undefined && range[0] < span.end;
}

/**
 * The units marked by `ranges`:
 * - word i (unit 2i) when at least one character of its span is covered: a
 *   partial highlight is rounded up to the whole word (decision 2026-11-27);
 * - gap i (unit 2i + 1) when it is joinable, both its words are marked and its
 *   whole span is covered.
 * The result is passed through `normalizeMarks`.
 */
export function unitsFromRanges(
  ranges: readonly (readonly [number, number])[],
  words: readonly Span[],
  gaps: readonly Gap[],
): number[] {
  const merged = mergeRanges(ranges);
  const units: number[] = [];
  words.forEach((word, i) => {
    if (touched(merged, word)) units.push(2 * i);
  });
  const marked = new Set(units);
  gaps.forEach((gap, i) => {
    if (gap.joinable && marked.has(2 * i) && marked.has(2 * i + 2) && covered(merged, gap)) {
      units.push(2 * i + 1);
    }
  });
  return normalizeMarks(units, words.length, joinableFromGaps(gaps));
}

/**
 * The character ranges of the marked units: the union of the spans of the
 * marked words and gaps, merged. The units are normalized first, so a gap that
 * is not joinable or lacks one of its words is never highlighted.
 */
export function rangesFromUnits(
  units: readonly number[],
  words: readonly Span[],
  gaps: readonly Gap[],
): CharRange[] {
  const valid = normalizeMarks(units, words.length, joinableFromGaps(gaps));
  const spans: CharRange[] = [];
  for (const unit of valid) {
    const span = unit % 2 === 0 ? words[unit / 2] : gaps[(unit - 1) / 2];
    if (span) spans.push([span.start, span.end]);
  }
  return mergeRanges(spans);
}
