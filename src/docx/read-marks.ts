/**
 * Existing darkYellow highlights → marks (phase 4.5, D-07, D-08).
 *
 * Mirrors `run_highlighted()` and `read_marks()` in a.prototype/docx_roundtrip.py:
 * a character is highlighted when its run's *direct* w:rPr has
 * `w:highlight w:val="darkYellow"` (that is what WPS writes, and what the
 * writer will change). Other highlight colours are not marks: they are kept in
 * the file and shown through `RunView.originalHighlight`.
 *
 * The decisions of 2026-11-27 apply through `unitsFromRanges`: a word with at
 * least one darkYellow character is marked whole; a gap only when it is fully
 * covered and both its words are marked.
 */
import { type CharRange, mergeRanges, rangesFromUnits, unitsFromRanges } from '@/core/marks';
import type { CharRef, ParagraphModel } from './types';
import { wAttr, wChild } from './xml';

export const DARK_YELLOW = 'darkYellow';

/** True when the run's direct rPr has a darkYellow highlight. */
export function isDarkYellowRun(run: Element): boolean {
  return wAttr(wChild(wChild(run, 'rPr'), 'highlight'), 'val') === DARK_YELLOW;
}

/** The darkYellow characters of a flattened paragraph, as raw merged ranges. */
export function darkYellowRanges(charMap: readonly CharRef[]): CharRange[] {
  const ranges: CharRange[] = [];
  const cache = new Map<Element, boolean>();
  charMap.forEach(({ run }, i) => {
    let on = cache.get(run);
    if (on === undefined) {
      on = isDarkYellowRun(run);
      cache.set(run, on);
    }
    if (on) ranges.push([i, i + 1]);
  });
  return mergeRanges(ranges);
}

export type ParagraphMarks = {
  /** Marked units (word i = 2i, gap i = 2i + 1), normalized. */
  units: number[];
  /** The canonical ranges of those units: whole words and joined gaps only. */
  ranges: CharRange[];
};

/** The marks of one paragraph, read from its darkYellow runs. */
export function readParagraphMarks(
  paragraph: Pick<ParagraphModel, 'charMap' | 'words' | 'gaps'>,
): ParagraphMarks {
  const raw = darkYellowRanges(paragraph.charMap);
  const units = unitsFromRanges(raw, paragraph.words, paragraph.gaps);
  return { units, ranges: rangesFromUnits(units, paragraph.words, paragraph.gaps) };
}

/**
 * The initial marks of a document (ERD A.2 "initial rows"): one entry per
 * paragraph that has at least one mark, keyed by paragraphIndex.
 */
export function readDocumentMarks(doc: {
  paragraphs: readonly Pick<ParagraphModel, 'index' | 'charMap' | 'words' | 'gaps'>[];
}): Map<number, ParagraphMarks> {
  const out = new Map<number, ParagraphMarks>();
  for (const paragraph of doc.paragraphs) {
    const marks = readParagraphMarks(paragraph);
    if (marks.units.length > 0) out.set(paragraph.index, marks);
  }
  return out;
}
