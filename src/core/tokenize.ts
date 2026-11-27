/**
 * Paragraph tokenizer (D-03, D-05), mirroring `tokenize()` in
 * a.prototype/docx_roundtrip.py.
 *
 * - A token is a run of non-whitespace characters (`\S+`).
 * - A word is a token with its leading and trailing non-core characters
 *   trimmed. Core characters are letters, combining marks and digits
 *   (`\p{L}\p{M}\p{N}`, decision 2026-11-27: a decomposed accent such as
 *   `e` + U+0301 stays with its word). So `I'm`, `full-stack`, `Node.js` and
 *   `2026` are one word each; `good,` gives `good`.
 * - A token with no core character (e.g. `—`) is not a word: it becomes part
 *   of the gap around it.
 * - Gap i is everything between word i and word i + 1 (spaces and any
 *   punctuation). It is joinable unless it contains a hard break (w:tab,
 *   w:br, w:cr).
 *
 * Offsets are UTF-16 indices into `text` (JavaScript string indices), [start, end).
 * Text before the first word and after the last one is not a unit.
 */

/** Bump when the word or gap rule changes (stored with each document, ERD A.1). */
export const TOKENIZER_VERSION = 1;

/** A span of the paragraph text, [start, end). */
export type Span = { start: number; end: number };

/** The text between two words. `joinable` is false when it contains a hard break. */
export type Gap = Span & { joinable: boolean };

export type Tokens = {
  /** Word i is unit 2i. */
  words: Span[];
  /** Gap i (between word i and i + 1) is unit 2i + 1. Always `words.length - 1` long (or 0). */
  gaps: Gap[];
};

const TOKEN = /\S+/gu;
const CORE = /^[\p{L}\p{M}\p{N}]$/u;

function isCore(char: string): boolean {
  return CORE.test(char);
}

/**
 * Splits a flattened paragraph into words and gaps.
 *
 * @param text the visible text of the paragraph (feasibility report §5.1)
 * @param hardBreaks offsets in `text` of the w:tab / w:br / w:cr characters
 */
export function tokenize(text: string, hardBreaks: readonly number[] = []): Tokens {
  const words: Span[] = [];
  for (const match of text.matchAll(TOKEN)) {
    let start = match.index;
    let end = start + match[0].length;
    // Trim whole code points, so a letter outside the BMP is never cut in half.
    while (start < end) {
      const char = String.fromCodePoint(text.codePointAt(start)!);
      if (isCore(char)) break;
      start += char.length;
    }
    while (end > start) {
      const char = lastCodePoint(text, end);
      if (isCore(char)) break;
      end -= char.length;
    }
    if (start < end) words.push({ start, end });
  }

  const breaks = [...hardBreaks].sort((a, b) => a - b);
  const gaps: Gap[] = [];
  for (let i = 0; i + 1 < words.length; i += 1) {
    const start = words[i]!.end;
    const end = words[i + 1]!.start;
    gaps.push({ start, end, joinable: !breaks.some((offset) => offset >= start && offset < end) });
  }
  return { words, gaps };
}

/** The code point that ends just before `end` (one or two UTF-16 units). */
function lastCodePoint(text: string, end: number): string {
  const low = text.charCodeAt(end - 1);
  if (low >= 0xdc00 && low <= 0xdfff && end >= 2) {
    const high = text.charCodeAt(end - 2);
    if (high >= 0xd800 && high <= 0xdbff) return text.slice(end - 2, end);
  }
  return text.slice(end - 1, end);
}
