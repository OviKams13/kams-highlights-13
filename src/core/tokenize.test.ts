import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { TOKENIZER_VERSION, tokenize } from './tokenize';

/** Words and gaps as text, for readable expectations. */
function split(text: string, hardBreaks: number[] = []) {
  const { words, gaps } = tokenize(text, hardBreaks);
  return {
    words: words.map(({ start, end }) => text.slice(start, end)),
    gaps: gaps.map(({ start, end }) => text.slice(start, end)),
    joinable: gaps.map((gap) => gap.joinable),
  };
}

/** Offsets of the characters that stand for w:tab / w:br / w:cr in these tests. */
function breaksOf(text: string): number[] {
  return [...text.matchAll(/[\t\n]/g)].map((match) => match.index);
}

describe('tokenize (D-03, D-05)', () => {
  it('has version 1', () => {
    expect(TOKENIZER_VERSION).toBe(1);
  });

  it('keeps apostrophes, hyphens and inner dots inside a word', () => {
    expect(split("I'm a full-stack developer using Node.js").words).toEqual([
      "I'm",
      'a',
      'full-stack',
      'developer',
      'using',
      'Node.js',
    ]);
  });

  it('trims trailing punctuation: "2026." gives "2026", "TypeScript," gives "TypeScript"', () => {
    const result = split('Since 2026. I use TypeScript, React and Vite.');
    expect(result.words).toEqual([
      'Since',
      '2026',
      'I',
      'use',
      'TypeScript',
      'React',
      'and',
      'Vite',
    ]);
    expect(result.gaps).toEqual([' ', '. ', ' ', ' ', ', ', ' ', ' ']);
  });

  it('trims leading punctuation: "(see" gives "see"', () => {
    const result = split('(see below)');
    expect(result.words).toEqual(['see', 'below']);
  });

  it('drops a punctuation-only token into the gap: "a — b"', () => {
    const result = split('a — b');
    expect(result.words).toEqual(['a', 'b']);
    expect(result.gaps).toEqual([' — ']);
    expect(result.joinable).toEqual([true]);
  });

  it('trims the trailing apostrophe of "students\'"', () => {
    expect(split("the students' work").words).toEqual(['the', 'students', 'work']);
  });

  it('handles curly quotes and curly apostrophes', () => {
    const result = split('“pleased,” he said. I’m ‘fine’');
    expect(result.words).toEqual(['pleased', 'he', 'said', 'I’m', 'fine']);
    expect(result.gaps).toEqual([',” ', ' ', '. ', ' ‘']);
  });

  it('keeps numbers whole: 3.5, 1,000, 24/7', () => {
    expect(split('3.5 1,000 24/7').words).toEqual(['3.5', '1,000', '24/7']);
  });

  it('makes several spaces one gap', () => {
    const result = split('a   b');
    expect(result.gaps).toEqual(['   ']);
    expect(tokenize('a   b').gaps).toEqual([{ start: 1, end: 4, joinable: true }]);
  });

  it('makes a gap with a tab not joinable', () => {
    const text = 'Name:\tKams';
    const result = split(text, breaksOf(text));
    expect(result.words).toEqual(['Name', 'Kams']);
    expect(result.gaps).toEqual([':\t']);
    expect(result.joinable).toEqual([false]);
  });

  it('makes a gap with a line break not joinable, and leaves the others joinable', () => {
    const text = 'one two\nthree four';
    const result = split(text, breaksOf(text));
    expect(result.words).toEqual(['one', 'two', 'three', 'four']);
    expect(result.joinable).toEqual([true, false, true]);
  });

  it('decides joinable from hardBreaks only (in any order)', () => {
    expect(tokenize('a b c', [3, 1]).gaps.map((gap) => gap.joinable)).toEqual([false, false]);
    expect(tokenize('a b c', [10]).gaps.map((gap) => gap.joinable)).toEqual([true, true]);
  });

  it('splits on a no-break space and a narrow no-break space (joinable gaps)', () => {
    expect(split('a b').words).toEqual(['a', 'b']);
    expect(split('a b').joinable).toEqual([true]);
    expect(split('prix : 10 €').words).toEqual(['prix', '10']);
  });

  it('returns nothing for empty, blank or punctuation-only text', () => {
    for (const text of ['', '   ', '\t', '— , .', '“”']) {
      expect(tokenize(text)).toEqual({ words: [], gaps: [] });
    }
  });

  it('gives one word and no gap for a single word', () => {
    expect(tokenize('  hello  ')).toEqual({ words: [{ start: 2, end: 7 }], gaps: [] });
  });

  it('trims whole code points outside the BMP', () => {
    expect(split('x😀 𝐀𝐁 😀').words).toEqual(['x', '𝐀𝐁']);
    expect(split('😀').words).toEqual([]);
  });

  it('keeps letters from other scripts', () => {
    expect(split('café naïve Ελλάδα 東京 مرحبا').words).toEqual([
      'café',
      'naïve',
      'Ελλάδα',
      '東京',
      'مرحبا',
    ]);
  });

  it('keeps a decomposed accent in its word (core = \\p{L}\\p{M}\\p{N}, decision 2026-11-27)', () => {
    const cafe = 'café'; // "café" as e + COMBINING ACUTE ACCENT
    const text = `un ${cafe}. ${cafe}`;
    expect(split(text).words).toEqual(['un', cafe, cafe]);
    expect(tokenize(text).words[1]).toEqual({ start: 3, end: 8 });
    expect(split(`(${cafe})`).words).toEqual([cafe]);
  });

  it('always returns sorted words with core ends and gaps between them (property)', () => {
    const char = fc.constantFrom('a', 'Z', '7', 'é', '́', ' ', '\t', '\n', ' ', '.', ',', '—', "'");
    fc.assert(
      fc.property(fc.array(char, { maxLength: 40 }), (chars) => {
        const text = chars.join('');
        const hard = breaksOf(text);
        const { words, gaps } = tokenize(text, hard);
        expect(gaps).toHaveLength(Math.max(0, words.length - 1));
        words.forEach((word, i) => {
          expect(word.end).toBeGreaterThan(word.start);
          expect(text[word.start]).toMatch(/[\p{L}\p{M}\p{N}]/u);
          expect(text[word.end - 1]).toMatch(/[\p{L}\p{M}\p{N}]/u);
          expect(text.slice(word.start, word.end)).not.toMatch(/\s/);
          const gap = gaps[i];
          if (gap) {
            expect(gap.start).toBe(word.end);
            expect(gap.end).toBe(words[i + 1]!.start);
            expect(text.slice(gap.start, gap.end)).toMatch(/\s/);
            expect(gap.joinable).toBe(!/[\t\n]/.test(text.slice(gap.start, gap.end)));
          }
        });
      }),
    );
  });
});
