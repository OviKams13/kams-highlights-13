/**
 * Writes the file for the real WPS test (phase 5.4): a.fixtures/test.docx with
 * the words "Cyprus International" highlighted and joined, saved next to it as
 * a.fixtures/test_highlighted.docx. The highlights already in the file are kept.
 *
 * Run with `npm run export-sample`. Local only: a.fixtures/ is never committed.
 *
 * This file runs directly under Node's type stripping: keep it free of enums,
 * namespaces and parameter properties. The app code in src/ is written for Vite
 * and the browser, so before loading it this script gives Node:
 * - a resolve hook for the `@/` alias and the imports without `.ts`;
 * - DOMParser and XMLSerializer, from jsdom (a dev dependency).
 * That is why src/ is loaded with `import()`: static imports would run first.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { createRequire, registerHooks } from 'node:module';
import { dirname, extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const PROJECT_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const SRC = new URL('../../src/', import.meta.url).href;
const FIXTURES_DIR = join(PROJECT_ROOT, 'a.fixtures');
const SOURCE_NAME = 'test.docx';
const FIRST = 'Cyprus';
const SECOND = 'International';

registerHooks({
  resolve(specifier, context, nextResolve) {
    const alias = specifier.startsWith('@/');
    const fromSource = specifier.startsWith('.') && context.parentURL?.startsWith(SRC) === true;
    if (!alias && !fromSource) return nextResolve(specifier, context);
    const target = alias ? new URL(specifier.slice(2), SRC).href : specifier;
    return nextResolve(extname(target) === '.ts' ? target : `${target}.ts`, context);
  },
});

const { JSDOM } = createRequire(import.meta.url)('jsdom') as {
  JSDOM: new (html: string) => { window: typeof globalThis };
};
const { window } = new JSDOM('');
globalThis.DOMParser = window.DOMParser;
globalThis.XMLSerializer = window.XMLSerializer;

const { joinableFromGaps, toggleMark } = await import('../../src/core/units.ts');
const { rangesFromUnits } = await import('../../src/core/marks.ts');
const { openDocx } = await import('../../src/docx/open.ts');
const { readDocumentMarks } = await import('../../src/docx/read-marks.ts');
const { exportDocx } = await import('../../src/docx/export.ts');

const sourcePath = join(FIXTURES_DIR, SOURCE_NAME);
if (!existsSync(sourcePath)) {
  throw new Error(`${sourcePath} is missing: this script needs the real WPS file.`);
}
const doc = await openDocx(readFileSync(sourcePath), SOURCE_NAME);

// The marks already in the file. A paragraph left out of the map would lose its highlights.
const marks = new Map(
  [...readDocumentMarks(doc)].map(([index, { units }]) => [index, units] as const),
);

// The first place where the two words follow each other.
const found = doc.paragraphs
  .flatMap((paragraph) => {
    const textOf = (i: number) => {
      const word = paragraph.words[i];
      return word && paragraph.text.slice(word.start, word.end);
    };
    return paragraph.words.map((_, i) => ({
      paragraph,
      i,
      match: textOf(i) === FIRST && textOf(i + 1) === SECOND,
    }));
  })
  .find((candidate) => candidate.match);
if (!found) throw new Error(`"${FIRST} ${SECOND}" is not in ${SOURCE_NAME}.`);

// Same rules as a tap in the app: the two words first, then the gap between them.
const { paragraph, i } = found;
const joinable = joinableFromGaps(paragraph.gaps);
let units = marks.get(paragraph.index) ?? [];
for (const unit of [2 * i, 2 * i + 2, 2 * i + 1]) {
  if (!units.includes(unit)) {
    units = toggleMark(units, unit, paragraph.words.length, joinable);
  }
}
if (!units.includes(2 * i + 1)) throw new Error(`"${FIRST}" and "${SECOND}" cannot be joined.`);
marks.set(paragraph.index, units);

const file = await exportDocx(doc, marks);
const outPath = join(FIXTURES_DIR, file.name);
writeFileSync(outPath, new Uint8Array(await file.arrayBuffer()));

// Read the written file again and show what is highlighted in that paragraph.
const written = await openDocx(readFileSync(outPath), file.name);
const result = written.paragraphs[paragraph.index]!;
const saved = readDocumentMarks(written).get(paragraph.index)?.units ?? [];
const highlighted = rangesFromUnits(saved, result.words, result.gaps).map(([start, end]) =>
  result.text.slice(start, end),
);
if (!highlighted.some((text) => text.includes(`${FIRST} ${SECOND}`))) {
  throw new Error(`${outPath} was written, but "${FIRST} ${SECOND}" is not highlighted in it.`);
}
console.log(`Wrote ${outPath}`);
console.log(
  `Highlighted in paragraph ${paragraph.index}: ${highlighted.map((t) => `[${t}]`).join(' ')}`,
);
