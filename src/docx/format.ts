/**
 * Formatting for the preview (phase 4.4, feasibility report §5.2, D-10: faithful
 * but not pixel-exact).
 *
 * - Run formatting is resolved in layers, lowest first:
 *   docDefaults → paragraph style chain → character style chain → direct w:rPr.
 *   Each layer overrides only the properties it sets.
 * - Paragraph props: style id, heading level, alignment, indentation and list
 *   marker (bullet `•`, or a number counted per numId / ilvl in document order).
 *
 * WPS writes numeric style ids ("1" = Normal, "2" = heading 1), so headings are
 * recognised by style *name* ("heading N") or outline level, never by id.
 */
import type {
  Alignment,
  CharRef,
  Indent,
  ListMarker,
  ParagraphProps,
  RunFormat,
  RunView,
} from './types';
import { A, onOff, wAttr, wChild, wChildren, wInt } from './xml';

/* ------------------------------------------------------------------ */
/* Style table (word/styles.xml + theme fonts)                         */
/* ------------------------------------------------------------------ */

export type StyleDef = {
  id: string;
  type: string;
  name?: string;
  basedOn?: string;
  pPr?: Element;
  rPr?: Element;
};

export type StyleTable = {
  styles: Map<string, StyleDef>;
  defaultParagraphStyleId?: string;
  defaultCharacterStyleId?: string;
  docDefaultsRPr?: Element;
  docDefaultsPPr?: Element;
  /** Latin fonts of the theme, for `w:asciiTheme="minorHAnsi"` and similar. */
  themeFonts: { major?: string; minor?: string };
};

export function parseStyles(stylesXml?: XMLDocument, themeXml?: XMLDocument): StyleTable {
  const table: StyleTable = { styles: new Map(), themeFonts: parseThemeFonts(themeXml) };
  const root = stylesXml?.documentElement;
  if (!root) return table;

  const docDefaults = wChild(root, 'docDefaults');
  table.docDefaultsRPr = wChild(wChild(docDefaults, 'rPrDefault'), 'rPr');
  table.docDefaultsPPr = wChild(wChild(docDefaults, 'pPrDefault'), 'pPr');

  for (const style of wChildren(root, 'style')) {
    const id = wAttr(style, 'styleId');
    if (id === undefined) continue;
    const type = wAttr(style, 'type') ?? 'paragraph';
    const def: StyleDef = {
      id,
      type,
      name: wAttr(wChild(style, 'name'), 'val'),
      basedOn: wAttr(wChild(style, 'basedOn'), 'val'),
      pPr: wChild(style, 'pPr'),
      rPr: wChild(style, 'rPr'),
    };
    table.styles.set(id, def);
    if (['1', 'true', 'on'].includes(wAttr(style, 'default') ?? '')) {
      if (type === 'paragraph') table.defaultParagraphStyleId ??= id;
      if (type === 'character') table.defaultCharacterStyleId ??= id;
    }
  }
  return table;
}

function parseThemeFonts(themeXml?: XMLDocument): StyleTable['themeFonts'] {
  if (!themeXml) return {};
  const latin = (local: string) => {
    const font = themeXml.getElementsByTagNameNS(A, local)[0];
    const el = font ? Array.from(font.children).find((c) => c.localName === 'latin') : undefined;
    return el?.getAttribute('typeface') || undefined;
  };
  return { major: latin('majorFont'), minor: latin('minorFont') };
}

/** The style and its bases, base first (cycle-safe). */
export function styleChain(table: StyleTable, id: string | undefined): StyleDef[] {
  const chain: StyleDef[] = [];
  const seen = new Set<string>();
  for (let current = id; current !== undefined && !seen.has(current);) {
    seen.add(current);
    const def = table.styles.get(current);
    if (!def) break;
    chain.unshift(def);
    current = def.basedOn;
  }
  return chain;
}

/* ------------------------------------------------------------------ */
/* Numbering table (word/numbering.xml)                                */
/* ------------------------------------------------------------------ */

export type LevelDef = {
  start: number;
  numFmt: string;
  lvlText: string;
  pPr?: Element;
};

export type NumberingTable = {
  /** abstractNumId → ilvl → level. */
  abstracts: Map<string, Map<number, LevelDef>>;
  /** numId → abstractNumId plus per-level overrides. */
  nums: Map<string, { abstractId: string; overrides: Map<number, Partial<LevelDef>> }>;
};

function levelOf(lvl: Element): LevelDef {
  return {
    start: wInt(wChild(lvl, 'start'), 'val') ?? 1,
    numFmt: wAttr(wChild(lvl, 'numFmt'), 'val') ?? 'decimal',
    lvlText: wAttr(wChild(lvl, 'lvlText'), 'val') ?? '',
    pPr: wChild(lvl, 'pPr'),
  };
}

export function parseNumbering(numberingXml?: XMLDocument): NumberingTable {
  const table: NumberingTable = { abstracts: new Map(), nums: new Map() };
  const root = numberingXml?.documentElement;
  if (!root) return table;

  for (const abstract of wChildren(root, 'abstractNum')) {
    const id = wAttr(abstract, 'abstractNumId');
    if (id === undefined) continue;
    const levels = new Map<number, LevelDef>();
    for (const lvl of wChildren(abstract, 'lvl')) levels.set(wInt(lvl, 'ilvl') ?? 0, levelOf(lvl));
    table.abstracts.set(id, levels);
  }
  for (const num of wChildren(root, 'num')) {
    const id = wAttr(num, 'numId');
    const abstractId = wAttr(wChild(num, 'abstractNumId'), 'val');
    if (id === undefined || abstractId === undefined) continue;
    const overrides = new Map<number, Partial<LevelDef>>();
    for (const override of wChildren(num, 'lvlOverride')) {
      const ilvl = wInt(override, 'ilvl') ?? 0;
      const lvl = wChild(override, 'lvl');
      const start = wInt(wChild(override, 'startOverride'), 'val');
      overrides.set(ilvl, {
        ...(lvl ? levelOf(lvl) : {}),
        ...(start !== undefined ? { start } : {}),
      });
    }
    table.nums.set(id, { abstractId, overrides });
  }
  return table;
}

/** The level definition used by numId / ilvl, or undefined. */
export function numberingLevel(
  table: NumberingTable,
  numId: string,
  ilvl: number,
): LevelDef | undefined {
  const num = table.nums.get(numId);
  if (!num) return undefined;
  const base = table.abstracts.get(num.abstractId)?.get(ilvl);
  const override = num.overrides.get(ilvl);
  if (!base && !override) return undefined;
  return { start: 1, numFmt: 'decimal', lvlText: '', ...base, ...override };
}

function toRoman(value: number): string {
  if (value <= 0 || value >= 4000) return String(value);
  const table: [number, string][] = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let out = '';
  let rest = value;
  for (const [n, s] of table) {
    while (rest >= n) {
      out += s;
      rest -= n;
    }
  }
  return out;
}

/** a, b, …, z, aa, bb, … (the Word style). */
function toLetter(value: number): string {
  if (value <= 0) return String(value);
  const letter = String.fromCharCode(97 + ((value - 1) % 26));
  return letter.repeat(Math.floor((value - 1) / 26) + 1);
}

/** Formats one counter value for a numFmt (approximate: unknown formats fall back to decimal). */
export function formatNumber(value: number, numFmt: string): string {
  switch (numFmt) {
    case 'none':
      return '';
    case 'decimalZero':
      return value < 10 ? `0${value}` : String(value);
    case 'upperRoman':
      return toRoman(value);
    case 'lowerRoman':
      return toRoman(value).toLowerCase();
    case 'upperLetter':
      return toLetter(value).toUpperCase();
    case 'lowerLetter':
      return toLetter(value);
    default:
      return String(value);
  }
}

/**
 * Counts list items in document order. Call `next` once per list paragraph, in
 * paragraph order. Counters are kept per numId; a level restarts after a
 * paragraph at a shallower level of the same list (approximate, D-10).
 */
export function createListCounter(numbering: NumberingTable) {
  const counters = new Map<string, (number | undefined)[]>();
  return {
    next(numId: string, ilvl: number): ListMarker | undefined {
      const level = numberingLevel(numbering, numId, ilvl);
      if (!level) return undefined;
      if (level.numFmt === 'bullet') return { numId, ilvl, kind: 'bullet', text: '•' };

      const values = counters.get(numId) ?? [];
      counters.set(numId, values);
      values[ilvl] = (values[ilvl] ?? level.start - 1) + 1;
      values.length = ilvl + 1; // deeper levels restart

      const text = level.lvlText.replace(/%([1-9])/g, (_, digit: string) => {
        const n = Number(digit) - 1;
        const def = numberingLevel(numbering, numId, n);
        const value = values[n] ?? def?.start ?? 1;
        return formatNumber(value, def?.numFmt ?? 'decimal');
      });
      return { numId, ilvl, kind: 'number', text };
    },
  };
}

/* ------------------------------------------------------------------ */
/* Run formatting                                                      */
/* ------------------------------------------------------------------ */

/** What one rPr layer sets. `null` means "explicitly reset" (w:color auto, w:highlight none). */
type RunLayer = {
  font?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string | null;
  highlight?: string | null;
};

function themeFont(table: StyleTable, theme: string): string | undefined {
  return theme.startsWith('major') ? table.themeFonts.major : table.themeFonts.minor;
}

function runLayer(table: StyleTable, rPr: Element | undefined): RunLayer {
  const layer: RunLayer = {};
  if (!rPr) return layer;

  const fonts = wChild(rPr, 'rFonts');
  if (fonts) {
    // Within one layer a theme font wins over the explicit name (ECMA-376 §17.3.2.26).
    const theme = wAttr(fonts, 'asciiTheme') ?? wAttr(fonts, 'hAnsiTheme');
    const font = theme ? themeFont(table, theme) : (wAttr(fonts, 'ascii') ?? wAttr(fonts, 'hAnsi'));
    if (font) layer.font = font;
  }
  const sz = wInt(wChild(rPr, 'sz'), 'val');
  if (sz !== undefined) layer.sizePt = sz / 2;
  const bold = onOff(wChild(rPr, 'b'));
  if (bold !== undefined) layer.bold = bold;
  const italic = onOff(wChild(rPr, 'i'));
  if (italic !== undefined) layer.italic = italic;
  const u = wChild(rPr, 'u');
  if (u) layer.underline = (wAttr(u, 'val') ?? 'single') !== 'none';
  const color = wAttr(wChild(rPr, 'color'), 'val');
  if (color !== undefined)
    layer.color = /^[0-9a-f]{6}$/i.test(color) ? `#${color.toUpperCase()}` : null;
  const highlight = wAttr(wChild(rPr, 'highlight'), 'val');
  if (highlight !== undefined) layer.highlight = highlight === 'none' ? null : highlight;
  return layer;
}

function applyLayer(target: RunLayer, layer: RunLayer): void {
  for (const key of Object.keys(layer) as (keyof RunLayer)[]) {
    (target as Record<string, unknown>)[key] = layer[key];
  }
}

/** The effective paragraph style id: w:pStyle, or the default paragraph style. */
export function paragraphStyleId(table: StyleTable, p: Element): string | undefined {
  return wAttr(wChild(wChild(p, 'pPr'), 'pStyle'), 'val') ?? table.defaultParagraphStyleId;
}

/** Resolves the formatting of one w:r inside a paragraph of style `pStyleId`. */
export function resolveRunFormat(
  table: StyleTable,
  pStyleId: string | undefined,
  run: Element,
): RunFormat {
  const out: RunLayer = {};
  applyLayer(out, runLayer(table, table.docDefaultsRPr));
  for (const style of styleChain(table, pStyleId)) applyLayer(out, runLayer(table, style.rPr));
  const rPr = wChild(run, 'rPr');
  const rStyle = wAttr(wChild(rPr, 'rStyle'), 'val') ?? table.defaultCharacterStyleId;
  for (const style of styleChain(table, rStyle)) applyLayer(out, runLayer(table, style.rPr));
  applyLayer(out, runLayer(table, rPr));

  const format: RunFormat = {};
  if (out.font) format.fontFamily = out.font;
  if (out.sizePt !== undefined) format.sizePt = out.sizePt;
  if (out.bold) format.bold = true;
  if (out.italic) format.italic = true;
  if (out.underline) format.underline = true;
  if (out.color) format.color = out.color;
  if (out.highlight) format.highlight = out.highlight;
  return format;
}

/**
 * Splits the flattened text into pieces of one run each, with their resolved
 * formatting. darkYellow is not copied here: it is shown through the marks (D-08).
 */
export function buildRunViews(
  table: StyleTable,
  pStyleId: string | undefined,
  charMap: readonly CharRef[],
): RunView[] {
  const views: RunView[] = [];
  const cache = new Map<Element, RunFormat>();
  let start = 0;
  while (start < charMap.length) {
    const run = charMap[start]!.run;
    let end = start + 1;
    while (end < charMap.length && charMap[end]!.run === run) end += 1;

    let format = cache.get(run);
    if (!format) {
      format = resolveRunFormat(table, pStyleId, run);
      cache.set(run, format);
    }
    const { highlight, ...rest } = format;
    const view: RunView = { start, end, ...rest };
    if (highlight && highlight !== 'darkYellow') view.originalHighlight = highlight;
    views.push(view);
    start = end;
  }
  return views;
}

/* ------------------------------------------------------------------ */
/* Paragraph props                                                     */
/* ------------------------------------------------------------------ */

const TWIPS_PER_POINT = 20;

type ParaLayer = {
  alignment?: Alignment;
  indent: Indent;
  numId?: string;
  ilvl?: number;
  outlineLvl?: number;
};

function alignmentOf(jc: string): Alignment | undefined {
  switch (jc) {
    case 'left':
    case 'start':
      return 'left';
    case 'center':
      return 'center';
    case 'right':
    case 'end':
      return 'right';
    case 'both':
    case 'distribute':
    case 'mediumKashida':
    case 'highKashida':
    case 'lowKashida':
    case 'thaiDistribute':
      return 'justify';
    default:
      return undefined;
  }
}

function applyParaLayer(target: ParaLayer, pPr: Element | undefined): void {
  if (!pPr) return;
  const jc = wAttr(wChild(pPr, 'jc'), 'val');
  const alignment = jc === undefined ? undefined : alignmentOf(jc);
  if (alignment) target.alignment = alignment;

  const ind = wChild(pPr, 'ind');
  if (ind) {
    const pt = (...names: string[]) => {
      for (const name of names) {
        const value = wInt(ind, name);
        if (value !== undefined) return value / TWIPS_PER_POINT;
      }
      return undefined;
    };
    const left = pt('left', 'start');
    const right = pt('right', 'end');
    const firstLine = pt('firstLine');
    const hanging = pt('hanging');
    if (left !== undefined) target.indent.leftPt = left;
    if (right !== undefined) target.indent.rightPt = right;
    // firstLine and hanging exclude each other: the later layer replaces both.
    if (firstLine !== undefined || hanging !== undefined) {
      delete target.indent.firstLinePt;
      delete target.indent.hangingPt;
      if (hanging !== undefined) target.indent.hangingPt = hanging;
      else if (firstLine !== undefined) target.indent.firstLinePt = firstLine;
    }
  }

  const numPr = wChild(pPr, 'numPr');
  if (numPr) {
    const numId = wAttr(wChild(numPr, 'numId'), 'val');
    const ilvl = wInt(wChild(numPr, 'ilvl'), 'val');
    if (numId !== undefined) target.numId = numId;
    if (ilvl !== undefined) target.ilvl = ilvl;
  }
  const outlineLvl = wInt(wChild(pPr, 'outlineLvl'), 'val');
  if (outlineLvl !== undefined) target.outlineLvl = outlineLvl;
}

/** "heading 1" … "heading 9" (Word and WPS built-in names, any case). */
const HEADING_NAME = /^heading\s*([1-9])$/i;

/**
 * Resolves the props of one paragraph. `counter` must be shared by all the
 * paragraphs of the document and called in document order.
 */
export function resolveParagraphProps(
  table: StyleTable,
  numbering: NumberingTable,
  counter: ReturnType<typeof createListCounter>,
  p: Element,
): ParagraphProps {
  const styleId = paragraphStyleId(table, p);
  const chain = styleChain(table, styleId);
  const directPPr = wChild(p, 'pPr');

  // Layers without numbering indentation first, to learn numId / ilvl.
  const layer: ParaLayer = { indent: {} };
  applyParaLayer(layer, table.docDefaultsPPr);
  for (const style of chain) applyParaLayer(layer, style.pPr);
  const styleOnly: ParaLayer = { ...layer, indent: { ...layer.indent } };
  applyParaLayer(layer, directPPr);

  const props: ParagraphProps = {};
  if (styleId !== undefined) props.styleId = styleId;

  let list: ListMarker | undefined;
  if (layer.numId !== undefined && layer.numId !== '0') {
    const ilvl = layer.ilvl ?? 0;
    list = counter.next(layer.numId, ilvl);
    // Indentation: style → numbering level → direct.
    const level = numberingLevel(numbering, layer.numId, ilvl);
    if (level?.pPr) {
      const withList: ParaLayer = { ...styleOnly, indent: { ...styleOnly.indent } };
      applyParaLayer(withList, level.pPr);
      applyParaLayer(withList, directPPr);
      layer.indent = withList.indent;
    }
  }
  if (list) props.list = list;

  for (let i = chain.length - 1; i >= 0; i -= 1) {
    const match = HEADING_NAME.exec(chain[i]!.name ?? '');
    if (match) {
      props.headingLevel = Number(match[1]);
      break;
    }
  }
  if (props.headingLevel === undefined && layer.outlineLvl !== undefined && layer.outlineLvl < 9) {
    props.headingLevel = layer.outlineLvl + 1;
  }
  if (layer.alignment) props.alignment = layer.alignment;
  if (Object.keys(layer.indent).length > 0) props.indent = layer.indent;
  return props;
}
