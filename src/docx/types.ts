/**
 * In-memory document model (ERD Part B). Rebuilt from the file every time it is opened.
 */
import type JSZip from 'jszip';
import type { Gap, Span } from '@/core/tokenize';
import type { NumberingTable, StyleTable } from './format';

export type DocModel = {
  /** SHA-256 of the file bytes, lowercase hex (= documents.id in IndexedDB). */
  id: string;
  /** Original file name, e.g. `answer.docx`. */
  name: string;
  /** Size of the file in bytes. */
  size: number;
  /** The original package, never modified here. */
  zip: JSZip;
  /** Parsed word/document.xml. */
  xml: XMLDocument;
  styles: StyleTable;
  numbering: NumberingTable;
  /** Body in order, for rendering. */
  blocks: Block[];
  /** Every w:p in document order (index = paragraphIndex), text boxes excluded. */
  paragraphs: ParagraphModel[];
};

export type TableCell = {
  blocks: Block[];
  gridSpan: number;
  vMerge?: 'restart' | 'continue';
};

export type Block =
  { kind: 'paragraph'; paragraphIndex: number } | { kind: 'table'; rows: { cells: TableCell[] }[] };

/** Which XML node a character of the flattened text comes from. */
export type CharRef = {
  /** The w:r that holds the character. */
  run: Element;
  /** The w:t (or w:tab / w:br / w:cr / w:noBreakHyphen) element. */
  node: Element;
  /** Index in the w:t text; null for tab, br, cr and noBreakHyphen. */
  offset: number | null;
};

/** Result of flattening one paragraph (feasibility report §5.1). */
export type FlatParagraph = {
  /** Visible text: w:t, plus `\t` for w:tab, `\n` for w:br / w:cr, `-` for w:noBreakHyphen. */
  text: string;
  /** One entry per UTF-16 unit of `text`. */
  charMap: CharRef[];
  /** Offsets of the w:tab / w:br / w:cr characters. */
  hardBreaks: number[];
  /** The w:r elements of this paragraph, in order (nested paragraphs excluded). */
  runs: Element[];
};

export type Alignment = 'left' | 'center' | 'right' | 'justify';

/** Indentation in points (twips ÷ 20). */
export type Indent = {
  leftPt?: number;
  rightPt?: number;
  firstLinePt?: number;
  hangingPt?: number;
};

export type ListMarker = {
  numId: string;
  ilvl: number;
  kind: 'bullet' | 'number';
  /** Text to draw before the paragraph: `•` for bullets, e.g. `2.` or `1.3.` for numbers (approximate). */
  text: string;
};

export type ParagraphProps = {
  /** w:pStyle (or the default paragraph style). */
  styleId?: string;
  /** 1–9 for "heading N" styles or an outline level; undefined for body text. */
  headingLevel?: number;
  alignment?: Alignment;
  indent?: Indent;
  list?: ListMarker;
};

/** Resolved run formatting (direct rPr → character style → paragraph style → docDefaults). */
export type RunFormat = {
  fontFamily?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** `#RRGGBB`, undefined for `auto`. */
  color?: string;
  /** Highlight colour name (`darkYellow`, `yellow`, …), undefined when none. */
  highlight?: string;
};

/** A piece of text with one formatting, for rendering. */
export type RunView = {
  start: number;
  end: number;
  fontFamily?: string;
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  /** A highlight colour other than darkYellow (D-08). darkYellow is shown through the marks. */
  originalHighlight?: string;
};

export type ParagraphModel = {
  index: number;
  element: Element;
  text: string;
  charMap: CharRef[];
  hardBreaks: number[];
  /** Word i = unit 2i. */
  words: Span[];
  /** Gap i (between word i and i + 1) = unit 2i + 1. */
  gaps: Gap[];
  props: ParagraphProps;
  runs: RunView[];
};
