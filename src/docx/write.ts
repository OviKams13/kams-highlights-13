/**
 * Marks → runs (phase 5.1, feasibility report §5.3).
 *
 * Mirrors `apply_marks()` and `set_highlight()` in a.prototype/docx_roundtrip.py:
 * each run is split where the wanted state changes, every piece gets a deep copy
 * of the original w:rPr, and only the direct `w:highlight` of that copy changes:
 * - wanted piece → `darkYellow`;
 * - other piece → the original highlight, removed if it was darkYellow and kept
 *   if it was another colour (ERD "Export").
 *
 * A run that already has the right highlight everywhere is not touched at all.
 */
import type { CharRange } from '@/core/marks';
import { flattenParagraph } from './paragraphs';
import { DARK_YELLOW } from './read-marks';
import { W, isW, wAttr, wChild } from './xml';

const XML_NS = 'http://www.w3.org/XML/1998/namespace';

/**
 * The w:rPr children that come after w:highlight in the schema order (CT_RPr).
 * Elements from other namespaces (w14:*, …) also come after it.
 */
const AFTER_HIGHLIGHT: ReadonlySet<string> = new Set([
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
  'eastAsianLayout',
  'specVanish',
  'oMath',
  'rPrChange',
]);

/** A part of a run whose characters all have the same wanted state. */
type Piece = { on: boolean; nodes: ChildNode[] };

/** Creates `w:<local>` with the same prefix as the element `like`. */
function createW(like: Element, local: string): Element {
  const name = like.prefix ? `${like.prefix}:${local}` : local;
  return like.ownerDocument.createElementNS(W, name);
}

/** Sets or removes `w:highlight` in the direct w:rPr of `run`. */
function setHighlight(run: Element, value: string | undefined): void {
  let rPr = wChild(run, 'rPr');
  let highlight = wChild(rPr, 'highlight');

  if (value === undefined) {
    if (!rPr || !highlight) return;
    highlight.remove();
    // An rPr that only held the highlight goes away with it.
    if (rPr.children.length === 0 && rPr.attributes.length === 0) rPr.remove();
    return;
  }

  if (!rPr) {
    rPr = createW(run, 'rPr');
    run.insertBefore(rPr, run.firstChild);
  }
  if (!highlight) {
    highlight = createW(run, 'highlight');
    const next = Array.from(rPr.children).find(
      (child) => child.namespaceURI !== W || AFTER_HIGHLIGHT.has(child.localName),
    );
    rPr.insertBefore(highlight, next ?? null);
  }
  highlight.setAttributeNS(W, `${run.prefix ?? 'w'}:val`, value);
}

/** A copy of the w:t `node` holding `text`, with `xml:space="preserve"` when it needs it. */
function textPiece(node: Element, text: string): Element {
  const copy = node.cloneNode(false) as Element;
  copy.textContent = text;
  if (/^\s|\s$/.test(text)) copy.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  else copy.removeAttributeNS(XML_NS, 'space');
  return copy;
}

/**
 * The content of `run` (without w:rPr) cut into same-state pieces; adjacent
 * content with the same state stays in one piece. Content without characters
 * (w:softHyphen, w:drawing, w:fldChar, …) stays with the piece before it, or
 * with the first piece. A run without any character gives no piece.
 */
function runPieces(
  run: Element,
  firstIndex: ReadonlyMap<Node, number>,
  wanted: readonly boolean[],
): Piece[] {
  const pieces: Piece[] = [];
  let leading: ChildNode[] = [];
  const add = (on: boolean, node: ChildNode) => {
    const last = pieces.at(-1);
    if (last && last.on === on) last.nodes.push(node);
    else {
      pieces.push({ on, nodes: [...leading, node] });
      leading = [];
    }
  };

  for (const child of Array.from(run.childNodes)) {
    if (isW(child, 'rPr')) continue;
    const start = firstIndex.get(child);
    if (start === undefined) {
      const last = pieces.at(-1);
      if (last) last.nodes.push(child);
      else leading.push(child);
      continue;
    }
    if (!isW(child, 't')) {
      add(wanted[start] === true, child);
      continue;
    }
    const text = child.textContent ?? '';
    let from = 0;
    for (let i = 1; i <= text.length; i += 1) {
      const on = wanted[start + from] === true;
      if (i < text.length && (wanted[start + i] === true) === on) continue;
      // The original element is kept when it does not need to be cut.
      add(on, from === 0 && i === text.length ? child : textPiece(child, text.slice(from, i)));
      from = i;
    }
  }
  return pieces;
}

function writeRun(
  run: Element,
  firstIndex: ReadonlyMap<Node, number>,
  wanted: readonly boolean[],
): void {
  const pieces = runPieces(run, firstIndex, wanted);
  const rPr = wChild(run, 'rPr');
  const original = wAttr(wChild(rPr, 'highlight'), 'val');
  const target = (on: boolean) => {
    if (on) return DARK_YELLOW;
    return original === DARK_YELLOW ? undefined : original;
  };

  const [only, ...rest] = pieces;
  if (!only) return;
  if (rest.length === 0) {
    if (target(only.on) !== original) setHighlight(run, target(only.on));
    return;
  }

  for (const piece of pieces) {
    const copy = run.cloneNode(false) as Element; // keeps the w:r attributes (rsid, …)
    if (rPr) copy.appendChild(rPr.cloneNode(true));
    copy.append(...piece.nodes);
    setHighlight(copy, target(piece.on));
    run.before(copy);
  }
  run.remove();
}

/**
 * Makes exactly the characters of `ranges` darkYellow in the w:p `paragraph`
 * (changed in place). `ranges` are offsets in the flattened paragraph text.
 */
export function applyMarks(paragraph: Element, ranges: readonly CharRange[]): void {
  const { text, charMap, runs } = flattenParagraph(paragraph);
  const wanted = new Array<boolean>(text.length).fill(false);
  for (const [start, end] of ranges) {
    wanted.fill(true, Math.max(0, start), Math.min(text.length, end));
  }
  const firstIndex = new Map<Node, number>();
  charMap.forEach(({ node }, i) => {
    if (!firstIndex.has(node)) firstIndex.set(node, i);
  });
  for (const run of runs) writeRun(run, firstIndex, wanted);
}
