/**
 * Body block tree for rendering (phase 4.3): paragraphs and tables in order,
 * with gridSpan / vMerge and nested tables.
 *
 * Block-level content controls (w:sdt → w:sdtContent) and w:customXml wrappers
 * are transparent: their content is read as if it were in place. Anything else
 * (w:sectPr, bookmarks, …) is not rendered.
 */
import type { Block, TableCell } from './types';
import { W, isW, wAttr, wChild, wInt } from './xml';

/** Children of `parent`, with w:sdt / w:customXml wrappers opened. */
function contentChildren(parent: Element): Element[] {
  const out: Element[] = [];
  for (const child of Array.from(parent.children)) {
    if (isW(child, 'sdt')) {
      const content = wChild(child, 'sdtContent');
      if (content) out.push(...contentChildren(content));
    } else if (isW(child, 'customXml')) {
      out.push(...contentChildren(child));
    } else {
      out.push(child);
    }
  }
  return out;
}

function cellOf(tc: Element, indexOf: ReadonlyMap<Element, number>): TableCell {
  const tcPr = wChild(tc, 'tcPr');
  const cell: TableCell = {
    blocks: blocksOf(tc, indexOf),
    gridSpan: Math.max(1, wInt(wChild(tcPr, 'gridSpan'), 'val') ?? 1),
  };
  const vMerge = wChild(tcPr, 'vMerge');
  if (vMerge) {
    // A bare <w:vMerge/> means "continue".
    cell.vMerge = wAttr(vMerge, 'val') === 'restart' ? 'restart' : 'continue';
  }
  return cell;
}

function tableOf(tbl: Element, indexOf: ReadonlyMap<Element, number>): Block {
  const rows = contentChildren(tbl)
    .filter((child) => isW(child, 'tr'))
    .map((tr) => ({
      cells: contentChildren(tr)
        .filter((child) => isW(child, 'tc'))
        .map((tc) => cellOf(tc, indexOf)),
    }));
  return { kind: 'table', rows };
}

/** Blocks directly inside `container` (w:body or w:tc). */
export function blocksOf(container: Element, indexOf: ReadonlyMap<Element, number>): Block[] {
  const blocks: Block[] = [];
  for (const child of contentChildren(container)) {
    if (isW(child, 'p')) {
      const paragraphIndex = indexOf.get(child);
      if (paragraphIndex !== undefined) blocks.push({ kind: 'paragraph', paragraphIndex });
    } else if (isW(child, 'tbl')) {
      blocks.push(tableOf(child, indexOf));
    }
  }
  return blocks;
}

/**
 * The body as a block tree. `paragraphs` is the ordered list from
 * `collectParagraphs`; a block's `paragraphIndex` is the position in it.
 */
export function buildBlocks(xml: XMLDocument, paragraphs: readonly Element[]): Block[] {
  const body = xml.getElementsByTagNameNS(W, 'body')[0];
  if (!body) return [];
  const indexOf = new Map(paragraphs.map((p, i) => [p, i] as const));
  return blocksOf(body, indexOf);
}
