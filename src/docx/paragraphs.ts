/**
 * Paragraph collection and flattening (phase 4.2, feasibility report §5.1).
 *
 * Mirrors `paragraph_runs()` and `flatten()` in a.prototype/docx_roundtrip.py:
 * - every w:p in document order, at every depth (body and table cells), gets an
 *   index — except paragraphs inside text boxes (w:txbxContent), which are out of V1;
 * - a paragraph's text comes from the w:t of its own runs, including runs inside
 *   w:hyperlink, w:ins, w:smartTag, w:sdtContent, w:fldSimple…;
 * - w:tab → `\t`, w:br / w:cr → `\n` (hard breaks); w:noBreakHyphen → `-`;
 *   w:softHyphen, w:delText and w:instrText are not visible text and are skipped.
 *
 * Headers, footers and notes live in other parts of the package, so they are
 * never seen here.
 */
import type { CharRef, FlatParagraph } from './types';
import { MC, W, hasWAncestor, isW } from './xml';

/** Every w:p of the document in order, text-box paragraphs excluded. */
export function collectParagraphs(xml: XMLDocument): Element[] {
  return Array.from(xml.getElementsByTagNameNS(W, 'p')).filter(
    (p) => !hasWAncestor(p, 'txbxContent'),
  );
}

/**
 * Containers that never hold visible runs of this paragraph:
 * - properties (w:pPr, w:rPr);
 * - moved-away text in tracked changes (w:moveFrom), like w:del;
 * - the mc:Fallback copy of an mc:AlternateContent (the mc:Choice is read instead).
 */
function skipSubtree(el: Element): boolean {
  if (el.namespaceURI === MC && el.localName === 'Fallback') return true;
  return isW(el, 'pPr') || isW(el, 'rPr') || isW(el, 'moveFrom');
}

/** The w:r elements that belong to `p` itself (not to a paragraph nested in it), in order. */
export function paragraphRuns(p: Element): Element[] {
  const runs: Element[] = [];
  const walk = (parent: Element) => {
    for (const child of Array.from(parent.children)) {
      if (isW(child, 'r')) runs.push(child);
      else if (isW(child, 'p') || skipSubtree(child)) continue;
      else walk(child);
    }
  };
  walk(p);
  return runs;
}

/** Concatenates the visible text of `p` and maps each character back to its run and node. */
export function flattenParagraph(p: Element): FlatParagraph {
  let text = '';
  const charMap: CharRef[] = [];
  const hardBreaks: number[] = [];
  const runs = paragraphRuns(p);

  for (const run of runs) {
    for (const node of Array.from(run.children)) {
      if (node.namespaceURI !== W) continue;
      switch (node.localName) {
        case 't': {
          const value = node.textContent ?? '';
          for (let offset = 0; offset < value.length; offset += 1) {
            charMap.push({ run, node, offset });
          }
          text += value;
          break;
        }
        case 'tab':
        case 'br':
        case 'cr':
          hardBreaks.push(text.length);
          charMap.push({ run, node, offset: null });
          text += node.localName === 'tab' ? '\t' : '\n';
          break;
        case 'noBreakHyphen':
          charMap.push({ run, node, offset: null });
          text += '-';
          break;
        default:
          // w:softHyphen, w:delText, w:instrText, w:fldChar, w:drawing, …: not visible text.
          break;
      }
    }
  }
  return { text, charMap, hardBreaks, runs };
}
