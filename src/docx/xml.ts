/**
 * Small namespace-aware helpers over the browser DOM (DOMParser output).
 *
 * WordprocessingML elements are matched by namespace URI + local name, never by
 * prefix: a file may bind `w:` to the main namespace under any prefix.
 */

/** WordprocessingML main namespace (`w:`). */
export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
/** Markup compatibility namespace (`mc:`). */
export const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
/** DrawingML main namespace (`a:`), used by the theme part. */
export const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** True when `node` is the element `w:<local>`. */
export function isW(node: Node | null | undefined, local: string): node is Element {
  return (
    node !== null &&
    node !== undefined &&
    node.nodeType === 1 &&
    (node as Element).namespaceURI === W &&
    (node as Element).localName === local
  );
}

/** Element children of `parent` (any namespace). */
export function elementChildren(parent: Element): Element[] {
  return Array.from(parent.children);
}

/** First `w:<local>` child of `parent`, or undefined. */
export function wChild(parent: Element | null | undefined, local: string): Element | undefined {
  if (!parent) return undefined;
  for (const child of Array.from(parent.children)) {
    if (isW(child, local)) return child;
  }
  return undefined;
}

/** All `w:<local>` children of `parent`. */
export function wChildren(parent: Element | null | undefined, local: string): Element[] {
  if (!parent) return [];
  return Array.from(parent.children).filter((child) => isW(child, local));
}

/** The `w:<name>` attribute of `el`, or undefined. */
export function wAttr(el: Element | null | undefined, name: string): string | undefined {
  return el?.getAttributeNS(W, name) ?? undefined;
}

/** Integer value of `w:<name>`, or undefined. */
export function wInt(el: Element | null | undefined, name: string): number | undefined {
  const raw = wAttr(el, name);
  if (raw === undefined) return undefined;
  const value = Number.parseInt(raw, 10);
  return Number.isFinite(value) ? value : undefined;
}

/**
 * Value of an on/off property such as `<w:b/>` or `<w:b w:val="0"/>`:
 * undefined when the element is absent, true when it has no value or a true value.
 */
export function onOff(el: Element | undefined): boolean | undefined {
  if (!el) return undefined;
  const value = wAttr(el, 'val');
  return value === undefined || !['0', 'false', 'off', 'none'].includes(value);
}

/** True when `el` has an ancestor `w:<local>` (stops at the document root). */
export function hasWAncestor(el: Element, local: string): boolean {
  for (let node: Element | null = el.parentElement; node !== null; node = node.parentElement) {
    if (node.namespaceURI === W && node.localName === local) return true;
  }
  return false;
}

export class XmlParseError extends Error {
  constructor(part: string) {
    super(`${part} is not well-formed XML.`);
    this.name = 'XmlParseError';
  }
}

/** Parses an XML part with the browser DOMParser; throws `XmlParseError` when it is not well-formed. */
export function parseXml(text: string, part: string): XMLDocument {
  let doc: XMLDocument;
  try {
    doc = new DOMParser().parseFromString(text, 'application/xml');
  } catch {
    throw new XmlParseError(part);
  }
  if (doc.getElementsByTagName('parsererror').length > 0 || !doc.documentElement) {
    throw new XmlParseError(part);
  }
  return doc;
}
