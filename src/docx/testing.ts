/**
 * Test helpers for src/docx (used by *.test.ts only; never imported by app code).
 * Builds small WordprocessingML documents from body XML strings.
 */
import JSZip from 'jszip';
import { W } from './xml';

const NAMESPACES =
  `xmlns:w="${W}" ` +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** word/document.xml text around `body`. */
export function documentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><w:document ${NAMESPACES}><w:body>${body}</w:body></w:document>`;
}

/** Parses a document whose body is `body`. */
export function parseBody(body: string): XMLDocument {
  return new DOMParser().parseFromString(documentXml(body), 'application/xml');
}

/** The first w:p of a document whose body is `body`. */
export function firstParagraph(body: string): Element {
  const p = parseBody(body).getElementsByTagNameNS(W, 'p')[0];
  if (!p) throw new Error('no w:p in test body');
  return p;
}

/** A styles.xml or numbering.xml part from its inner XML. */
export function partXml(root: 'styles' | 'numbering', inner: string): XMLDocument {
  return new DOMParser().parseFromString(
    `<?xml version="1.0" encoding="UTF-8"?><w:${root} xmlns:w="${W}">${inner}</w:${root}>`,
    'application/xml',
  );
}

/** A minimal .docx package; pass `undefined` for a part to leave it out. */
export async function makeDocx(parts: Record<string, string | undefined>): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const [path, content] of Object.entries(parts)) {
    if (content !== undefined) zip.file(path, content, { createFolders: false });
  }
  return zip.generateAsync({ type: 'uint8array' });
}
