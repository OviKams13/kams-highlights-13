/**
 * Export (phase 5.2, ERD "Export"): the original package with the marks written
 * into word/document.xml. Every other zip record is copied byte for byte (./zip.ts).
 *
 * The open model is not changed: marks are written into a copy of the XML, so a
 * document can be exported again after more taps.
 */
import { rangesFromUnits } from '@/core/marks';
import { DOCUMENT_PART } from './open';
import { collectParagraphs } from './paragraphs';
import type { DocModel } from './types';
import { applyMarks } from './write';
import { replaceZipEntry } from './zip';

/** Default of the `filenameSuffix` setting (ERD A.3). */
export const DEFAULT_SUFFIX = '_highlighted';

export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Marked units per paragraphIndex (ERD "UI state"). A missing paragraph has no highlight. */
export type DocumentMarks = ReadonlyMap<number, readonly number[]>;

/** `<name><suffix>.docx`, where `name` is the original name without its extension. */
export function exportFileName(name: string, suffix: string = DEFAULT_SUFFIX): string {
  return `${name.replace(/\.docx$/i, '')}${suffix}.docx`;
}

/** A copy of the document XML where the darkYellow highlights are exactly `marks`. */
export function writeMarks(
  doc: Pick<DocModel, 'xml' | 'paragraphs'>,
  marks: DocumentMarks,
): XMLDocument {
  const xml = doc.xml.cloneNode(true) as XMLDocument;
  const elements = collectParagraphs(xml);
  for (const { index, words, gaps } of doc.paragraphs) {
    const element = elements[index];
    if (element) applyMarks(element, rangesFromUnits(marks.get(index) ?? [], words, gaps));
  }
  return xml;
}

/**
 * The XML text with the declaration `standalone="yes"`, followed by the same
 * line break as in `source` (WPS writes `\r\n`). XMLSerializer does not give
 * the same declaration in every engine, so the nodes are serialized without it.
 */
export function serializeDocument(xml: XMLDocument, source: string): string {
  const lineBreak = /^\uFEFF?<\?xml\s[^>]*\?>(\s*)/.exec(source)?.[1] ?? '\r\n';
  const serializer = new XMLSerializer();
  const body = Array.from(xml.childNodes)
    // Some engines expose the declaration as a processing instruction.
    .filter((node) => node.nodeName !== 'xml')
    .map((node) => serializer.serializeToString(node))
    .join('');
  return DECLARATION + lineBreak + body;
}

/**
 * The .docx to download: the original file with `marks`, named
 * `<name><suffix>.docx`. A `File` is a `Blob` that carries its name.
 *
 * @throws ZipFormatError when the original package cannot be rewritten.
 */
export async function exportDocx(
  doc: DocModel,
  marks: DocumentMarks,
  suffix: string = DEFAULT_SUFFIX,
): Promise<File> {
  const part = doc.zip.file(DOCUMENT_PART);
  if (!part) throw new Error(`${DOCUMENT_PART} is missing from the open document.`);
  const xml = serializeDocument(writeMarks(doc, marks), await part.async('string'));
  const bytes = await replaceZipEntry(doc.bytes, DOCUMENT_PART, new TextEncoder().encode(xml));
  return new File([bytes], exportFileName(doc.name, suffix), { type: DOCX_MIME });
}
