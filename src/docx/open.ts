/**
 * Opening a .docx (phase 4.1): JSZip load, SHA-256 fingerprint, namespace-aware
 * parse of word/document.xml, word/styles.xml and word/numbering.xml, then the
 * full in-memory model of ERD Part B.
 *
 * Nothing here changes the package: export (Phase 5) rewrites word/document.xml only.
 */
import JSZip from 'jszip';
import { tokenize } from '@/core/tokenize';
import { buildBlocks } from './blocks';
import {
  buildRunViews,
  createListCounter,
  paragraphStyleId,
  parseNumbering,
  parseStyles,
  resolveParagraphProps,
  type NumberingTable,
  type StyleTable,
} from './format';
import { collectParagraphs, flattenParagraph } from './paragraphs';
import type { DocModel, ParagraphModel } from './types';
import { XmlParseError, parseXml } from './xml';

export const DOCUMENT_PART = 'word/document.xml';
export const STYLES_PART = 'word/styles.xml';
export const NUMBERING_PART = 'word/numbering.xml';
export const THEME_PART = 'word/theme/theme1.xml';

export type DocxOpenErrorCode = 'legacy-format' | 'not-zip' | 'no-document' | 'bad-xml';

const MESSAGES: Record<DocxOpenErrorCode, string> = {
  'legacy-format':
    'This is an old .doc or .wps file. In WPS, open it and use "Save as" → .docx, then import the .docx.',
  'not-zip': 'This file is not a .docx document. In WPS, use "Save as" → .docx, then import it.',
  'no-document':
    'This file has no Word document inside (word/document.xml is missing). In WPS, use "Save as" → .docx.',
  'bad-xml': 'This .docx is damaged: its text could not be read. Try saving it again from WPS.',
};

/** A file that cannot be opened, with a message for the user. */
export class DocxOpenError extends Error {
  readonly code: DocxOpenErrorCode;
  constructor(code: DocxOpenErrorCode, detail?: string) {
    super(MESSAGES[code]);
    this.name = 'DocxOpenError';
    this.code = code;
    if (detail) this.cause = detail;
  }
}

/** OLE2 compound file signature: .doc, .wps, .xls… */
const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
/** Zip local file header ("PK\3\4") or empty archive ("PK\5\6"). */
const ZIP_SIGNATURES = [
  [0x50, 0x4b, 0x03, 0x04],
  [0x50, 0x4b, 0x05, 0x06],
];

const startsWith = (bytes: Uint8Array, signature: readonly number[]) =>
  bytes.length >= signature.length && signature.every((b, i) => bytes[i] === b);

/** Checks the file signature before unzipping, for a clear error message. */
export function sniffFormat(bytes: Uint8Array, name = ''): 'zip' | DocxOpenErrorCode {
  if (ZIP_SIGNATURES.some((s) => startsWith(bytes, s))) return 'zip';
  if (startsWith(bytes, OLE_SIGNATURE) || /\.(doc|wps|dot|wpt)$/i.test(name)) {
    return 'legacy-format';
  }
  return 'not-zip';
}

/** SHA-256 of the bytes, lowercase hex (D-09, ERD A.1 `id`). */
export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  // Copy into a plain ArrayBuffer-backed view (subtle.digest refuses shared buffers).
  const digest = await globalThis.crypto.subtle.digest('SHA-256', new Uint8Array(bytes));
  return Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, '0')).join('');
}

/** Realm-safe (no `instanceof`): bytes may come from another window, a worker or Node. */
async function toBytes(input: Blob | ArrayBuffer | Uint8Array): Promise<Uint8Array> {
  if (ArrayBuffer.isView(input))
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  if ('arrayBuffer' in input) return new Uint8Array(await input.arrayBuffer());
  return new Uint8Array(input);
}

async function readPart(zip: JSZip, path: string): Promise<XMLDocument | undefined> {
  const file = zip.file(path);
  if (!file) return undefined;
  try {
    return parseXml(await file.async('string'), path);
  } catch (error) {
    if (error instanceof XmlParseError) throw new DocxOpenError('bad-xml', error.message);
    throw error;
  }
}

/** Builds every ParagraphModel, in document order (list numbering depends on that order). */
export function buildParagraphModels(
  xml: XMLDocument,
  styles: StyleTable,
  numbering: NumberingTable,
): { elements: Element[]; paragraphs: ParagraphModel[] } {
  const elements = collectParagraphs(xml);
  const counter = createListCounter(numbering);
  const paragraphs = elements.map((element, index): ParagraphModel => {
    const { text, charMap, hardBreaks } = flattenParagraph(element);
    const { words, gaps } = tokenize(text, hardBreaks);
    const props = resolveParagraphProps(styles, numbering, counter, element);
    const runs = buildRunViews(styles, paragraphStyleId(styles, element), charMap);
    return { index, element, text, charMap, hardBreaks, words, gaps, props, runs };
  });
  return { elements, paragraphs };
}

/**
 * Opens a .docx and builds its in-memory model.
 *
 * @throws DocxOpenError for a .doc/.wps file, a file that is not a zip, a zip
 *   without word/document.xml, or XML that is not well-formed.
 */
export async function openDocx(
  input: Blob | ArrayBuffer | Uint8Array,
  name: string,
): Promise<DocModel> {
  const bytes = await toBytes(input);
  const format = sniffFormat(bytes, name);
  if (format !== 'zip') throw new DocxOpenError(format);

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch (error) {
    throw new DocxOpenError('not-zip', error instanceof Error ? error.message : String(error));
  }

  const [xml, stylesXml, numberingXml, themeXml, id] = await Promise.all([
    readPart(zip, DOCUMENT_PART),
    readPart(zip, STYLES_PART),
    readPart(zip, NUMBERING_PART),
    // The theme only gives fallback font names: a damaged theme is ignored.
    readPart(zip, THEME_PART).catch(() => undefined),
    sha256Hex(bytes),
  ]);
  if (!xml) throw new DocxOpenError('no-document');

  const styles = parseStyles(stylesXml, themeXml);
  const numbering = parseNumbering(numberingXml);
  const { elements, paragraphs } = buildParagraphModels(xml, styles, numbering);
  const blocks = buildBlocks(xml, elements);

  return { id, name, size: bytes.length, zip, xml, styles, numbering, blocks, paragraphs };
}
