/**
 * Replacing one entry of a zip package without rewriting the others (phase 5.2).
 *
 * JSZip would regenerate every record and write them in central-directory order,
 * but WPS stores them in another order. So the original bytes are spliced here:
 * every other local record (header, data, data descriptor) and every other
 * central-directory record is copied as is, in its original place. Only the
 * offsets that move because the new entry has another size are updated.
 *
 * Zip layout: APPNOTE.TXT §4.3. zip64 packages are refused (a .docx does not need it).
 */
import JSZip from 'jszip';

const LOCAL_SIGNATURE = 0x04034b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const END_SIGNATURE = 0x06054b50;
const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;

const LOCAL_HEADER_SIZE = 30;
const CENTRAL_HEADER_SIZE = 46;
const END_SIZE = 22;
const MAX_COMMENT = 0xffff;

const DEFLATE = 8;
/** General purpose flags: bits 1–2 (deflate level) and 3 (data descriptor follows). */
const CLEARED_FLAGS = 0b1110;

export class ZipFormatError extends Error {
  constructor(detail: string) {
    super(`Unsupported zip package: ${detail}.`);
    this.name = 'ZipFormatError';
  }
}

export type ZipEntry = {
  name: string;
  /** Position and size of the record in the central directory. */
  centralStart: number;
  centralLength: number;
  /** Position of the local header. The record ends where the next one starts. */
  localStart: number;
  localEnd: number;
  compressedSize: number;
};

export type ZipLayout = {
  /** In central-directory order (the order zip tools list). */
  entries: ZipEntry[];
  directoryStart: number;
  /** Position of the end-of-central-directory record. */
  endStart: number;
};

const viewOf = (bytes: Uint8Array) => new DataView(bytes.buffer, bytes.byteOffset, bytes.length);

function findEnd(view: DataView): number {
  const last = view.byteLength - END_SIZE;
  for (let at = last; at >= Math.max(0, last - MAX_COMMENT); at -= 1) {
    if (
      view.getUint32(at, true) === END_SIGNATURE &&
      at + END_SIZE + view.getUint16(at + 20, true) === view.byteLength
    ) {
      return at;
    }
  }
  throw new ZipFormatError('no end of central directory');
}

/** Reads where every record of the package is. */
export function readZipLayout(bytes: Uint8Array): ZipLayout {
  const view = viewOf(bytes);
  const endStart = findEnd(view);
  const count = view.getUint16(endStart + 10, true);
  const directorySize = view.getUint32(endStart + 12, true);
  const directoryStart = view.getUint32(endStart + 16, true);
  const zip64 =
    count === 0xffff ||
    directoryStart === 0xffffffff ||
    (endStart >= 20 && view.getUint32(endStart - 20, true) === ZIP64_LOCATOR_SIGNATURE);
  if (zip64) throw new ZipFormatError('zip64');
  if (directoryStart + directorySize !== endStart) {
    throw new ZipFormatError('unexpected data around the central directory');
  }

  const decoder = new TextDecoder();
  const entries: ZipEntry[] = [];
  let at = directoryStart;
  for (let i = 0; i < count; i += 1) {
    if (at + CENTRAL_HEADER_SIZE > endStart || view.getUint32(at, true) !== CENTRAL_SIGNATURE) {
      throw new ZipFormatError('damaged central directory');
    }
    const nameLength = view.getUint16(at + 28, true);
    const centralLength =
      CENTRAL_HEADER_SIZE +
      nameLength +
      view.getUint16(at + 30, true) +
      view.getUint16(at + 32, true);
    const localStart = view.getUint32(at + 42, true);
    if (localStart >= directoryStart || view.getUint32(localStart, true) !== LOCAL_SIGNATURE) {
      throw new ZipFormatError('damaged local header');
    }
    const nameStart = at + CENTRAL_HEADER_SIZE;
    entries.push({
      name: decoder.decode(bytes.subarray(nameStart, nameStart + nameLength)),
      centralStart: at,
      centralLength,
      localStart,
      localEnd: directoryStart,
      compressedSize: view.getUint32(at + 20, true),
    });
    at += centralLength;
  }

  const physical = [...entries].sort((a, b) => a.localStart - b.localStart);
  physical.forEach((entry, i) => {
    entry.localEnd = physical[i + 1]?.localStart ?? directoryStart;
  });
  return { entries, directoryStart, endStart };
}

/** Length of the local header (fixed part, name and extra field). */
function localHeaderLength(view: DataView, start: number): number {
  return LOCAL_HEADER_SIZE + view.getUint16(start + 26, true) + view.getUint16(start + 28, true);
}

/** Deflates `content` with JSZip and returns the compressed data and its checksum. */
async function deflate(content: Uint8Array) {
  const zip = new JSZip();
  zip.file('part', content);
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const entry = readZipLayout(bytes).entries[0]!;
  const view = viewOf(bytes);
  const dataStart = entry.localStart + localHeaderLength(view, entry.localStart);
  return {
    data: bytes.subarray(dataStart, dataStart + entry.compressedSize),
    crc32: view.getUint32(entry.centralStart + 16, true),
  };
}

/** Writes method, CRC-32 and sizes into a header; `at` is the position of the flags field. */
function patchHeader(
  view: DataView,
  at: number,
  values: { crc32: number; compressedSize: number; size: number },
): void {
  view.setUint16(at - 2, Math.max(view.getUint16(at - 2, true), 20), true); // version needed
  view.setUint16(at, view.getUint16(at, true) & ~CLEARED_FLAGS, true);
  view.setUint16(at + 2, DEFLATE, true);
  view.setUint32(at + 8, values.crc32, true);
  view.setUint32(at + 12, values.compressedSize, true);
  view.setUint32(at + 16, values.size, true);
}

function concat(parts: readonly Uint8Array[]): Uint8Array<ArrayBuffer> {
  const out = new Uint8Array(parts.reduce((n, part) => n + part.length, 0));
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

/**
 * A copy of the package `source` where the content of the entry `name` is
 * `content`. No entry is added, removed or moved; the replaced entry keeps its
 * name, date and attributes.
 *
 * @throws ZipFormatError when the package has no such entry or cannot be rewritten.
 */
export async function replaceZipEntry(
  source: Uint8Array,
  name: string,
  content: Uint8Array,
): Promise<Uint8Array<ArrayBuffer>> {
  const { entries, directoryStart, endStart } = readZipLayout(source);
  // With duplicate names, the last one is the one JSZip (and so the reader) uses.
  const target = entries.findLast((entry) => entry.name === name);
  if (!target) throw new ZipFormatError(`no entry ${name}`);

  const { data, crc32 } = await deflate(content);
  const values = { crc32, compressedSize: data.length, size: content.length };

  const headerLength = localHeaderLength(viewOf(source), target.localStart);
  const header = source.slice(target.localStart, target.localStart + headerLength);
  patchHeader(viewOf(header), 6, values);
  const record = concat([header, data]);
  const shift = record.length - (target.localEnd - target.localStart);

  const directory = entries.map((entry) => {
    const copy = source.slice(entry.centralStart, entry.centralStart + entry.centralLength);
    const view = viewOf(copy);
    if (entry === target) patchHeader(view, 8, values);
    if (entry.localStart > target.localStart) view.setUint32(42, entry.localStart + shift, true);
    return copy;
  });

  const end = source.slice(endStart);
  viewOf(end).setUint32(16, directoryStart + shift, true);

  return concat([
    source.subarray(0, target.localStart),
    record,
    source.subarray(target.localEnd, directoryStart),
    ...directory,
    end,
  ]);
}
