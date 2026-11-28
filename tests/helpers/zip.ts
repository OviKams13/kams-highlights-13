/**
 * Test-only zip helpers, written independently of src/docx/zip.ts so that the
 * tests do not check that module with itself.
 */

export type RawRecord = {
  name: string;
  /** Local header, data and data descriptor. */
  local: Uint8Array;
  /** Central-directory record with the local-header offset set to 0. */
  central: Uint8Array;
  offset: number;
};

const view = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength);

/** The records of a zip without comment or zip64, in central-directory order. */
export function rawRecords(bytes: Uint8Array): RawRecord[] {
  const data = view(bytes);
  const end = bytes.length - 22;
  const count = data.getUint16(end + 10, true);
  const directory = data.getUint32(end + 16, true);
  let at = directory;
  const records = Array.from({ length: count }, () => {
    const nameLength = data.getUint16(at + 28, true);
    const length = 46 + nameLength + data.getUint16(at + 30, true) + data.getUint16(at + 32, true);
    const central = new Uint8Array(bytes.subarray(at, at + length));
    const offset = data.getUint32(at + 42, true);
    view(central).setUint32(42, 0, true);
    at += length;
    const name = new TextDecoder().decode(central.subarray(46, 46 + nameLength));
    return { name, central, offset, local: new Uint8Array() };
  });
  const starts = records.map((r) => r.offset).sort((a, b) => a - b);
  for (const record of records) {
    const next = starts[starts.indexOf(record.offset) + 1] ?? directory;
    record.local = bytes.slice(record.offset, next);
  }
  return records;
}

/** Entry names in the order their data is stored in the file. */
export function physicalOrder(bytes: Uint8Array): string[] {
  return rawRecords(bytes)
    .sort((a, b) => a.offset - b.offset)
    .map((r) => r.name);
}

/**
 * The same package with the local records stored in reverse order (the central
 * directory keeps its order), as in WPS files where the two orders differ.
 */
export function reversePhysicalOrder(bytes: Uint8Array): Uint8Array {
  const records = rawRecords(bytes);
  const out = new Uint8Array(bytes.length);
  let at = 0;
  const offsets = new Map<string, number>();
  for (const record of [...records].sort((a, b) => b.offset - a.offset)) {
    offsets.set(record.name, at);
    out.set(record.local, at);
    at += record.local.length;
  }
  const directory = at;
  for (const record of records) {
    const central = new Uint8Array(record.central);
    view(central).setUint32(42, offsets.get(record.name)!, true);
    out.set(central, at);
    at += central.length;
  }
  out.set(bytes.subarray(bytes.length - 22), at);
  view(out).setUint32(at + 16, directory, true);
  return out;
}
