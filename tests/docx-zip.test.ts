// Phase 5.2: replacing one zip entry without rewriting the others.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { ZipFormatError, readZipLayout, replaceZipEntry } from '@/docx/zip';
import { physicalOrder, rawRecords, reversePhysicalOrder } from './helpers/zip.ts';

const encode = (text: string) => new TextEncoder().encode(text);
const DOCUMENT = 'word/document.xml';
const NAMES = ['[Content_Types].xml', '_rels/', '_rels/.rels', 'word/', DOCUMENT, 'word/a.bin'];

/** A package with folder entries and a stored entry; `streamFiles` adds data descriptors. */
async function sample(streamFiles = false): Promise<Uint8Array> {
  const zip = new JSZip();
  const date = new Date('2020-05-06T07:08:10Z');
  for (const name of NAMES) {
    if (name.endsWith('/')) zip.file(name, null, { dir: true, date, createFolders: false });
    else {
      const compression = name === 'word/a.bin' ? 'STORE' : 'DEFLATE';
      zip.file(name, `old content of ${name} `.repeat(20), {
        date,
        compression,
        createFolders: false,
      });
    }
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', streamFiles });
}

describe.each([
  ['plain records', () => sample()],
  ['data descriptors', () => sample(true)],
  ['reversed physical order', async () => reversePhysicalOrder(await sample())],
])('replaceZipEntry: %s', (_title, build) => {
  it.each([
    ['shorter', 'new'],
    ['longer', Array.from({ length: 4000 }, (_, i) => i).join(' ')],
  ])('replaces only that entry (%s content)', async (_size, content) => {
    const source = await build();
    const out = await replaceZipEntry(source, DOCUMENT, encode(content));

    const before = rawRecords(source);
    const after = rawRecords(out);
    expect(after.map((r) => r.name)).toEqual(NAMES);
    expect(physicalOrder(out)).toEqual(physicalOrder(source));
    for (const [i, record] of before.entries()) {
      if (record.name === DOCUMENT) continue;
      expect(after[i]!.local).toEqual(record.local);
      expect(after[i]!.central).toEqual(record.central);
    }
    // Date, time and attributes of the replaced entry are kept.
    const at = NAMES.indexOf(DOCUMENT);
    expect(after[at]!.central.subarray(12, 16)).toEqual(before[at]!.central.subarray(12, 16));
    expect(after[at]!.central.subarray(34, 42)).toEqual(before[at]!.central.subarray(34, 42));

    const reopened = await JSZip.loadAsync(out, { checkCRC32: true });
    expect(Object.keys(reopened.files)).toEqual(NAMES);
    expect(await reopened.file(DOCUMENT)!.async('string')).toBe(content);
    expect(readZipLayout(out).entries).toHaveLength(NAMES.length);
  });
});

describe('replaceZipEntry: errors', () => {
  it('refuses a missing entry', async () => {
    await expect(replaceZipEntry(await sample(), 'nope.xml', encode(''))).rejects.toThrow(
      ZipFormatError,
    );
  });

  it('refuses bytes that are not a whole zip', async () => {
    const bytes = await sample();
    expect(() => readZipLayout(bytes.subarray(0, bytes.length - 5))).toThrow(ZipFormatError);
    expect(() => readZipLayout(encode('not a zip'))).toThrow(ZipFormatError);
  });
});
