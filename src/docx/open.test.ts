import { describe, expect, it } from 'vitest';
import { DocxOpenError, openDocx, sha256Hex, sniffFormat } from './open';
import { documentXml, makeDocx } from './testing';

const BODY =
  '<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r>' +
  '<w:r><w:rPr><w:highlight w:val="darkYellow"/></w:rPr><w:t>world</w:t></w:r></w:p>';

async function openError(input: Uint8Array, name = 'file.docx'): Promise<DocxOpenError> {
  try {
    await openDocx(input, name);
  } catch (error) {
    if (error instanceof DocxOpenError) return error;
    throw error;
  }
  throw new Error('expected openDocx to fail');
}

describe('sha256Hex', () => {
  it('matches the FIPS 180-2 test vector', async () => {
    expect(await sha256Hex(new TextEncoder().encode('abc'))).toBe(
      'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
    );
  });
});

describe('sniffFormat', () => {
  it('recognises zip, OLE (.doc/.wps) and other files', () => {
    expect(sniffFormat(Uint8Array.of(0x50, 0x4b, 0x03, 0x04, 0))).toBe('zip');
    expect(sniffFormat(Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1))).toBe(
      'legacy-format',
    );
    expect(sniffFormat(new TextEncoder().encode('{\\rtf1'), 'old.wps')).toBe('legacy-format');
    expect(sniffFormat(new TextEncoder().encode('%PDF-1.7'), 'a.pdf')).toBe('not-zip');
    expect(sniffFormat(new Uint8Array(0))).toBe('not-zip');
  });
});

describe('openDocx', () => {
  it('opens a minimal package and builds the model', async () => {
    const bytes = await makeDocx({ 'word/document.xml': documentXml(BODY) });
    const doc = await openDocx(bytes, 'hello.docx');
    expect(doc.name).toBe('hello.docx');
    expect(doc.size).toBe(bytes.length);
    expect(doc.id).toBe(await sha256Hex(bytes));
    expect(doc.id).toMatch(/^[0-9a-f]{64}$/);
    expect(doc.paragraphs).toHaveLength(1);
    expect(doc.paragraphs[0]).toMatchObject({ index: 0, text: 'Hello world' });
    expect(doc.blocks).toEqual([{ kind: 'paragraph', paragraphIndex: 0 }]);
    // No styles.xml / numbering.xml: empty tables, not an error.
    expect(doc.styles.styles.size).toBe(0);
    expect(doc.numbering.nums.size).toBe(0);
  });

  it('accepts an ArrayBuffer or a Blob and gives the same id', async () => {
    const bytes = await makeDocx({ 'word/document.xml': documentXml(BODY) });
    const buffer = bytes.slice().buffer;
    const fromBuffer = await openDocx(buffer, 'a.docx');
    const fromBlob = await openDocx(new Blob([buffer]), 'a.docx');
    expect(fromBuffer.id).toBe(fromBlob.id);
    expect(fromBlob.paragraphs[0]?.text).toBe('Hello world');
  });

  it('rejects a .doc / .wps file with a clear message', async () => {
    const ole = Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0);
    const error = await openError(ole, 'cv.doc');
    expect(error.code).toBe('legacy-format');
    expect(error.message).toMatch(/\.doc or \.wps/);
    expect((await openError(ole, 'cv.wps')).code).toBe('legacy-format');
  });

  it('rejects a file that is not a zip', async () => {
    const error = await openError(new TextEncoder().encode('just some text'), 'notes.docx');
    expect(error.code).toBe('not-zip');
    expect(error.message).toMatch(/not a \.docx/);
  });

  it('rejects a damaged zip', async () => {
    const error = await openError(Uint8Array.of(0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4, 5));
    expect(error.code).toBe('not-zip');
  });

  it('rejects a zip without word/document.xml (e.g. an .odt)', async () => {
    const bytes = await makeDocx({
      mimetype: 'application/vnd.oasis.opendocument.text',
      'content.xml': '<x/>',
    });
    const error = await openError(bytes, 'notes.odt');
    expect(error.code).toBe('no-document');
    expect(error.message).toMatch(/word\/document\.xml/);
  });

  it('rejects XML that is not well-formed', async () => {
    const bytes = await makeDocx({ 'word/document.xml': '<w:document><w:body>' });
    expect((await openError(bytes)).code).toBe('bad-xml');
  });
});
