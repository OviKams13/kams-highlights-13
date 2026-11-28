import { describe, expect, it } from 'vitest';
import { DOCX_MIME, exportDocx, exportFileName, serializeDocument, writeMarks } from './export';
import { openDocx } from './open';
import { readDocumentMarks } from './read-marks';
import { documentXml, makeDocx, parseBody } from './testing';
import { W } from './xml';

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const BODY = '<w:p><w:r><w:t>one two three</w:t></w:r></w:p><w:p/>';

const open = async (name = 'answer.docx') =>
  openDocx(await makeDocx({ 'word/document.xml': documentXml(BODY) }), name);

describe('exportFileName', () => {
  it('is <name><suffix>.docx', () => {
    expect(exportFileName('answer.docx')).toBe('answer_highlighted.docx');
    expect(exportFileName('My answer.v2.DOCX', ' (marked)')).toBe('My answer.v2 (marked).docx');
    expect(exportFileName('answer', '')).toBe('answer.docx');
  });
});

describe('serializeDocument', () => {
  const xml = parseBody('<w:p/>');

  it('writes standalone="yes" and keeps the line break of the source', () => {
    for (const lineBreak of ['\r\n', '\n', '']) {
      const out = serializeDocument(xml, `${DECLARATION}${lineBreak}<w:document/>`);
      expect(out.startsWith(`${DECLARATION}${lineBreak}<w:document `)).toBe(true);
      expect(out.match(/<\?xml /g)).toHaveLength(1);
    }
  });

  it('adds the declaration when the source has none', () => {
    expect(serializeDocument(xml, '<w:document/>').startsWith(`${DECLARATION}\r\n<`)).toBe(true);
  });
});

describe('exportDocx', () => {
  it('returns a named .docx Blob', async () => {
    const file = await exportDocx(await open(), new Map(), '_x');
    expect(file).toBeInstanceOf(Blob);
    expect(file.name).toBe('answer_x.docx');
    expect(file.type).toBe(DOCX_MIME);
  });

  it('does not change the open model, so it can export again', async () => {
    const doc = await open();
    const before = new XMLSerializer().serializeToString(doc.xml);
    const bytes = new Uint8Array(doc.bytes);

    const marked = await openDocx(await exportDocx(doc, new Map([[0, [2]]])), 'a.docx');
    const other = await openDocx(await exportDocx(doc, new Map([[0, [0, 1, 2]]])), 'b.docx');

    expect(new XMLSerializer().serializeToString(doc.xml)).toBe(before);
    expect(doc.bytes).toEqual(bytes);
    expect(readDocumentMarks(marked).get(0)?.units).toEqual([2]);
    expect(readDocumentMarks(other).get(0)?.units).toEqual([0, 1, 2]);
  });

  it('keeps its own copy of the input bytes', async () => {
    const input = await makeDocx({ 'word/document.xml': documentXml(BODY) });
    const doc = await openDocx(input, 'answer.docx');
    input.fill(0);
    const out = await openDocx(await exportDocx(doc, new Map()), 'a.docx');
    expect(out.paragraphs[0]?.text).toBe('one two three');
  });

  it('ignores marks for paragraphs and units that do not exist', async () => {
    const doc = await open();
    const xml = writeMarks(
      doc,
      new Map([
        [0, [99]],
        [1, [0]],
        [7, [0]],
      ]),
    );
    expect(xml.getElementsByTagNameNS(W, 'highlight')).toHaveLength(0);
  });
});
