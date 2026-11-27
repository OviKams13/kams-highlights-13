import { describe, expect, it } from 'vitest';
import { buildBlocks } from './blocks';
import { collectParagraphs } from './paragraphs';
import { parseBody } from './testing';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const tc = (content: string, tcPr = '') => `<w:tc><w:tcPr>${tcPr}</w:tcPr>${content}</w:tc>`;

const blocksOf = (body: string) => {
  const xml = parseBody(body);
  return buildBlocks(xml, collectParagraphs(xml));
};

describe('buildBlocks', () => {
  it('lists body paragraphs in order and ignores sectPr', () => {
    expect(blocksOf(p('a') + p('b') + '<w:sectPr/>')).toEqual([
      { kind: 'paragraph', paragraphIndex: 0 },
      { kind: 'paragraph', paragraphIndex: 1 },
    ]);
  });

  it('reads gridSpan, vMerge (restart / bare = continue) and nested tables', () => {
    const nested = `<w:tbl><w:tr>${tc(p('in'))}</w:tr></w:tbl>`;
    const blocks = blocksOf(
      p('before') +
        '<w:tbl><w:tblPr/><w:tblGrid/>' +
        `<w:tr>${tc(p('span'), '<w:gridSpan w:val="2"/>')}${tc(p('r'), '<w:vMerge w:val="restart"/>')}</w:tr>` +
        `<w:tr>${tc(nested + p('after'))}${tc(p('x'))}${tc('<w:p/>', '<w:vMerge/>')}</w:tr>` +
        '</w:tbl>' +
        p('end'),
    );
    expect(blocks).toEqual([
      { kind: 'paragraph', paragraphIndex: 0 },
      {
        kind: 'table',
        rows: [
          {
            cells: [
              { blocks: [{ kind: 'paragraph', paragraphIndex: 1 }], gridSpan: 2 },
              {
                blocks: [{ kind: 'paragraph', paragraphIndex: 2 }],
                gridSpan: 1,
                vMerge: 'restart',
              },
            ],
          },
          {
            cells: [
              {
                blocks: [
                  {
                    kind: 'table',
                    rows: [
                      {
                        cells: [
                          { blocks: [{ kind: 'paragraph', paragraphIndex: 3 }], gridSpan: 1 },
                        ],
                      },
                    ],
                  },
                  { kind: 'paragraph', paragraphIndex: 4 },
                ],
                gridSpan: 1,
              },
              { blocks: [{ kind: 'paragraph', paragraphIndex: 5 }], gridSpan: 1 },
              {
                blocks: [{ kind: 'paragraph', paragraphIndex: 6 }],
                gridSpan: 1,
                vMerge: 'continue',
              },
            ],
          },
        ],
      },
      { kind: 'paragraph', paragraphIndex: 7 },
    ]);
  });

  it('opens block-level content controls (w:sdt) and customXml', () => {
    const blocks = blocksOf(
      `<w:sdt><w:sdtPr/><w:sdtContent>${p('a')}<w:tbl><w:tr><w:sdt><w:sdtContent>${tc(p('b'))}</w:sdtContent></w:sdt></w:tr></w:tbl></w:sdtContent></w:sdt>` +
        `<w:customXml w:element="x">${p('c')}</w:customXml>`,
    );
    expect(blocks).toEqual([
      { kind: 'paragraph', paragraphIndex: 0 },
      {
        kind: 'table',
        rows: [{ cells: [{ blocks: [{ kind: 'paragraph', paragraphIndex: 1 }], gridSpan: 1 }] }],
      },
      { kind: 'paragraph', paragraphIndex: 2 },
    ]);
  });
});
