import { describe, it, expect, vi } from 'vitest';

// Each schema node "translates" to the XML node stored on it, so the tests control
// exactly what translateChildNodes receives back from the exporter.
vi.mock('../../../exporter.js', () => ({
  exportSchemaToJson: ({ node }) => node.xml,
}));

import { translateChildNodes, HOST_PARAGRAPH } from './translateChildNodes.js';

const run = (text) => ({ name: 'w:r', elements: [{ name: 'w:t', elements: [{ type: 'text', text }] }] });
const pPr = { name: 'w:pPr', elements: [{ name: 'w:jc', attributes: { 'w:val': 'both' } }] };
const paragraph = (...elements) => ({ xml: { name: 'w:p', elements } });
const hoisted = (host, label) => {
  const xml = { name: 'w:p', elements: [run(label)] };
  xml[HOST_PARAGRAPH] = host;
  return { xml };
};
const translate = (...content) => translateChildNodes({ node: { content } });
const texts = (p) => p.elements.map((el) => (el.name === 'w:r' ? el.elements[0].elements[0].text : el.name));

describe('translateChildNodes hoisted drawings', () => {
  it('moves a drawing hoisted from the start of a paragraph back into it, after w:pPr', () => {
    const out = translate(hoisted('next', 'box'), paragraph(pPr, run('Body')));
    expect(out).toHaveLength(1);
    expect(texts(out[0])).toEqual(['w:pPr', 'box', 'Body']);
  });

  it('moves a drawing hoisted from the end of a paragraph back into it', () => {
    const out = translate(paragraph(pPr, run('Body')), hoisted('previous', 'box'));
    expect(out).toHaveLength(1);
    expect(texts(out[0])).toEqual(['w:pPr', 'Body', 'box']);
  });

  it('keeps the original order of several drawings that shared one paragraph', () => {
    const out = translate(hoisted('next', 'a'), hoisted('next', 'b'), paragraph(run('Body')), hoisted('previous', 'c'));
    expect(out).toHaveLength(1);
    expect(texts(out[0])).toEqual(['a', 'b', 'Body', 'c']);
  });

  it('keeps the wrapper paragraph when there is no neighbouring paragraph to host the drawing', () => {
    const table = { xml: { name: 'w:tbl', elements: [] } };
    const out = translate(hoisted('next', 'box'), table);
    expect(out.map((n) => n.name)).toEqual(['w:p', 'w:tbl']);
  });

  it('never leaks the marker into the XML object keys', () => {
    const out = translate(hoisted('next', 'box'), paragraph(run('Body')));
    expect(Object.keys(out[0])).toEqual(['name', 'elements']);
  });
});
