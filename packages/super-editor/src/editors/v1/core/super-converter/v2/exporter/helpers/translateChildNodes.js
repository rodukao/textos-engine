import { exportSchemaToJson } from '../../../exporter.js';

/**
 * Marks a translated `<w:p>` that only wraps a drawing hoisted out of a paragraph on
 * import. Value: 'next' | 'previous' — the sibling paragraph the drawing belongs to.
 * A symbol, so it never reaches the serialized XML.
 */
export const HOST_PARAGRAPH = Symbol('hostParagraph');

const isParagraph = (node) => node?.name === 'w:p';

/**
 * Put hoisted drawings back inside the paragraph they came from: their runs go to the
 * start of the next paragraph or the end of the previous one, and the wrapper <w:p>
 * is dropped. Without this, each such drawing becomes an extra empty paragraph.
 *
 * @param {Array} nodes Translated sibling nodes (mutated)
 * @returns {Array}
 */
function mergeHoistedDrawings(nodes) {
  // 'next' from the end, so consecutive drawings keep their original order.
  for (let i = nodes.length - 1; i >= 0; i--) {
    const node = nodes[i];
    if (node?.[HOST_PARAGRAPH] !== 'next' || !isParagraph(nodes[i + 1])) continue;
    const target = nodes[i + 1];
    const runs = (node.elements || []).filter((el) => el?.name !== 'w:pPr');
    const elements = target.elements || [];
    const pPrIndex = elements.findIndex((el) => el?.name === 'w:pPr');
    target.elements = [...elements.slice(0, pPrIndex + 1), ...runs, ...elements.slice(pPrIndex + 1)];
    nodes.splice(i, 1);
  }
  for (let i = 0; i < nodes.length; i++) {
    const node = nodes[i];
    if (node?.[HOST_PARAGRAPH] !== 'previous' || !isParagraph(nodes[i - 1])) continue;
    const target = nodes[i - 1];
    const runs = (node.elements || []).filter((el) => el?.name !== 'w:pPr');
    target.elements = [...(target.elements || []), ...runs];
    nodes.splice(i, 1);
    i--;
  }
  return nodes;
}

/**
 * Process child nodes, ignoring any that are not valid
 *
 * @param {import('@converter/exporter').SchemaNode[]} nodes The input nodes
 * @returns {import('@converter/exporter').XmlReadyNode[]} The processed child nodes
 */
export function translateChildNodes(params) {
  const { content: nodes } = params.node;
  if (!nodes) return [];

  const translatedNodes = [];
  const hyperlinkGroup = [];
  let index = 0;
  while (index < nodes.length) {
    const node = nodes[index];
    const linkMark = _isLinkNode(node);

    // Group adjacent hyperlinks together to avoid breaking them up
    if (linkMark) {
      hyperlinkGroup.push(node);
      const nextNode = index + 1 < nodes.length ? nodes[index + 1] : null;
      const nextIsLink = _isLinkNode(nextNode, linkMark);
      if (nextIsLink) {
        index++;
        continue; // Continue to the next iteration to check the next node
      } else {
        let translatedLinkGroup = exportSchemaToJson({
          ...params,
          node: hyperlinkGroup[0],
          extraParams: { ...(params.extraParams || {}), hyperlinkGroup: hyperlinkGroup.slice() },
        });
        if (translatedLinkGroup instanceof Array) translatedNodes.push(...translatedLinkGroup);
        else translatedNodes.push(translatedLinkGroup);
        hyperlinkGroup.length = 0; // Clear the group
        index++;
        continue;
      }
    }

    let translatedNode = exportSchemaToJson({ ...params, node });

    if (translatedNode instanceof Array) translatedNodes.push(...translatedNode);
    else translatedNodes.push(translatedNode);

    index++;
  }

  // Filter out any null nodes
  return mergeHoistedDrawings(translatedNodes.filter((n) => n));
}

/**
 * Check if a node has a link mark, optionally matching a reference mark's attributes
 * @param {import('@converter/exporter').SchemaNode} node The node to check
 * @param {import('@converter/exporter').MarkType|null} [referenceMark] An optional reference mark to match attributes against
 * @returns {import('@converter/exporter').MarkType|null} The link mark if found (and matches), otherwise null
 */
function _isLinkNode(node, referenceMark = null) {
  if (!node || (!node.marks && !node.attrs?.marksAsAttrs)) return null;
  const marks = node.marks || node.attrs.marksAsAttrs;
  const linkMark = marks.find((mark) => mark.type === 'link');
  if (!linkMark) return null;
  if (referenceMark) {
    // If a reference mark is provided, ensure the link attributes match
    return _isSameLinkMark(linkMark.attrs, referenceMark.attrs) ? linkMark : null;
  }
  return linkMark;
}

function _isSameLinkMark(attrsA, attrsB) {
  for (let key of ['anchor', 'docLocation', 'history', 'href', 'rId', 'target']) {
    if (attrsA[key] != attrsB[key]) {
      return false;
    }
  }
  return true;
}
