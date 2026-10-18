import {
  cloneXmlElement,
  isSectPrElement,
  readSectPrMargins,
  readSectPrPageSetup,
  type XmlElement,
} from './sections-xml.js';

/**
 * Converter fields that mirror the body-level `w:sectPr`.
 *
 * Layout reads `converter.bodySectPr` (via `editor.getJSON()`) and
 * `converter.pageStyles`, while the document history only tracks
 * `doc.attrs.bodySectPr`. Every write to the body section must keep these
 * copies in step with the doc attribute.
 */
export interface ConverterWithBodySection {
  bodySectPr?: unknown;
  savedTagsToRestore?: Array<Record<string, unknown>>;
  pageStyles?: {
    pageSize?: { width?: number; height?: number };
    pageMargins?: Record<string, number | undefined>;
    [key: string]: unknown;
  } | null;
}

/** Copy a body sectPr into the converter (bodySectPr, saved w:body tag and pageStyles). */
export function syncConverterBodySectPr(
  converter: ConverterWithBodySection | null | undefined,
  sectPr: XmlElement,
): void {
  if (!converter) return;
  converter.bodySectPr = cloneXmlElement(sectPr);

  const savedBodyNode = converter.savedTagsToRestore?.find((entry) => entry?.name === 'w:body');
  if (savedBodyNode && Array.isArray(savedBodyNode.elements)) {
    const preservedChildren = (savedBodyNode.elements as Array<Record<string, unknown>>).filter(
      (entry) => entry?.name !== 'w:sectPr',
    );
    preservedChildren.push(cloneXmlElement(sectPr) as unknown as Record<string, unknown>);
    savedBodyNode.elements = preservedChildren;
  }

  const margins = readSectPrMargins(sectPr);
  const pageSetup = readSectPrPageSetup(sectPr);
  if (!converter.pageStyles) converter.pageStyles = {};
  if (!converter.pageStyles.pageSize) converter.pageStyles.pageSize = {};
  if (pageSetup?.width !== undefined) converter.pageStyles.pageSize.width = pageSetup.width;
  if (pageSetup?.height !== undefined) converter.pageStyles.pageSize.height = pageSetup.height;
  if (!converter.pageStyles.pageMargins) converter.pageStyles.pageMargins = {};
  const pageMargins = converter.pageStyles.pageMargins;
  if (margins.top !== undefined) pageMargins.top = margins.top;
  if (margins.right !== undefined) pageMargins.right = margins.right;
  if (margins.bottom !== undefined) pageMargins.bottom = margins.bottom;
  if (margins.left !== undefined) pageMargins.left = margins.left;
  if (margins.header !== undefined) pageMargins.header = margins.header;
  if (margins.footer !== undefined) pageMargins.footer = margins.footer;
  if (margins.gutter !== undefined) pageMargins.gutter = margins.gutter;
}

/**
 * Re-align the converter with `doc.attrs.bodySectPr` after a transaction that
 * changed the attribute without going through a section mutation (undo, redo,
 * remote steps). Returns true when the converter was updated.
 */
export function resyncConverterBodySectPrFromDoc(
  converter: ConverterWithBodySection | null | undefined,
  previousBodySectPr: unknown,
  nextBodySectPr: unknown,
): boolean {
  if (!converter) return false;
  if (previousBodySectPr === nextBodySectPr) return false;
  if (!isSectPrElement(nextBodySectPr)) return false;
  if (JSON.stringify(converter.bodySectPr ?? null) === JSON.stringify(nextBodySectPr)) return false;
  syncConverterBodySectPr(converter, nextBodySectPr);
  return true;
}
