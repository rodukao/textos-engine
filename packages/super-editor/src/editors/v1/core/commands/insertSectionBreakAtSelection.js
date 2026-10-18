import { TextSelection } from 'prosemirror-state';
import { canSplit } from 'prosemirror-transform';
import { updateSectionMargins, getSectPrMargins } from '@converter/section-properties.js';
import { Attribute } from '../Attribute.js';

const SECTION_BREAK_TYPES = ['nextPage', 'continuous', 'evenPage', 'oddPage'];

const cloneJson = (value) => JSON.parse(JSON.stringify(value));

const isSectPr = (value) => Boolean(value) && typeof value === 'object' && value.name === 'w:sectPr';

/** A paragraph sectPr only ends a section when it carries properties (see shouldIgnoreSectionBreak). */
const readParagraphSectPr = (node) => {
  const sectPr = node?.attrs?.paragraphProperties?.sectPr;
  if (!isSectPr(sectPr)) return null;
  return Array.isArray(sectPr.elements) && sectPr.elements.length > 0 ? sectPr : null;
};

const buildSectionMargins = (sectPr) => {
  const margins = getSectPrMargins(sectPr) || {};
  return {
    top: margins.top ?? null,
    right: margins.right ?? null,
    bottom: margins.bottom ?? null,
    left: margins.left ?? null,
    header: margins.header ?? null,
    footer: margins.footer ?? null,
  };
};

/**
 * Write the section-start type. `nextPage` is the OOXML default, so it is
 * written by dropping `<w:type>` (as Word does) instead of an explicit value.
 */
const writeSectionType = (sectPr, breakType) => {
  const elements = (sectPr.elements ?? []).filter((el) => el?.name !== 'w:type');
  if (breakType !== 'nextPage') {
    elements.unshift({ type: 'element', name: 'w:type', attributes: { 'w:val': breakType } });
  }
  sectPr.elements = elements;
};

const readSectionType = (sectPr) => {
  const typeEl = sectPr?.elements?.find((el) => el?.name === 'w:type');
  return typeEl?.attributes?.['w:val'] ?? 'nextPage';
};

/**
 * Find the sectPr of the section that contains the top-level node at `index`:
 * the first paragraph from there on that ends a section, else the body sectPr.
 */
const findGoverningSection = (doc, index, bodySectPr) => {
  let pos = 0;
  for (let i = 0; i < doc.childCount; i += 1) {
    const child = doc.child(i);
    if (i >= index && child.type.name === 'paragraph') {
      const sectPr = readParagraphSectPr(child);
      if (sectPr) return { kind: 'paragraph', pos, node: child, sectPr };
    }
    pos += child.nodeSize;
  }
  return isSectPr(bodySectPr) ? { kind: 'body', sectPr: bodySectPr } : null;
};

/**
 * Insert a section break at the selection, as Word's "Breaks > Section Breaks".
 *
 * The paragraph is split at the cursor. The text before the cursor ends the
 * current section: its paragraph gets a copy of the current section's sectPr
 * (OOXML end-tagged semantics). The text after the cursor starts the new
 * section, whose sectPr records how it begins (`breakType`, default `nextPage`).
 *
 * Optionally applies header/footer distances (inches) to the ended section.
 *
 * @param {{ breakType?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage', headerInches?: number, footerInches?: number }} params
 * @param {string} [params.breakType] - How the new section starts (default 'nextPage')
 * @param {number} [params.headerInches] - Distance from page top to content area (must be >= 0)
 * @param {number} [params.footerInches] - Distance from page bottom to content area (must be >= 0)
 * @returns {import('./types/index.js').Command}
 */
export const insertSectionBreakAtSelection =
  ({ breakType = 'nextPage', headerInches, footerInches } = {}) =>
  ({ tr, state, editor, dispatch }) => {
    if (!state || !editor) {
      console.warn('[insertSectionBreakAtSelection] Missing state or editor');
      return false;
    }

    if (!SECTION_BREAK_TYPES.includes(breakType)) {
      console.warn('[insertSectionBreakAtSelection] Unknown breakType:', breakType);
      return false;
    }
    if (typeof headerInches === 'number' && headerInches < 0) {
      console.warn('[insertSectionBreakAtSelection] headerInches must be >= 0, got:', headerInches);
      return false;
    }
    if (typeof footerInches === 'number' && footerInches < 0) {
      console.warn('[insertSectionBreakAtSelection] footerInches must be >= 0, got:', footerInches);
      return false;
    }

    if (!tr.selection.empty) tr.deleteSelection();

    const { $from } = tr.selection;
    // Sections are delimited by body-level paragraphs only.
    let paragraphDepth = -1;
    for (let d = $from.depth; d > 0; d -= 1) {
      if ($from.node(d).type.name === 'paragraph') {
        paragraphDepth = d;
        break;
      }
    }
    if (paragraphDepth !== 1) {
      console.warn('[insertSectionBreakAtSelection] No body paragraph found at selection');
      return false;
    }

    const paragraph = $from.node(paragraphDepth);
    const paragraphPos = $from.before(paragraphDepth);
    const bodySectPr = tr.doc.attrs?.bodySectPr ?? editor.converter?.bodySectPr ?? null;
    const governing = findGoverningSection(tr.doc, $from.index(0), bodySectPr);
    const governedByThisParagraph = governing?.kind === 'paragraph' && governing.node === paragraph;

    // The ended section keeps the current section's properties.
    const endedSectPr = governing ? cloneJson(governing.sectPr) : { type: 'element', name: 'w:sectPr', elements: [] };
    const updates = {};
    if (headerInches != null) updates.headerInches = headerInches;
    if (footerInches != null) updates.footerInches = footerInches;
    if (Object.keys(updates).length) {
      updateSectionMargins({ type: 'sectPr', sectPr: endedSectPr }, updates);
    }

    // The new section starts after the break: its sectPr carries the break type.
    let followingSectPr = governing ? governing.sectPr : null;
    if (followingSectPr && readSectionType(followingSectPr) !== breakType) {
      followingSectPr = cloneJson(followingSectPr);
      writeSectionType(followingSectPr, breakType);
    }

    // Second half: a continuation of the paragraph (fresh ids). If this
    // paragraph already ended a section, that sectPr moves along with it.
    const paraProps = paragraph.attrs?.paragraphProperties || {};
    const splitAttrs = Attribute.getSplittedAttributes(
      editor.extensionService?.attributes ?? [],
      paragraph.type.name,
      paragraph.attrs,
    );
    const secondProps = { ...(splitAttrs.paragraphProperties ?? paraProps) };
    delete secondProps.sectPr;
    const secondAttrs = { ...splitAttrs, paragraphProperties: secondProps };
    delete secondAttrs.pageBreakSource;
    delete secondAttrs.sectionMargins;
    const ownSectPr = readParagraphSectPr(paragraph);
    if (ownSectPr) {
      const carried = governedByThisParagraph ? followingSectPr : ownSectPr;
      secondProps.sectPr = carried;
      secondAttrs.pageBreakSource = 'sectPr';
      secondAttrs.sectionMargins = buildSectionMargins(carried);
    }

    const splitDepth = $from.depth - paragraphDepth + 1;
    const typesAfter = [{ type: paragraph.type, attrs: secondAttrs }];
    for (let d = paragraphDepth + 1; d <= $from.depth; d += 1) typesAfter.push(null);
    if (!canSplit(tr.doc, $from.pos, splitDepth, typesAfter)) {
      console.warn('[insertSectionBreakAtSelection] Selection cannot be split');
      return false;
    }

    if (!dispatch) return true;

    tr.split($from.pos, splitDepth, typesAfter);
    tr.setNodeMarkup(
      paragraphPos,
      undefined,
      {
        ...paragraph.attrs,
        paragraphProperties: { ...paraProps, sectPr: endedSectPr },
        pageBreakSource: 'sectPr',
        sectionMargins: buildSectionMargins(endedSectPr),
      },
      paragraph.marks,
    );

    if (governing && !governedByThisParagraph && followingSectPr !== governing.sectPr) {
      if (governing.kind === 'body') {
        tr.setDocAttribute('bodySectPr', followingSectPr);
      } else {
        const pos = tr.mapping.map(governing.pos);
        const node = tr.doc.nodeAt(pos);
        if (node) {
          tr.setNodeMarkup(
            pos,
            undefined,
            {
              ...node.attrs,
              paragraphProperties: { ...(node.attrs.paragraphProperties || {}), sectPr: followingSectPr },
              sectionMargins: buildSectionMargins(followingSectPr),
            },
            node.marks,
          );
        }
      }
    }

    // Caret at the start of the new section, like Word.
    const caretPos = paragraphPos + tr.doc.nodeAt(paragraphPos).nodeSize + 1;
    tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(caretPos, tr.doc.content.size))));
    tr.setMeta('forceUpdatePagination', true);
    tr.scrollIntoView();
    return true;
  };
