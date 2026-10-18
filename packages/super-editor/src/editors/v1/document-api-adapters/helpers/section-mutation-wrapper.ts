import type { SectionAddress, SectionMutationResult, MutationOptions } from '@superdoc/document-api';
import type { Editor } from '../../core/Editor.js';
import { DocumentApiAdapterError } from '../errors.js';
import { applyDirectMutationMeta } from './transaction-meta.js';
import { checkRevision } from '../plan-engine/revision-tracker.js';
import { clearIndexCache } from './index-cache.js';
import { rejectTrackedMode } from './mutation-helpers.js';
import { resolveSectionProjections, getBodySectPrFromEditor, type SectionProjection } from './sections-resolver.js';
import { ensureSectPrElement, readSectPrMargins, type XmlElement } from './sections-xml.js';
import { syncConverterBodySectPr, type ConverterWithBodySection } from './body-section-sync.js';
import { readTargetSectPr } from './section-projection-access.js';

function toSectionFailure(
  code: 'NO_OP' | 'INVALID_TARGET' | 'CAPABILITY_UNAVAILABLE',
  message: string,
): SectionMutationResult {
  return { success: false, failure: { code, message } };
}

function toSectionSuccess(section: SectionAddress): SectionMutationResult {
  return { success: true, section };
}

function buildSectionMarginsForAttrs(sectPr: XmlElement): Record<string, number | null> {
  const margins = readSectPrMargins(sectPr);
  return {
    top: margins.top ?? null,
    right: margins.right ?? null,
    bottom: margins.bottom ?? null,
    left: margins.left ?? null,
    header: margins.header ?? null,
    footer: margins.footer ?? null,
  };
}

function syncConverterBodySection(editor: Editor, sectPr: XmlElement): void {
  syncConverterBodySectPr(getConverter(editor) as ConverterWithBodySection | undefined, sectPr);
}

interface ConverterLike {
  bodySectPr?: unknown;
  savedTagsToRestore?: Array<Record<string, unknown>>;
  pageStyles?: {
    pageSize?: { width?: number; height?: number };
    pageMargins?: Record<string, number | undefined>;
  };
}

function getConverter(editor: Editor): ConverterLike | undefined {
  return (editor as unknown as { converter?: ConverterLike }).converter;
}

export function applySectPrToProjection(
  editor: Editor,
  projection: SectionProjection,
  sectPr: XmlElement,
  options?: { addToHistory?: boolean },
): void {
  const addToHistory = options?.addToHistory ?? true;

  if (projection.target.kind === 'paragraph') {
    const paragraph = projection.target.node;
    const attrs = (paragraph.attrs ?? {}) as Record<string, unknown>;
    const paragraphProperties = {
      ...((attrs.paragraphProperties ?? {}) as Record<string, unknown>),
      sectPr,
    };
    const nextAttrs: Record<string, unknown> = {
      ...attrs,
      paragraphProperties,
      pageBreakSource: 'sectPr',
      sectionMargins: buildSectionMarginsForAttrs(sectPr),
    };

    const tr = applyDirectMutationMeta(editor.state.tr);
    if (!addToHistory) {
      tr.setMeta('addToHistory', false);
    }
    tr.setNodeMarkup(projection.target.pos, undefined, nextAttrs, paragraph.marks);
    tr.setMeta('forceUpdatePagination', true);
    editor.dispatch(tr);
    return;
  }

  const tr = applyDirectMutationMeta(editor.state.tr);
  if (!addToHistory) {
    tr.setMeta('addToHistory', false);
  }
  tr.setDocAttribute('bodySectPr', sectPr);
  tr.setMeta('forceUpdatePagination', true);
  editor.dispatch(tr);
  syncConverterBodySection(editor, sectPr);
}

/**
 * Generic sectPr mutation wrapper.
 * Handles: tracked mode rejection, revision check, dry-run, NO_OP detection,
 * applying changes to editor state, clearing index cache.
 */
export function sectionMutationBySectPr<TInput extends { target: SectionAddress }>(
  editor: Editor,
  input: TInput,
  options: MutationOptions | undefined,
  operationName: string,
  mutate: (
    sectPr: XmlElement,
    projection: SectionProjection,
    sections: SectionProjection[],
    dryRun: boolean,
  ) => SectionMutationResult | void,
): SectionMutationResult {
  rejectTrackedMode(operationName, options);
  checkRevision(editor, options?.expectedRevision);

  const sections = resolveSectionProjections(editor);
  const projection = sections.find((entry) => entry.sectionId === input.target.sectionId);
  if (!projection) {
    throw new DocumentApiAdapterError('TARGET_NOT_FOUND', 'Section target was not found.', { target: input.target });
  }

  const dryRun = options?.dryRun === true;

  const currentSectPr = readTargetSectPr(editor, projection);
  const nextSectPr = ensureSectPrElement(currentSectPr);
  const before = JSON.stringify(nextSectPr);
  const earlyResult = mutate(nextSectPr, projection, sections, dryRun);
  if (earlyResult) return earlyResult;

  const changed = before !== JSON.stringify(nextSectPr);
  if (!changed) {
    return toSectionFailure('NO_OP', `${operationName} did not produce a section change.`);
  }

  if (options?.dryRun) {
    return toSectionSuccess(projection.address);
  }

  applySectPrToProjection(editor, projection, nextSectPr);
  clearIndexCache(editor);
  return toSectionSuccess(projection.address);
}

export { toSectionFailure, toSectionSuccess };
