/* @vitest-environment jsdom */

import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import { initTestEditor, loadTestDataForEditorTests } from '@tests/helpers/helpers.js';
import { toFlowBlocks } from '@core/layout-adapter';
import type { Editor } from '../core/Editor.js';

type LoadedDocData = Awaited<ReturnType<typeof loadTestDataForEditorTests>>;

type AnyEditor = Editor & {
  doc: any;
  commands: any;
  converter: any;
};

function sectionsOf(editor: AnyEditor) {
  return editor.doc.sections.list();
}

function bodyTarget(editor: AnyEditor) {
  return sectionsOf(editor).items[0].address;
}

function paragraphTexts(editor: AnyEditor): string[] {
  const texts: string[] = [];
  editor.state.doc.forEach((node) => texts.push(node.textContent));
  return texts;
}

function setParagraphs(editor: AnyEditor, texts: string[]): void {
  const { schema } = editor.state;
  const paragraphs = texts.map((text) =>
    schema.nodes.paragraph.create(null, text ? schema.nodes.run.create(null, schema.text(text)) : null),
  );
  const tr = editor.state.tr.replaceWith(0, editor.state.doc.content.size, paragraphs);
  tr.setMeta('addToHistory', false);
  editor.dispatch(tr);
}

/** Put the caret after `offset` characters of the paragraph at `index`. */
function placeCaret(editor: AnyEditor, index: number, offset: number): void {
  let target = -1;
  let seen = 0;
  editor.state.doc.forEach((node, pos, i) => {
    if (i !== index) return;
    node.descendants((child, childPos) => {
      if (target >= 0 || !child.isText) return;
      const len = child.text?.length ?? 0;
      if (seen + len >= offset) target = pos + 1 + childPos + (offset - seen);
      seen += len;
    });
    if (target < 0) target = pos + 2;
  });
  editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, target)));
}

function layoutSectionBreaks(editor: AnyEditor) {
  const { blocks } = toFlowBlocks(editor.getJSON(), { emitSectionBreaks: true });
  return blocks
    .map((block: any, index: number) => ({ block, index }))
    .filter(({ block }) => block.kind === 'sectionBreak')
    .map(({ block, index }) => ({
      type: block.type,
      isFirst: block.attrs?.isFirstSection === true,
      nextText: (blocks[index + 1] as any)?.runs?.map((run: any) => run.text ?? '').join(''),
    }));
}

describe('section mutations and history', () => {
  let docData: LoadedDocData;
  let editor: AnyEditor | undefined;

  beforeAll(async () => {
    docData = await loadTestDataForEditorTests('blank-doc.docx');
  });

  afterEach(() => {
    editor?.destroy();
    editor = undefined;
  });

  function openEditor(): AnyEditor {
    ({ editor } = initTestEditor({
      content: docData.docx,
      media: docData.media,
      mediaFiles: docData.mediaFiles,
      fonts: docData.fonts,
      useImmediateSetTimeout: false,
    }) as { editor: AnyEditor });
    return editor!;
  }

  describe('undo/redo of body section changes', () => {
    it('restores converter page margins on undo and reapplies them on redo', () => {
      const ed = openEditor();
      expect(ed.converter.pageStyles.pageMargins.top).toBe(1);

      expect(ed.doc.sections.setPageMargins({ target: bodyTarget(ed), top: 2 }).success).toBe(true);
      expect(ed.converter.pageStyles.pageMargins.top).toBe(2);

      ed.commands.undo();
      expect(ed.converter.pageStyles.pageMargins.top).toBe(1);
      expect(ed.converter.bodySectPr).toEqual(ed.state.doc.attrs.bodySectPr);
      expect(sectionsOf(ed).items[0].margins.top).toBe(1);
      expect(editorJsonTopMargin(ed)).toBe('1440');

      ed.commands.redo();
      expect(ed.converter.pageStyles.pageMargins.top).toBe(2);
      expect(sectionsOf(ed).items[0].margins.top).toBe(2);
      expect(editorJsonTopMargin(ed)).toBe('2880');
    });

    it('restores page size on undo of setPageSetup', () => {
      const ed = openEditor();
      expect(ed.doc.sections.setPageSetup({ target: bodyTarget(ed), orientation: 'landscape' }).success).toBe(true);
      expect(ed.converter.pageStyles.pageSize.width).toBe(11);

      ed.commands.undo();
      expect(ed.converter.pageStyles.pageSize.width).toBe(8.5);
      expect(ed.converter.pageStyles.pageSize.height).toBe(11);
      expect(sectionsOf(ed).items[0].pageSetup.orientation).toBe('portrait');
    });

    it('restores columns seen by layout on undo of setColumns', () => {
      const ed = openEditor();
      expect(ed.doc.sections.setColumns({ target: bodyTarget(ed), count: 2 }).success).toBe(true);
      expect(sectionsOf(ed).items[0].columns.count).toBe(2);

      ed.commands.undo();
      expect(sectionsOf(ed).items[0].columns.count).toBe(1);
      const cols = (editor!.getJSON().attrs as any).bodySectPr.elements.find((el: any) => el.name === 'w:cols');
      expect(cols?.attributes?.['w:num'] ?? '1').toBe('1');
    });
  });

  describe('insertSectionBreakAtSelection', () => {
    it('splits the paragraph and starts a new section on the next page', () => {
      const ed = openEditor();
      setParagraphs(ed, ['HelloWorld', 'Tail']);
      placeCaret(ed, 0, 5);

      expect(ed.commands.insertSectionBreakAtSelection()).toBe(true);

      expect(paragraphTexts(ed)).toEqual(['Hello', 'World', 'Tail']);
      const list = sectionsOf(ed);
      expect(list.total).toBe(2);
      expect(list.items[0].range).toEqual({ startParagraphIndex: 0, endParagraphIndex: 0 });
      expect(list.items[1].range).toEqual({ startParagraphIndex: 1, endParagraphIndex: 2 });
      expect(list.items[1].breakType).toBe('nextPage');
      // The ended section keeps the current page setup.
      expect(list.items[0].margins).toEqual(list.items[1].margins);
      expect(list.items[0].pageSetup.width).toBe(list.items[1].pageSetup.width);

      const breaks = layoutSectionBreaks(ed);
      expect(breaks).toHaveLength(2);
      expect(breaks[1]).toMatchObject({ type: 'nextPage', isFirst: false, nextText: 'World' });

      // Caret lands at the start of the new section.
      expect(ed.state.selection.$from.index(0)).toBe(1);
    });

    it('works in the first paragraph and at its start', () => {
      const ed = openEditor();
      setParagraphs(ed, ['Only']);
      placeCaret(ed, 0, 0);

      expect(ed.commands.insertSectionBreakAtSelection()).toBe(true);
      expect(paragraphTexts(ed)).toEqual(['', 'Only']);
      expect(sectionsOf(ed).total).toBe(2);
    });

    it('writes a continuous type on the following section', () => {
      const ed = openEditor();
      setParagraphs(ed, ['One', 'Two']);
      placeCaret(ed, 0, 3);

      expect(ed.commands.insertSectionBreakAtSelection({ breakType: 'continuous' })).toBe(true);
      const list = sectionsOf(ed);
      expect(list.total).toBe(2);
      expect(list.items[1].breakType).toBe('continuous');
      expect(layoutSectionBreaks(ed)[1]).toMatchObject({ type: 'continuous' });
      // Converter follows the body sectPr change made by the command.
      expect(ed.converter.bodySectPr).toEqual(ed.state.doc.attrs.bodySectPr);
    });

    it('supports several breaks and keeps the later section boundaries', () => {
      const ed = openEditor();
      setParagraphs(ed, ['A', 'B', 'C']);
      placeCaret(ed, 0, 1);
      ed.commands.insertSectionBreakAtSelection();
      // Paragraphs now: A | (empty) B C — break inside "B"'s section again.
      placeCaret(ed, 2, 1);
      ed.commands.insertSectionBreakAtSelection();

      const list = sectionsOf(ed);
      expect(list.total).toBe(3);
      expect(list.items.map((item: any) => item.range.endParagraphIndex)).toEqual([0, 2, 4]);
    });

    it('is undone in one step', () => {
      const ed = openEditor();
      setParagraphs(ed, ['HelloWorld']);
      placeCaret(ed, 0, 5);
      ed.commands.insertSectionBreakAtSelection();
      expect(sectionsOf(ed).total).toBe(2);

      ed.commands.undo();
      expect(sectionsOf(ed).total).toBe(1);
      expect(paragraphTexts(ed)).toEqual(['HelloWorld']);
    });
  });
});

function editorJsonTopMargin(editor: AnyEditor): string | undefined {
  const sectPr = (editor.getJSON().attrs as any)?.bodySectPr;
  const pgMar = sectPr?.elements?.find((el: any) => el.name === 'w:pgMar');
  return pgMar?.attributes?.['w:top'];
}
