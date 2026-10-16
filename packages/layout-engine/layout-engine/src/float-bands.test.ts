import { describe, expect, it } from 'bun:test';
import type { Line, ParagraphBlock } from '@superdoc/contracts';
import { layoutLinesInBands, type LineBand } from './float-bands.js';

/** Fake measurer: every `charsPerLine(width)` characters make a line of height 20. */
const fakeLines = (block: ParagraphBlock, width: number): Line[] => {
  const text = block.runs.map((r) => (r as { text: string }).text).join('');
  const perLine = Math.max(1, Math.floor(width / 10));
  const lines: Line[] = [];
  for (let from = 0; from < text.length; from += perLine) {
    const to = Math.min(text.length, from + perLine);
    lines.push({
      fromRun: 0,
      fromChar: from,
      toRun: 0,
      toChar: to,
      width: (to - from) * 10,
      ascent: 15,
      descent: 5,
      lineHeight: 20,
    });
  }
  return lines;
};

describe('layoutLinesInBands', () => {
  const block: ParagraphBlock = {
    kind: 'paragraph',
    id: 'p',
    runs: [{ text: 'x'.repeat(100), fontFamily: 'Arial', fontSize: 16 }],
  };
  // A float on the left covering y 0..60: three 20px lines are narrowed to 200px, the rest is full width.
  const bandAt = (y: number, h: number): LineBand =>
    y < 60 && y + h > 0 ? { width: 200, offsetX: 300 } : { width: 500, offsetX: 0 };

  it('narrows only the lines beside the float', () => {
    const result = layoutLinesInBands({
      block,
      fullWidthLines: fakeLines(block, 500),
      startY: 0,
      columnWidth: 500,
      measureWidth: (w) => w,
      bandAt,
      remeasure: fakeLines,
    });
    expect(result).toBeDefined();
    const { lines, bands } = result!;
    // 3 lines of 20 chars beside the float, then 40 chars at 50 per line.
    expect(lines.map((l) => [l.fromChar, l.toChar])).toEqual([
      [0, 20],
      [20, 40],
      [40, 60],
      [60, 100],
    ]);
    expect(bands.map((b) => b.width)).toEqual([200, 200, 200, 500]);
  });

  it('keeps the full-width measure when no line is beside a float', () => {
    const full = fakeLines(block, 500);
    const result = layoutLinesInBands({
      block,
      fullWidthLines: full,
      startY: 100,
      columnWidth: 500,
      measureWidth: (w) => w,
      bandAt,
      remeasure: () => {
        throw new Error('should not remeasure');
      },
    });
    expect(result?.lines).toEqual(full);
  });
});
