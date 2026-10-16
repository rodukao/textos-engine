/**
 * Line-by-line wrapping around floating objects.
 *
 * Word narrows only the lines that sit beside a float: a paragraph that starts beside an image and
 * continues below it goes back to the full column width once it clears the image. Here the
 * paragraph is laid out in bands: lines are measured at the width available at their height, and
 * when that width changes the rest of the paragraph is remeasured from that line on.
 *
 * @module float-bands
 */
import type { Line, LineSegment, ParagraphBlock, ParagraphMeasure, Run } from '@superdoc/contracts';

/** Horizontal room for a line: width and offset from the column's left edge. */
export type LineBand = { width: number; offsetX: number };

type RunPosition = { run: number; char: number };

const BAND_EPSILON = 0.5;

/** Bands on each side of a float can differ by rounding only. */
export function sameBand(a: LineBand | undefined, b: LineBand | undefined): boolean {
  if (!a || !b) return a === b;
  return Math.abs(a.width - b.width) < BAND_EPSILON && Math.abs(a.offsetX - b.offsetX) < BAND_EPSILON;
}

export function isFullBand(band: LineBand, columnWidth: number): boolean {
  return band.width >= columnWidth - BAND_EPSILON && band.offsetX < BAND_EPSILON;
}

/**
 * Banded layout splits the paragraph between lines. List markers and drop caps belong to the first
 * line only and are measured with it, so those paragraphs keep the single-width layout.
 */
export function canLayoutInBands(block: ParagraphBlock, measure: ParagraphMeasure): boolean {
  const attrs = block.attrs;
  if (measure.marker || attrs?.wordLayout?.marker || attrs?.numberingProperties) return false;
  if (attrs?.dropCap || attrs?.dropCapDescriptor) return false;
  return Array.isArray(block.runs) && block.runs.length > 0;
}

const isTextLikeRun = (run: Run): run is Run & { text: string } =>
  (run.kind === undefined || run.kind === 'text') && typeof (run as { text?: unknown }).text === 'string';

/**
 * The paragraph from `start` on, as a block of its own. Its first line is a continuation line, so
 * the first-line and hanging indents do not apply to it.
 */
function sliceParagraphBlock(block: ParagraphBlock, start: RunPosition): ParagraphBlock | undefined {
  const runs = block.runs.slice(start.run);
  if (runs.length === 0) return undefined;
  if (start.char > 0) {
    const first = runs[0];
    if (!isTextLikeRun(first)) return undefined;
    const pmStart = (first as { pmStart?: number }).pmStart;
    runs[0] = {
      ...first,
      text: first.text.slice(start.char),
      ...(typeof pmStart === 'number' ? { pmStart: pmStart + start.char } : {}),
    } as Run;
  }
  const indent = block.attrs?.indent;
  return {
    ...block,
    runs,
    attrs: { ...block.attrs, ...(indent ? { indent: { ...indent, firstLine: 0, hanging: 0 } } : {}) },
  };
}

/** Maps a position inside a sliced block back to the original block. */
const toOriginal = (base: RunPosition, run: number, char: number): RunPosition => ({
  run: base.run + run,
  char: run === 0 ? base.char + char : char,
});

function remapLine(line: Line, base: RunPosition): Line {
  if (base.run === 0 && base.char === 0) return line;
  const from = toOriginal(base, line.fromRun, line.fromChar);
  const to = toOriginal(base, line.toRun, line.toChar);
  const segments = line.segments?.map((segment): LineSegment => {
    const segFrom = toOriginal(base, segment.runIndex, segment.fromChar);
    return {
      ...segment,
      runIndex: segFrom.run,
      fromChar: segFrom.char,
      toChar: segment.runIndex === 0 ? base.char + segment.toChar : segment.toChar,
    };
  });
  return {
    ...line,
    fromRun: from.run,
    fromChar: from.char,
    toRun: to.run,
    toChar: to.char,
    ...(segments ? { segments } : {}),
  };
}

export type BandLayoutInput = {
  block: ParagraphBlock;
  /** The paragraph measured at the full column width. */
  fullWidthLines: Line[];
  startY: number;
  columnWidth: number;
  /** Width to measure text at for a band (the band minus paragraph indents). */
  measureWidth: (bandWidth: number) => number;
  /** Room for a line at a given height. */
  bandAt: (y: number, lineHeight: number) => LineBand;
  remeasure: (part: ParagraphBlock, width: number) => Line[];
};

/**
 * Lays the paragraph out in bands. Returns undefined when it cannot be split (e.g. a line starts
 * inside a non-text run), so the caller keeps the single-width layout.
 */
export function layoutLinesInBands(input: BandLayoutInput): { lines: Line[]; bands: LineBand[] } | undefined {
  const { block, fullWidthLines, columnWidth, bandAt } = input;
  const lines: Line[] = [];
  const bands: LineBand[] = [];
  let y = input.startY;
  let base: RunPosition = { run: 0, char: 0 };
  let part: ParagraphBlock = block;
  let partLines = fullWidthLines;
  let partBand: LineBand = { width: columnWidth, offsetX: 0 };

  for (let guard = 0; guard < 500; guard += 1) {
    const band = bandAt(y, partLines[0]?.lineHeight ?? 0);
    if (!sameBand(band, partBand)) {
      partLines =
        isFullBand(band, columnWidth) && part === block
          ? fullWidthLines
          : input.remeasure(part, isFullBand(band, columnWidth) ? columnWidth : input.measureWidth(band.width));
      partBand = band;
    }
    if (partLines.length === 0) break;

    let taken = 0;
    while (taken < partLines.length) {
      const line = partLines[taken];
      // The first line of a band is always taken, so the loop advances even if the band changes
      // right at it (its own height may reach into the next band).
      if (taken > 0 && !sameBand(bandAt(y, line.lineHeight || 0), band)) break;
      lines.push(remapLine(line, base));
      bands.push(band);
      y += line.lineHeight || 0;
      taken += 1;
    }
    if (taken >= partLines.length) return { lines, bands };

    const next = partLines[taken];
    base = toOriginal(base, next.fromRun, next.fromChar);
    const nextPart = sliceParagraphBlock(block, base);
    if (!nextPart) return undefined;
    part = nextPart;
    // Force a remeasure of the remaining text for whatever band comes next.
    partBand = { width: -1, offsetX: -1 };
  }
  return undefined;
}
