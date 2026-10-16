/**
 * Draw text at the font's design advances.
 *
 * The layout measures text from design advances, as Word does (see measureUnhintedTextWidth in
 * @superdoc/font-utils), but at small sizes the browser may draw it with hinted advances
 * (Chromium on Windows at 100% scale draws Cambria ~1.8% wider). A per-run letter-spacing
 * correction makes the drawn width equal the measured one, so lines neither overflow the margin
 * nor end short.
 *
 * @module unhinted-advance
 */
import { unhintedTextCorrection } from '@superdoc/font-utils';

let context: CanvasRenderingContext2D | null | undefined;

const getContext = (doc: Document): CanvasRenderingContext2D | null => {
  if (context === undefined) {
    try {
      context = doc.createElement('canvas').getContext('2d');
    } catch {
      context = null;
    }
  }
  return context ?? null;
};

/**
 * Extra letter spacing (px per character) that brings the browser's drawn width of `text` to its
 * design-advance width. 0 when no canvas is available (e.g. tests without one).
 */
export function unhintedLetterSpacing(doc: Document, text: string, font: string): number {
  if (!text) return 0;
  const ctx = getContext(doc);
  if (!ctx) return 0;
  const correction = unhintedTextCorrection(ctx, text, font);
  if (correction === 0) return 0;
  // CSS letter-spacing is added after every character, including the last.
  let characters = 0;
  for (const _ of text) characters += 1;
  return correction / characters;
}
