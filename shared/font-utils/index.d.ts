export type WordFontFamily = 'swiss' | 'roman' | 'modern' | 'script' | 'decorative' | 'system' | 'auto' | (string & {});

export const FONT_FAMILY_FALLBACKS: Readonly<Record<string, string>>;
export const DEFAULT_GENERIC_FALLBACK: string;

export interface ToCssFontFamilyOptions {
  fallback?: string;
  wordFamily?: WordFontFamily | null;
}

export function mapWordFamilyFallback(wordFamily?: WordFontFamily | null): string;

export function toCssFontFamily(fontName?: string | null, options?: ToCssFontFamilyOptions): string | undefined | null;

/** Width of text from the font's design advances (no small-size hinting), as Word lays it out. */
export function measureUnhintedTextWidth(
  ctx: Pick<CanvasRenderingContext2D, 'font' | 'measureText'>,
  text: string,
  font: string,
): number;

/** Design width minus the browser's hinted width of text, in pixels (0 without hinting). */
export function unhintedTextCorrection(
  ctx: Pick<CanvasRenderingContext2D, 'font' | 'measureText'>,
  text: string,
  font: string,
): number;
