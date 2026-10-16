/**
 * Word font family fallbacks derived from common DOCX classifications.
 *
 * Maps DOCX font family types (w:family attribute values) to CSS font-family fallback chains.
 * This provides sensible defaults when a specific font is not available.
 *
 * Reference: ISO/IEC 29500-1:2016 (OOXML) Section 17.4.68 - ST_FontFamily
 *
 * @constant {Readonly<Record<string, string>>}
 * @property {string} swiss - Sans-serif fonts (e.g., Arial, Helvetica) → "Arial, sans-serif"
 * @property {string} roman - Serif fonts (e.g., Times New Roman) → "Times New Roman, serif"
 * @property {string} modern - Monospace fonts (e.g., Courier) → "Courier New, monospace"
 * @property {string} script - Cursive/script fonts → "cursive"
 * @property {string} decorative - Decorative/fantasy fonts → "fantasy"
 * @property {string} system - System default fonts → "system-ui"
 * @property {string} auto - Automatic font selection → "sans-serif"
 *
 * @example
 * // Using a DOCX font family classification
 * const fallback = FONT_FAMILY_FALLBACKS['swiss']; // "Arial, sans-serif"
 * const style = `font-family: Helvetica, ${fallback}`;
 */
export const FONT_FAMILY_FALLBACKS = Object.freeze({
  swiss: 'Arial, sans-serif',
  roman: 'Times New Roman, serif',
  modern: 'Courier New, monospace',
  script: 'cursive',
  decorative: 'fantasy',
  system: 'system-ui',
  auto: 'sans-serif',
});

/**
 * Default CSS generic font-family to use when no specific family is known.
 *
 * This is the most universally supported generic font family and provides
 * a safe fallback for all font rendering scenarios.
 *
 * @constant {string}
 *
 * @example
 * // Use as a final fallback
 * const fontFamily = specificFont || DEFAULT_GENERIC_FALLBACK; // "sans-serif"
 */
export const DEFAULT_GENERIC_FALLBACK = 'sans-serif';

/**
 * Known serif-like font families used as a heuristic when OOXML `w:family`
 * is unavailable. This keeps fallbacks closer to Word metrics for fonts like Cambria.
 */
const SERIF_LIKE_FONTS = new Set([
  'cambria',
  'cambria math',
  'times',
  'times new roman',
  'georgia',
  'garamond',
  'palatino',
  'palatino linotype',
  'book antiqua',
  'baskerville',
  'cochin',
  'hoefler text',
  'minion pro',
  'didot',
  'bodoni mt',
  'constantia',
]);

const normalizeFontNameForLookup = (fontName) => {
  if (!fontName || typeof fontName !== 'string') return '';
  return fontName
    .trim()
    .replace(/^["']|["']$/g, '')
    .toLowerCase();
};

const inferGenericFallbackFromFontName = (fontName) =>
  SERIF_LIKE_FONTS.has(normalizeFontNameForLookup(fontName)) ? 'serif' : DEFAULT_GENERIC_FALLBACK;

/**
 * Normalizes a comma-separated font-family string into an array of trimmed, non-empty parts.
 *
 * This internal helper splits a CSS font-family string by commas, trims whitespace
 * from each part, and filters out any empty strings. This is used to clean up
 * fallback chains before processing.
 *
 * @private
 * @param {string | undefined | null} value - The font-family string to normalize
 * @returns {string[]} Array of trimmed, non-empty font family names
 *
 * @example
 * normalizeParts('Arial, sans-serif'); // ['Arial', 'sans-serif']
 * normalizeParts('  Times  ,  serif  '); // ['Times', 'serif']
 * normalizeParts('Arial,,,serif'); // ['Arial', 'serif']
 * normalizeParts(''); // []
 * normalizeParts(null); // []
 */
const normalizeParts = (value) =>
  (value || '')
    .split(',')
    .map((part) => part.trim())
    .filter(Boolean);

/**
 * Splits a string by a delimiter, but only when the delimiter is outside of quotes.
 *
 * @private
 * @param {string} str - The string to split
 * @param {string} delimiter - The delimiter to split on (single character)
 * @returns {string[]} Array of parts, trimmed and filtered for empty strings
 */
const splitOutsideQuotes = (str, delimiter) => {
  // Fast path: no quotes, just split normally
  if (!str.includes('"') && !str.includes("'")) {
    return str
      .split(delimiter)
      .map((p) => p.trim())
      .filter(Boolean);
  }

  const parts = [];
  let current = '';
  let inQuote = false;
  let quoteChar = null;

  for (const char of str) {
    if (!inQuote && (char === '"' || char === "'")) {
      inQuote = true;
      quoteChar = char;
      current += char;
    } else if (inQuote && char === quoteChar) {
      inQuote = false;
      quoteChar = null;
      current += char;
    } else if (!inQuote && char === delimiter) {
      if (current.trim()) parts.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }

  if (current.trim()) parts.push(current.trim());

  return parts;
};

/**
 * Maps a DOCX font family classification to a CSS fallback string.
 *
 * Takes a Word/DOCX font family type (from the w:family attribute) and returns
 * an appropriate CSS font-family fallback chain. The mapping is case-insensitive.
 * If the family is unknown or falsy, returns the default generic fallback.
 *
 * @param {string | undefined | null} wordFamily - The DOCX font family classification
 *   (e.g., 'swiss', 'roman', 'modern', 'script', 'decorative', 'system', 'auto')
 * @returns {string} CSS font-family fallback string. Returns DEFAULT_GENERIC_FALLBACK
 *   ('sans-serif') if wordFamily is null, undefined, empty, or unrecognized.
 *
 * @example
 * // Map known DOCX families
 * mapWordFamilyFallback('swiss'); // "Arial, sans-serif"
 * mapWordFamilyFallback('roman'); // "Times New Roman, serif"
 * mapWordFamilyFallback('modern'); // "Courier New, monospace"
 *
 * @example
 * // Case-insensitive matching
 * mapWordFamilyFallback('SWISS'); // "Arial, sans-serif"
 * mapWordFamilyFallback('RoMaN'); // "Times New Roman, serif"
 *
 * @example
 * // Fallback for unknown/missing values
 * mapWordFamilyFallback('unknown'); // "sans-serif"
 * mapWordFamilyFallback(null); // "sans-serif"
 * mapWordFamilyFallback(''); // "sans-serif"
 */
export function mapWordFamilyFallback(wordFamily) {
  if (!wordFamily) return DEFAULT_GENERIC_FALLBACK;
  const mapped = FONT_FAMILY_FALLBACKS[wordFamily.toLowerCase()];
  return mapped || DEFAULT_GENERIC_FALLBACK;
}

/**
 * Composes a CSS font-family string with an appropriate fallback chain.
 *
 * This function takes a primary font name and intelligently appends a fallback chain,
 * ensuring proper CSS syntax and avoiding duplication. It implements the following algorithm:
 *
 * 1. **Passthrough for non-strings**: Returns the input unchanged if it's not a string
 *    (null, undefined, number, object, etc.)
 * 2. **Trimming**: Removes leading/trailing whitespace from the font name
 * 3. **Early return**: If the trimmed string is empty or already contains commas,
 *    returns it as-is (assumes it's already a complete font-family declaration)
 * 4. **Fallback resolution**: Determines the fallback chain using this precedence:
 *    - Explicit `options.fallback` (highest priority)
 *    - `options.wordFamily` mapped via mapWordFamilyFallback()
 *    - DEFAULT_GENERIC_FALLBACK ('sans-serif')
 * 5. **Duplicate detection**: If the font name already appears in the fallback chain
 *    (case-insensitive), returns just the fallback chain
 * 6. **Composition**: Prepends the font name to the fallback chain
 *
 * **Edge cases handled:**
 * - Preserves quotes in font names (e.g., "Times New Roman")
 * - Case-insensitive duplicate detection prevents "Arial, Arial, sans-serif"
 * - Normalizes whitespace in fallback chains (trims each part, filters empties)
 * - Empty fallback results in just the font name (no trailing comma)
 *
 * @param {string | undefined | null} fontName - The primary font name to use.
 *   Can be null/undefined (returned as-is), or a string font name.
 * @param {{ fallback?: string; wordFamily?: string | null }} [options={}] - Configuration options
 * @param {string} [options.fallback] - Explicit CSS fallback string (e.g., 'Arial, sans-serif').
 *   Takes precedence over wordFamily. Can be a comma-separated list.
 * @param {string | null} [options.wordFamily] - DOCX font family classification
 *   (e.g., 'swiss', 'roman'). Mapped to CSS fallback via mapWordFamilyFallback().
 *   Ignored if options.fallback is provided.
 *
 * @returns {string | undefined | null} A complete CSS font-family string, or the original
 *   value if it was not a string. Never throws errors; returns passthrough values safely.
 *
 * @throws {never} This function does not throw. Invalid inputs are returned as-is.
 *
 * @example
 * // Basic usage with default fallback
 * toCssFontFamily('Arial'); // "Arial, sans-serif"
 * toCssFontFamily('Times New Roman'); // "Times New Roman, serif"
 *
 * @example
 * // Custom explicit fallback
 * toCssFontFamily('Georgia', { fallback: 'Times, serif' });
 * // "Georgia, Times, serif"
 *
 * @example
 * // DOCX wordFamily option
 * toCssFontFamily('Helvetica', { wordFamily: 'swiss' });
 * // "Helvetica, Arial, sans-serif"
 *
 * toCssFontFamily('Courier', { wordFamily: 'modern' });
 * // "Courier, Courier New, monospace"
 *
 * @example
 * // Fallback takes precedence over wordFamily
 * toCssFontFamily('MyFont', { fallback: 'serif', wordFamily: 'swiss' });
 * // "MyFont, serif" (wordFamily is ignored)
 *
 * @example
 * // Duplicate detection (case-insensitive)
 * toCssFontFamily('Arial', { fallback: 'Arial, sans-serif' });
 * // "Arial, sans-serif" (not "Arial, Arial, sans-serif")
 *
 * toCssFontFamily('arial', { fallback: 'Arial, sans-serif' });
 * // "Arial, sans-serif" (case-insensitive match)
 *
 * @example
 * // Passthrough for non-strings and special cases
 * toCssFontFamily(null); // null
 * toCssFontFamily(undefined); // undefined
 * toCssFontFamily(''); // '' (empty string)
 * toCssFontFamily('  '); // '' (whitespace-only)
 * toCssFontFamily('Arial, sans-serif'); // "Arial, sans-serif" (already has comma)
 * toCssFontFamily(123); // 123 (non-string passthrough)
 *
 * @example
 * // Whitespace handling
 * toCssFontFamily('  Arial  '); // "Arial, sans-serif"
 * toCssFontFamily('Arial', { fallback: '  serif  ,  monospace  ' });
 * // "Arial, serif, monospace"
 */
export function toCssFontFamily(fontName, options = {}) {
  if (!fontName || typeof fontName !== 'string') return fontName;
  let trimmed = fontName.trim();
  if (!trimmed || trimmed.includes(',')) return trimmed;
  // Replace semicolon font fallback separators (e.g., "Liberation Sans;Arial" from LibreOffice).
  // Only split on semicolons outside of quotes to preserve font names like "Foo;Bar".
  if (trimmed.includes(';')) {
    trimmed = splitOutsideQuotes(trimmed, ';').join(', ');
  }

  const { fallback, wordFamily } = options;
  const fallbackValue =
    fallback ??
    (wordFamily ? mapWordFamilyFallback(wordFamily) : undefined) ??
    inferGenericFallbackFromFontName(trimmed);

  const fallbackParts = normalizeParts(fallbackValue);
  if (fallbackParts.length === 0) {
    return trimmed;
  }

  const normalizedName = trimmed.toLowerCase();
  const includesName = fallbackParts.some((part) => part.toLowerCase() === normalizedName);
  if (includesName) {
    return fallbackParts.join(', ');
  }

  return [trimmed, ...fallbackParts].join(', ');
}

/** Size characters are measured at to get the font's design advances, free of small-size hinting. */
const UNHINTED_REFERENCE_SIZE_PX = 1000;
const FONT_SIZE_PATTERN = /(\d*\.?\d+)px/;

/**
 * Width of text from the font's design advances, which is what Word lays out with.
 *
 * At small sizes browsers may use hinted glyph advances: Chromium on Windows at 100% scale draws
 * Cambria 12pt about 1.8% wider than its design widths, so lines would break earlier than in Word.
 * Hinting rounds each glyph's advance on its own, so the text is measured at its size as usual and
 * each character's difference (design minus hinted advance, measured once per font) is added; the
 * result matches measuring at a large size and scaling down, kerning included, at a fraction of
 * the cost.
 *
 * @param {CanvasRenderingContext2D} ctx - canvas context; its font is left set to `font`
 * @param {string} text
 * @param {string} font - CSS font shorthand, e.g. "italic bold 16px Cambria"
 * @returns {number} width in pixels at the font's size
 */
export function measureUnhintedTextWidth(ctx, text, font) {
  const correction = unhintedTextCorrection(ctx, text, font);
  ctx.font = font;
  return ctx.measureText(text).width + correction;
}

/**
 * Design width minus the browser's (hinted) width of `text`, in pixels: 0 where the browser
 * already uses design advances. Painters add it as letter spacing so drawn text matches the layout.
 *
 * @param {CanvasRenderingContext2D} ctx
 * @param {string} text
 * @param {string} font - CSS font shorthand
 * @returns {number}
 */
export function unhintedTextCorrection(ctx, text, font) {
  const deltas = characterDeltasFor(ctx, font);
  if (!deltas) return 0;
  let correction = 0;
  let missing = null;
  for (let i = 0; i < text.length; i += 1) {
    const code = text.codePointAt(i);
    if (code > 0xffff) i += 1;
    const delta = deltas.get(code);
    if (delta === undefined) (missing ??= new Set()).add(code);
    else correction += delta;
  }
  if (missing) {
    // Changing the canvas font is the expensive part: measure all new characters in one go.
    const codes = [...missing];
    for (const [index, delta] of measureCharacterDeltas(ctx, codes, font).deltas.entries()) {
      deltas.set(codes[index], delta);
    }
    for (let i = 0; i < text.length; i += 1) {
      const code = text.codePointAt(i);
      if (code > 0xffff) i += 1;
      if (missing.has(code)) correction += deltas.get(code);
    }
  }
  return correction;
}

/**
 * Per font: design minus hinted advance of each character (by code point) seen so far, or null for
 * fonts the browser draws at their design advances anyway.
 */
const deltasByFont = new Map();
const MAX_FONTS = 200;
/** Fonts whose hinted widths stay within this fraction of the design widths are left alone. */
const NEGLIGIBLE_HINTING = 0.001;
/** Hinting moves widths by a few percent at most; beyond this the context ignores the size (a stub). */
const IMPLAUSIBLE_HINTING = 0.25;
/**
 * Frequent letters: enough to tell whether a font is hinted at a size. Measuring a glyph for the
 * first time in a face/size is the costly part, so the probe stays short.
 */
const PROBE_CODES = Array.from('aeimnorst', (character) => character.codePointAt(0));

function characterDeltasFor(ctx, font) {
  const known = deltasByFont.get(font);
  if (known !== undefined) return known;
  const match = FONT_SIZE_PATTERN.exec(font);
  const size = match ? Number.parseFloat(match[1]) : 0;
  let deltas = null;
  if (size > 0 && size < UNHINTED_REFERENCE_SIZE_PX && typeof ctx === 'object' && ctx !== null) {
    const { deltas: probeDeltas, hinted } = measureCharacterDeltas(ctx, PROBE_CODES, font);
    const drift = hinted > 0 ? Math.abs(probeDeltas.reduce((sum, delta) => sum + delta, 0)) / hinted : Infinity;
    if (drift > NEGLIGIBLE_HINTING && drift < IMPLAUSIBLE_HINTING) {
      deltas = new Map(PROBE_CODES.map((code, index) => [code, probeDeltas[index]]));
    }
  }
  if (deltasByFont.size >= MAX_FONTS) deltasByFont.clear();
  deltasByFont.set(font, deltas);
  return deltas;
}

/**
 * Per face (the font at the reference size, e.g. "bold 1000px Cambria"): design advance of each
 * character at the reference size. Shared by every size of the face, so the reference size is
 * measured once per face rather than once per size.
 */
const designAdvancesByFace = new Map();

function referenceFont(font) {
  return font.replace(FONT_SIZE_PATTERN, `${UNHINTED_REFERENCE_SIZE_PX}px`);
}

/** Design minus hinted advance of each character at `font`, and their total hinted width. */
function measureCharacterDeltas(ctx, codes, font) {
  const size = Number.parseFloat(FONT_SIZE_PATTERN.exec(font)[1]);
  const face = referenceFont(font);
  let advances = designAdvancesByFace.get(face);
  if (!advances) {
    if (designAdvancesByFace.size >= MAX_FONTS) designAdvancesByFace.clear();
    advances = new Map();
    designAdvancesByFace.set(face, advances);
  }
  const unknown = codes.filter((code) => !advances.has(code));
  if (unknown.length > 0) {
    ctx.font = face;
    for (const code of unknown) advances.set(code, ctx.measureText(String.fromCodePoint(code)).width);
  }
  ctx.font = font;
  let hinted = 0;
  const deltas = codes.map((code) => {
    const drawn = ctx.measureText(String.fromCodePoint(code)).width;
    hinted += drawn;
    const delta = (advances.get(code) * size) / UNHINTED_REFERENCE_SIZE_PX - drawn;
    return Number.isFinite(delta) ? delta : 0;
  });
  return { deltas, hinted };
}
