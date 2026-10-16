/**
 * Word "single" line height
 *
 * Word does not use a fixed 1.15 × font size for single line spacing. On Windows it takes the
 * font's GDI text metrics: tmHeight (usWinAscent + usWinDescent) plus tmExternalLeading, where
 * the external leading is max(0, hhea.lineGap - ((winAscent + winDescent) - (hhea.ascender -
 * hhea.descender))). The result depends on the font: Calibri is 1.2207 × size, Cambria 1.1724,
 * Times New Roman and Arial 1.1499, Segoe UI 1.3301.
 *
 * "Multiple" spacing (lineRule="auto") multiplies that single height, and Word rounds every line
 * height up to 1/960 inch (0.075 pt = 0.1 px). Both rules were measured against Word 16 for 7 fonts,
 * 4 sizes and 4 spacings (112 cases, all exact), plus empty paragraphs, multiples below 1 and
 * "at least" spacing.
 *
 * The ratios below are keyed by the font named in the document (not the font the browser ends up
 * rendering), so the layout follows Word even when a metric-compatible substitute is used.
 *
 * @module wordLineMetrics
 */

/**
 * Single line height / font size, from each font's OS/2 and hhea tables (regular face; the bold
 * and italic faces of these families have the same vertical metrics).
 * Keys are lowercase family names.
 */
const WORD_SINGLE_LINE_RATIOS: Record<string, number> = {
  arial: 1.1499,
  'arial black': 1.41016,
  'arial narrow': 1.14746,
  bahnschrift: 1.2002,
  'bell mt': 1.11182,
  'book antiqua': 1.24268,
  'bookman old style': 1.17383,
  calibri: 1.2207,
  'calibri light': 1.2207,
  'calisto mt': 1.1543,
  cambria: 1.17236,
  candara: 1.2207,
  'candara light': 1.2207,
  centaur: 1.14014,
  century: 1.20215,
  'century gothic': 1.22607,
  'comic sans ms': 1.39355,
  consolas: 1.1709,
  constantia: 1.2207,
  corbel: 1.2207,
  'corbel light': 1.2207,
  'courier new': 1.13281,
  'footlight mt light': 1.05566,
  'franklin gothic medium': 1.13379,
  'franklin gothic medium cond': 1.13379,
  garamond: 1.125,
  georgia: 1.13623,
  'gill sans mt': 1.15967,
  'gill sans mt condensed': 1.2041,
  'gill sans ultra bold': 1.24609,
  'gill sans ultra bold condensed': 1.24512,
  impact: 1.21973,
  'lucida console': 1,
  'lucida sans': 1.17773,
  'microsoft sans serif': 1.13184,
  'palatino linotype': 1.34912,
  rockwell: 1.17432,
  'rockwell extra bold': 1.17432,
  'segoe ui': 1.33008,
  'segoe ui light': 1.33008,
  'segoe ui semibold': 1.33008,
  'segoe ui semilight': 1.33008,
  tahoma: 1.20703,
  'times new roman': 1.1499,
  'trebuchet ms': 1.16113,
  'tw cen mt': 1.08887,
  verdana: 1.21533,
};

/** Metric-compatible substitutes share the vertical metrics of the font they replace. */
const METRIC_COMPATIBLE_ALIASES: Record<string, string> = {
  carlito: 'calibri',
  caladea: 'cambria',
  'liberation serif': 'times new roman',
  tinos: 'times new roman',
  'liberation sans': 'arial',
  arimo: 'arial',
  'liberation mono': 'courier new',
  cousine: 'courier new',
};

/** Fallback when the font is unknown and the browser cannot report font metrics. */
export const DEFAULT_SINGLE_LINE_RATIO = 1.15;

/** Word rounds line heights up to 1/960 inch: 0.1 px at 96 px per inch. */
const WORD_LINE_HEIGHT_QUANTUM_PX = 0.1;

/** Values this close above a quantum step are float noise in Word's own arithmetic, not a new step. */
const QUANTUM_TOLERANCE = 0.025;

/** First family of a CSS font-family list, unquoted and lowercased. */
const primaryFamily = (fontFamily: string): string => {
  const first = fontFamily.split(',')[0] ?? '';
  return first
    .trim()
    .replace(/^["']|["']$/g, '')
    .trim()
    .toLowerCase();
};

/** Lookups by the exact font-family string, since measuring asks once per line. */
const ratioByFontFamily = new Map<string, number | undefined>();

/**
 * Returns Word's single line height ratio for a font family, or undefined when the family is not
 * in the table.
 */
export function getWordSingleLineRatio(fontFamily: string | undefined): number | undefined {
  if (!fontFamily) return undefined;
  if (ratioByFontFamily.has(fontFamily)) return ratioByFontFamily.get(fontFamily);
  const family = primaryFamily(fontFamily);
  const ratio = WORD_SINGLE_LINE_RATIOS[METRIC_COMPATIBLE_ALIASES[family] ?? family];
  if (ratioByFontFamily.size < 500) ratioByFontFamily.set(fontFamily, ratio);
  return ratio;
}

/** Rounds a line height up to Word's 1/960 inch grid. */
export function quantizeWordLineHeight(px: number): number {
  if (!Number.isFinite(px) || px <= 0) return px;
  const steps = Math.ceil(px / WORD_LINE_HEIGHT_QUANTUM_PX - QUANTUM_TOLERANCE);
  return Math.round(steps * WORD_LINE_HEIGHT_QUANTUM_PX * 1000) / 1000;
}
