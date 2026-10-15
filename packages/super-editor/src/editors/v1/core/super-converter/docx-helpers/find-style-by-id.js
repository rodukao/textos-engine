// @ts-check
import { DEFAULT_LINKED_STYLES } from '../exporter-docx-defs.js';

/**
 * Built-in Word style name for one of SuperDoc's default English style ids
 * (e.g. 'Heading1' → 'heading 1'), lower-cased. Undefined for other ids.
 * @param {string} styleId
 * @returns {string | undefined}
 */
const builtInStyleName = (styleId) =>
  DEFAULT_LINKED_STYLES[styleId]?.elements
    ?.find((/** @type {any} */ el) => el.name === 'w:name')
    ?.attributes?.['w:val']?.toLowerCase();

/**
 * Find a linked style by id. Localized Word builds keep the built-in style name
 * but translate the id (pt-BR: 'Ttulo1' for 'heading 1'), so a default English
 * id such as 'Heading1' falls back to the document's style with the same name.
 * Word compares style names case-insensitively.
 *
 * @param {Array<any> | null | undefined} styles Linked styles (`{ id, definition: { attrs: { name } } }`)
 * @param {string | null | undefined} styleId
 * @returns {any | undefined}
 */
export const findStyleById = (styles, styleId) => {
  if (!styleId || !Array.isArray(styles)) return undefined;
  const exact = styles.find((style) => style?.id === styleId);
  if (exact) return exact;
  const name = builtInStyleName(styleId);
  if (!name) return undefined;
  return styles.find((style) => style?.definition?.attrs?.name?.toLowerCase() === name);
};
