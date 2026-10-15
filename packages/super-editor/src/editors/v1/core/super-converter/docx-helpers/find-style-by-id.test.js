import { describe, it, expect } from 'vitest';
import { findStyleById } from './find-style-by-id.js';

const style = (id, name) => ({ id, type: 'paragraph', definition: { attrs: { name } } });

describe('findStyleById', () => {
  // pt-BR Word: built-in names, translated ids.
  const localized = [
    style('Normal', 'Normal'),
    style('Ttulo', 'Title'),
    style('Ttulo1', 'heading 1'),
    style('Custom', 'My style'),
  ];

  it('returns the style with the exact id when the document has it', () => {
    const styles = [...localized, style('Heading1', 'Something else')];
    expect(findStyleById(styles, 'Heading1').id).toBe('Heading1');
  });

  it('falls back to the localized style with the same built-in name', () => {
    expect(findStyleById(localized, 'Heading1').id).toBe('Ttulo1');
    expect(findStyleById(localized, 'Title').id).toBe('Ttulo');
  });

  it('compares built-in names case-insensitively', () => {
    expect(findStyleById([style('Ttulo1', 'Heading 1')], 'Heading1').id).toBe('Ttulo1');
  });

  it('does not alias ids that are not SuperDoc defaults', () => {
    expect(findStyleById(localized, 'MyStyle')).toBeUndefined();
  });

  it('returns undefined when the built-in style is absent, and for empty input', () => {
    expect(findStyleById(localized, 'Heading2')).toBeUndefined();
    expect(findStyleById(null, 'Heading1')).toBeUndefined();
    expect(findStyleById(localized, null)).toBeUndefined();
  });
});
