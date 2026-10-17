import { describe, expect, it } from 'vitest';
import { getWordSingleLineRatio } from './wordLineMetrics.js';

describe('getWordSingleLineRatio', () => {
  it('reads the first family of a CSS font list, ignoring quotes and case', () => {
    expect(getWordSingleLineRatio('"Calibri", sans-serif')).toBe(1.2207);
    expect(getWordSingleLineRatio('TIMES NEW ROMAN')).toBe(1.1499);
  });

  it('maps metric-compatible substitutes to the original font', () => {
    expect(getWordSingleLineRatio('Caladea')).toBe(getWordSingleLineRatio('Cambria'));
  });

  it('returns undefined for unknown fonts', () => {
    expect(getWordSingleLineRatio('Some Unknown Font')).toBeUndefined();
    expect(getWordSingleLineRatio(undefined)).toBeUndefined();
  });
});
