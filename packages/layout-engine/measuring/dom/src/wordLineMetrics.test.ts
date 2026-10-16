import { describe, expect, it } from 'vitest';
import { getWordSingleLineRatio, quantizeWordLineHeight } from './wordLineMetrics.js';

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

describe('quantizeWordLineHeight', () => {
  it('rounds up to 0.1px (1/960 inch)', () => {
    expect(quantizeWordLineHeight(19.531)).toBe(19.6);
    expect(quantizeWordLineHeight(18.4)).toBe(18.4);
  });

  it('does not step up on float noise just above a step', () => {
    // Times New Roman 11pt at 1.08: Word gives 13.65pt (18.2px) for a raw 18.2004px
    expect(quantizeWordLineHeight(18.2004)).toBe(18.2);
  });
});
