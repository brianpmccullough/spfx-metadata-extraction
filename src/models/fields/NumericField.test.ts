import { NumericField } from './NumericField';
import { FieldKind } from './FieldBase';

describe('NumericField', () => {
  const makeField = (
    value: number | null,
    overrides?: Partial<{ description: string; isRequired: boolean }>
  ): NumericField => {
    return new NumericField(
      'field-id',
      'Amount',
      'Total Amount',
      overrides?.description ?? '',
      overrides?.isRequired ?? false,
      value
    );
  };

  describe('fieldKind', () => {
    it('returns FieldKind.Numeric', () => {
      const field = makeField(42);
      expect(field.fieldKind).toBe(FieldKind.Numeric);
    });
  });

  describe('formatForDisplay', () => {
    it('returns the value as string when set', () => {
      const field = makeField(123.45);
      expect(field.formatForDisplay()).toBe('123.45');
    });

    it('returns "(empty)" when value is null', () => {
      const field = makeField(null);
      expect(field.formatForDisplay()).toBe('(empty)');
    });

    it('returns "0" for zero value', () => {
      const field = makeField(0);
      expect(field.formatForDisplay()).toBe('0');
    });

    it('handles negative numbers', () => {
      const field = makeField(-50);
      expect(field.formatForDisplay()).toBe('-50');
    });
  });

  describe('serializeForSharePoint', () => {
    it('returns the value as-is', () => {
      const field = makeField(42.5);
      expect(field.serializeForSharePoint()).toBe(42.5);
    });

    it('returns null when value is null', () => {
      const field = makeField(null);
      expect(field.serializeForSharePoint()).toBeNull();
    });

    it('returns 0 for zero value', () => {
      const field = makeField(0);
      expect(field.serializeForSharePoint()).toBe(0);
    });
  });

  describe('isValidExtractedValue', () => {
    it('returns true for any value', () => {
      const field = makeField(42);
      expect(field.isValidExtractedValue(100)).toBe(true);
    });
  });

  describe('resolveValueForApply', () => {
    it('returns the value as-is', () => {
      const field = makeField(42);
      expect(field.resolveValueForApply(100)).toBe(100);
    });
  });
});
