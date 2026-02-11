import { StringField } from './StringField';
import { FieldKind } from './FieldBase';

describe('StringField', () => {
  const makeField = (value: string | null, overrides?: Partial<{ description: string; isRequired: boolean; maxLength: number | null }>): StringField => {
    return new StringField(
      'field-id',
      'InternalName',
      'Display Title',
      overrides?.description ?? '',
      overrides?.isRequired ?? false,
      value,
      overrides?.maxLength
    );
  };

  describe('fieldKind', () => {
    it('returns FieldKind.String', () => {
      const field = makeField('test');
      expect(field.fieldKind).toBe(FieldKind.String);
    });
  });

  describe('formatForDisplay', () => {
    it('returns the value when set', () => {
      const field = makeField('Hello World');
      expect(field.formatForDisplay()).toBe('Hello World');
    });

    it('returns "(empty)" when value is null', () => {
      const field = makeField(null);
      expect(field.formatForDisplay()).toBe('(empty)');
    });

    it('returns empty string as-is (not "(empty)")', () => {
      const field = makeField('');
      expect(field.formatForDisplay()).toBe('');
    });
  });

  describe('serializeForSharePoint', () => {
    it('returns the value as-is', () => {
      const field = makeField('test value');
      expect(field.serializeForSharePoint()).toBe('test value');
    });

    it('returns null when value is null', () => {
      const field = makeField(null);
      expect(field.serializeForSharePoint()).toBeNull();
    });
  });

  describe('isValidExtractedValue', () => {
    it('returns true for any value when maxLength is null', () => {
      const field = makeField('test', { maxLength: null });
      expect(field.isValidExtractedValue('a'.repeat(1000))).toBe(true);
    });

    it('returns true for any value when maxLength is undefined (defaults to null)', () => {
      const field = makeField('test');
      expect(field.isValidExtractedValue('anything')).toBe(true);
    });

    it('returns true when value is within maxLength', () => {
      const field = makeField('test', { maxLength: 10 });
      expect(field.isValidExtractedValue('short')).toBe(true);
    });

    it('returns true when value is exactly at maxLength', () => {
      const field = makeField('test', { maxLength: 5 });
      expect(field.isValidExtractedValue('abcde')).toBe(true);
    });

    it('returns false when value exceeds maxLength', () => {
      const field = makeField('test', { maxLength: 5 });
      expect(field.isValidExtractedValue('abcdef')).toBe(false);
    });

    it('converts non-string values to string for length check', () => {
      const field = makeField('test', { maxLength: 3 });
      expect(field.isValidExtractedValue(12345)).toBe(false);
      expect(field.isValidExtractedValue(12)).toBe(true);
    });
  });

  describe('resolveValueForApply', () => {
    it('returns the value as-is', () => {
      const field = makeField('test');
      expect(field.resolveValueForApply('extracted text')).toBe('extracted text');
    });
  });
});
