import { StringField } from './StringField';
import { FieldKind } from './FieldBase';

describe('StringField', () => {
  const makeField = (value: string | null, overrides?: Partial<{ description: string; isRequired: boolean }>): StringField => {
    return new StringField(
      'field-id',
      'InternalName',
      'Display Title',
      overrides?.description ?? '',
      overrides?.isRequired ?? false,
      value
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
});
