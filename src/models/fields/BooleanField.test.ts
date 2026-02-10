import { BooleanField } from './BooleanField';
import { FieldKind } from './FieldBase';

describe('BooleanField', () => {
  const makeField = (
    value: boolean | null,
    overrides?: Partial<{ description: string; isRequired: boolean }>
  ): BooleanField => {
    return new BooleanField(
      'field-id',
      'IsActive',
      'Is Active',
      overrides?.description ?? '',
      overrides?.isRequired ?? false,
      value
    );
  };

  describe('fieldKind', () => {
    it('returns FieldKind.Boolean', () => {
      const field = makeField(true);
      expect(field.fieldKind).toBe(FieldKind.Boolean);
    });
  });

  describe('formatForDisplay', () => {
    it('returns "Yes" when value is true', () => {
      const field = makeField(true);
      expect(field.formatForDisplay()).toBe('Yes');
    });

    it('returns "No" when value is false', () => {
      const field = makeField(false);
      expect(field.formatForDisplay()).toBe('No');
    });

    it('returns "(empty)" when value is null', () => {
      const field = makeField(null);
      expect(field.formatForDisplay()).toBe('(empty)');
    });
  });

  describe('serializeForSharePoint', () => {
    it('returns true as-is', () => {
      const field = makeField(true);
      expect(field.serializeForSharePoint()).toBe(true);
    });

    it('returns false as-is', () => {
      const field = makeField(false);
      expect(field.serializeForSharePoint()).toBe(false);
    });

    it('returns null when value is null', () => {
      const field = makeField(null);
      expect(field.serializeForSharePoint()).toBeNull();
    });
  });
});
