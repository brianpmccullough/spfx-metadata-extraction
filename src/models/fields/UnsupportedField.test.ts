import { UnsupportedField } from './UnsupportedField';
import { FieldKind } from './FieldBase';

describe('UnsupportedField', () => {
  const makeField = (
    value: unknown,
    originalType: string
  ): UnsupportedField => {
    return new UnsupportedField(
      'field-id',
      'LookupField',
      'Related Item',
      'Links to another item',
      false,
      value,
      originalType
    );
  };

  describe('fieldKind', () => {
    it('returns FieldKind.Unsupported', () => {
      const field = makeField({ Id: 1, Title: 'Item' }, 'Lookup');
      expect(field.fieldKind).toBe(FieldKind.Unsupported);
    });
  });

  describe('originalType', () => {
    it('stores the original SharePoint field type', () => {
      const field = makeField(null, 'User');
      expect(field.originalType).toBe('User');
    });
  });

  describe('formatForDisplay', () => {
    it('returns original type in brackets when value exists', () => {
      const field = makeField({ some: 'value' }, 'Lookup');
      expect(field.formatForDisplay()).toBe('[Lookup]');
    });

    it('returns "(empty)" when value is null', () => {
      const field = makeField(null, 'User');
      expect(field.formatForDisplay()).toBe('(empty)');
    });

    it('returns "(empty)" when value is undefined', () => {
      const field = makeField(undefined, 'Calculated');
      expect(field.formatForDisplay()).toBe('(empty)');
    });
  });

  describe('serializeForSharePoint', () => {
    it('always returns null', () => {
      const field = makeField({ Id: 1, Title: 'Item' }, 'Lookup');
      expect(field.serializeForSharePoint()).toBeNull();
    });

    it('returns null even when value exists', () => {
      const field = makeField('some value', 'Computed');
      expect(field.serializeForSharePoint()).toBeNull();
    });
  });
});
