import { ChoiceField, MultiChoiceField } from './ChoiceField';
import { FieldKind } from './FieldBase';

describe('ChoiceField', () => {
  const choices = ['Draft', 'Review', 'Final'];

  const makeField = (
    value: string | null,
    overrides?: Partial<{ description: string; isRequired: boolean; choices: string[] }>
  ): ChoiceField => {
    return new ChoiceField(
      'field-id',
      'Status',
      'Document Status',
      overrides?.description ?? '',
      overrides?.isRequired ?? false,
      value,
      overrides?.choices ?? choices
    );
  };

  describe('fieldKind', () => {
    it('returns FieldKind.Choice', () => {
      const field = makeField('Draft');
      expect(field.fieldKind).toBe(FieldKind.Choice);
    });
  });

  describe('formatForDisplay', () => {
    it('returns the value when set', () => {
      const field = makeField('Review');
      expect(field.formatForDisplay()).toBe('Review');
    });

    it('returns "(empty)" when value is null', () => {
      const field = makeField(null);
      expect(field.formatForDisplay()).toBe('(empty)');
    });
  });

  describe('serializeForSharePoint', () => {
    it('returns the value as-is', () => {
      const field = makeField('Final');
      expect(field.serializeForSharePoint()).toBe('Final');
    });

    it('returns null when value is null', () => {
      const field = makeField(null);
      expect(field.serializeForSharePoint()).toBeNull();
    });
  });

  describe('isValidExtractedValue', () => {
    it('returns true for a matching choice', () => {
      const field = makeField('Draft');
      expect(field.isValidExtractedValue('Draft')).toBe(true);
    });

    it('matches case-insensitively', () => {
      const field = makeField('Draft');
      expect(field.isValidExtractedValue('draft')).toBe(true);
    });

    it('returns false for a non-matching value', () => {
      const field = makeField('Draft');
      expect(field.isValidExtractedValue('Unknown')).toBe(false);
    });
  });

  describe('resolveValueForApply', () => {
    it('returns the value as-is', () => {
      const field = makeField('Draft');
      expect(field.resolveValueForApply('Final')).toBe('Final');
    });
  });
});

describe('MultiChoiceField', () => {
  const choices = ['Red', 'Green', 'Blue', 'Yellow'];

  const makeField = (
    value: string[] | null,
    overrides?: Partial<{ description: string; isRequired: boolean; choices: string[] }>
  ): MultiChoiceField => {
    return new MultiChoiceField(
      'field-id',
      'Colors',
      'Favorite Colors',
      overrides?.description ?? '',
      overrides?.isRequired ?? false,
      value,
      overrides?.choices ?? choices
    );
  };

  describe('fieldKind', () => {
    it('returns FieldKind.MultiChoice', () => {
      const field = makeField(['Red']);
      expect(field.fieldKind).toBe(FieldKind.MultiChoice);
    });
  });

  describe('formatForDisplay', () => {
    it('returns comma-separated values', () => {
      const field = makeField(['Red', 'Blue']);
      expect(field.formatForDisplay()).toBe('Red, Blue');
    });

    it('returns "(empty)" when value is null', () => {
      const field = makeField(null);
      expect(field.formatForDisplay()).toBe('(empty)');
    });

    it('returns "(empty)" when value is empty array', () => {
      const field = makeField([]);
      expect(field.formatForDisplay()).toBe('(empty)');
    });

    it('returns single value without comma', () => {
      const field = makeField(['Green']);
      expect(field.formatForDisplay()).toBe('Green');
    });
  });

  describe('serializeForSharePoint', () => {
    it('returns SharePoint multi-choice format', () => {
      const field = makeField(['Red', 'Blue']);
      expect(field.serializeForSharePoint()).toBe(';#Red;#Blue;#');
    });

    it('returns null when value is null', () => {
      const field = makeField(null);
      expect(field.serializeForSharePoint()).toBeNull();
    });

    it('returns null when value is empty array', () => {
      const field = makeField([]);
      expect(field.serializeForSharePoint()).toBeNull();
    });

    it('handles single value', () => {
      const field = makeField(['Green']);
      expect(field.serializeForSharePoint()).toBe(';#Green;#');
    });
  });

  describe('isValidExtractedValue', () => {
    it('returns true when all values match choices', () => {
      const field = makeField(['Red']);
      expect(field.isValidExtractedValue('Red, Blue')).toBe(true);
    });

    it('matches case-insensitively', () => {
      const field = makeField(['Red']);
      expect(field.isValidExtractedValue('red, blue')).toBe(true);
    });

    it('returns false when any value does not match', () => {
      const field = makeField(['Red']);
      expect(field.isValidExtractedValue('Red, Purple')).toBe(false);
    });
  });

  describe('resolveValueForApply', () => {
    it('returns the value as-is', () => {
      const field = makeField(['Red']);
      expect(field.resolveValueForApply('Red, Blue')).toBe('Red, Blue');
    });
  });
});
