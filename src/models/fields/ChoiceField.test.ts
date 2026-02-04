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

  describe('generateLlmPrompt', () => {
    it('includes valid choices', () => {
      const field = makeField('Draft');
      const prompt = field.generateLlmPrompt();
      expect(prompt).toContain('Valid choices: ["Draft", "Review", "Final"]');
    });

    it('includes "(required)" when field is required', () => {
      const field = makeField('Draft', { isRequired: true });
      expect(field.generateLlmPrompt()).toContain('(required)');
    });

    it('describes as single-choice field', () => {
      const field = makeField('Draft');
      expect(field.generateLlmPrompt()).toContain('single-choice field');
    });

    it('includes current value', () => {
      const field = makeField('Review');
      expect(field.generateLlmPrompt()).toContain('Current value: Review');
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

  describe('generateLlmPrompt', () => {
    it('includes valid choices', () => {
      const field = makeField(['Red']);
      const prompt = field.generateLlmPrompt();
      expect(prompt).toContain('Valid choices: ["Red", "Green", "Blue", "Yellow"]');
    });

    it('describes as multi-choice field', () => {
      const field = makeField(['Red']);
      expect(field.generateLlmPrompt()).toContain('multi-choice field');
    });

    it('includes current values', () => {
      const field = makeField(['Red', 'Blue']);
      expect(field.generateLlmPrompt()).toContain('Current value: Red, Blue');
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
});
