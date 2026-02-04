import { DateTimeField } from './DateTimeField';
import { FieldKind } from './FieldBase';

describe('DateTimeField', () => {
  const makeField = (
    value: Date | null,
    includesTime: boolean,
    overrides?: Partial<{ description: string; isRequired: boolean }>
  ): DateTimeField => {
    return new DateTimeField(
      'field-id',
      'DueDate',
      'Due Date',
      overrides?.description ?? '',
      overrides?.isRequired ?? false,
      value,
      includesTime
    );
  };

  describe('fieldKind', () => {
    it('returns FieldKind.DateTime', () => {
      const field = makeField(new Date(), false);
      expect(field.fieldKind).toBe(FieldKind.DateTime);
    });
  });

  describe('formatForDisplay', () => {
    it('returns "(empty)" when value is null', () => {
      const field = makeField(null, false);
      expect(field.formatForDisplay()).toBe('(empty)');
    });

    it('calls toLocaleDateString when includesTime is false', () => {
      const date = new Date('2024-06-15T10:30:00Z');
      const field = makeField(date, false);
      expect(field.formatForDisplay()).toBe(date.toLocaleDateString());
    });

    it('calls toLocaleString when includesTime is true', () => {
      const date = new Date('2024-06-15T10:30:00Z');
      const field = makeField(date, true);
      expect(field.formatForDisplay()).toBe(date.toLocaleString());
    });
  });

  describe('generateLlmPrompt', () => {
    it('describes as "date only" when includesTime is false', () => {
      const field = makeField(null, false);
      expect(field.generateLlmPrompt()).toContain('date only field');
    });

    it('describes as "date and time" when includesTime is true', () => {
      const field = makeField(null, true);
      expect(field.generateLlmPrompt()).toContain('date and time field');
    });

    it('includes "(required)" when field is required', () => {
      const field = makeField(null, false, { isRequired: true });
      expect(field.generateLlmPrompt()).toContain('(required)');
    });

    it('includes description when provided', () => {
      const field = makeField(null, false, { description: 'When the task is due' });
      expect(field.generateLlmPrompt()).toContain('When the task is due');
    });

    it('includes current value', () => {
      const date = new Date('2024-06-15T10:30:00Z');
      const field = makeField(date, false);
      expect(field.generateLlmPrompt()).toContain(`Current value: ${date.toLocaleDateString()}`);
    });
  });

  describe('serializeForSharePoint', () => {
    it('returns ISO 8601 string', () => {
      const date = new Date('2024-06-15T10:30:00Z');
      const field = makeField(date, true);
      expect(field.serializeForSharePoint()).toBe('2024-06-15T10:30:00.000Z');
    });

    it('returns null when value is null', () => {
      const field = makeField(null, false);
      expect(field.serializeForSharePoint()).toBeNull();
    });
  });

  describe('includesTime property', () => {
    it('is accessible and returns true when set', () => {
      const field = makeField(new Date(), true);
      expect(field.includesTime).toBe(true);
    });

    it('is accessible and returns false when set', () => {
      const field = makeField(new Date(), false);
      expect(field.includesTime).toBe(false);
    });
  });
});
