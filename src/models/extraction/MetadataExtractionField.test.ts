import { MetadataExtractionField, MetadataExtractionFieldType } from './MetadataExtractionField';
import { StringField, NumericField, BooleanField, ChoiceField, DateTimeField, TaxonomyField, TaxonomyMultiField } from '../fields';

describe('MetadataExtractionField', () => {
  describe('constructor', () => {
    it('wraps a FieldBase and exposes it via field property', () => {
      const field = new StringField('id1', 'Notes', 'Notes', 'Some notes', false, 'value');
      const extractionField = new MetadataExtractionField(field);

      expect(extractionField.field).toBe(field);
    });

    it('defaults description to the field description', () => {
      const field = new StringField('id1', 'Notes', 'Notes', 'Field description here', false, 'value');
      const extractionField = new MetadataExtractionField(field);

      expect(extractionField.description).toBe('Field description here');
    });

    it('defaults description to empty string when field has no description', () => {
      const field = new StringField('id1', 'Notes', 'Notes', '', false, 'value');
      const extractionField = new MetadataExtractionField(field);

      expect(extractionField.description).toBe('');
    });
  });

  describe('inferExtractionType', () => {
    it('infers String type for StringField', () => {
      const field = new StringField('id1', 'Notes', 'Notes', '', false, 'value');
      const extractionField = new MetadataExtractionField(field);

      expect(extractionField.extractionType).toBe(MetadataExtractionFieldType.String);
    });

    it('infers Number type for NumericField', () => {
      const field = new NumericField('id1', 'Count', 'Count', '', false, 42);
      const extractionField = new MetadataExtractionField(field);

      expect(extractionField.extractionType).toBe(MetadataExtractionFieldType.Number);
    });

    it('infers Boolean type for BooleanField', () => {
      const field = new BooleanField('id1', 'Active', 'Active', '', false, true);
      const extractionField = new MetadataExtractionField(field);

      expect(extractionField.extractionType).toBe(MetadataExtractionFieldType.Boolean);
    });

    it('infers String type for ChoiceField', () => {
      const field = new ChoiceField('id1', 'Status', 'Status', '', false, 'Draft', ['Draft', 'Final']);
      const extractionField = new MetadataExtractionField(field);

      expect(extractionField.extractionType).toBe(MetadataExtractionFieldType.String);
    });

    it('infers String type for DateTimeField', () => {
      const field = new DateTimeField('id1', 'DueDate', 'Due Date', '', false, new Date(), false);
      const extractionField = new MetadataExtractionField(field);

      expect(extractionField.extractionType).toBe(MetadataExtractionFieldType.String);
    });
  });

  describe('clone', () => {
    it('returns a new instance', () => {
      const field = new StringField('id1', 'Notes', 'Notes', '', false, 'value');
      const original = new MetadataExtractionField(field);

      const cloned = original.clone();

      expect(cloned).not.toBe(original);
      expect(cloned).toBeInstanceOf(MetadataExtractionField);
    });

    it('preserves all mutable properties', () => {
      const field = new NumericField('id1', 'Count', 'Count', 'A count', false, 42);
      const original = new MetadataExtractionField(field);
      original.extractionType = MetadataExtractionFieldType.String;
      original.description = 'Custom description';
      original.extractedValue = 'extracted';
      original.confidence = 'green';

      const cloned = original.clone();

      expect(cloned.field).toBe(original.field);
      expect(cloned.extractionType).toBe(MetadataExtractionFieldType.String);
      expect(cloned.description).toBe('Custom description');
      expect(cloned.extractedValue).toBe('extracted');
      expect(cloned.confidence).toBe('green');
    });

    it('mutations on clone do not affect the original', () => {
      const field = new StringField('id1', 'Notes', 'Notes', '', false, 'value');
      const original = new MetadataExtractionField(field);
      original.extractedValue = 'original value';
      original.confidence = 'green';

      const cloned = original.clone();
      cloned.extractedValue = 'modified value';
      cloned.confidence = 'red';
      cloned.description = 'modified description';

      expect(original.extractedValue).toBe('original value');
      expect(original.confidence).toBe('green');
      expect(original.description).toBe('');
    });
  });

  describe('mutability', () => {
    it('allows extractionType to be changed', () => {
      const field = new StringField('id1', 'Notes', 'Notes', '', false, 'value');
      const extractionField = new MetadataExtractionField(field);

      extractionField.extractionType = MetadataExtractionFieldType.Number;

      expect(extractionField.extractionType).toBe(MetadataExtractionFieldType.Number);
    });

    it('allows description to be changed', () => {
      const field = new StringField('id1', 'Notes', 'Notes', 'Original', false, 'value');
      const extractionField = new MetadataExtractionField(field);

      extractionField.description = 'Updated description for LLM';

      expect(extractionField.description).toBe('Updated description for LLM');
    });

    it('does not affect the underlying field description when changed', () => {
      const field = new StringField('id1', 'Notes', 'Notes', 'Original', false, 'value');
      const extractionField = new MetadataExtractionField(field);

      extractionField.description = 'Updated';

      expect(field.description).toBe('Original');
    });
  });

  describe('canApply', () => {
    it('returns false when extractedValue is null', () => {
      const field = new StringField('id1', 'Title', 'Title', '', false, 'val');
      const ef = new MetadataExtractionField(field);
      ef.confidence = 'green';

      expect(ef.canApply()).toBe(false);
    });

    it('returns false when confidence is null', () => {
      const field = new StringField('id1', 'Title', 'Title', '', false, 'val');
      const ef = new MetadataExtractionField(field);
      ef.extractedValue = 'something';

      expect(ef.canApply()).toBe(false);
    });

    it('returns false when confidence is red', () => {
      const field = new StringField('id1', 'Title', 'Title', '', false, 'val');
      const ef = new MetadataExtractionField(field);
      ef.extractedValue = 'something';
      ef.confidence = 'red';

      expect(ef.canApply()).toBe(false);
    });

    it('returns true for green confidence with valid value', () => {
      const field = new StringField('id1', 'Title', 'Title', '', false, 'val');
      const ef = new MetadataExtractionField(field);
      ef.extractedValue = 'something';
      ef.confidence = 'green';

      expect(ef.canApply()).toBe(true);
    });

    it('returns true for yellow confidence with valid value', () => {
      const field = new StringField('id1', 'Title', 'Title', '', false, 'val');
      const ef = new MetadataExtractionField(field);
      ef.extractedValue = 'something';
      ef.confidence = 'yellow';

      expect(ef.canApply()).toBe(true);
    });

    it('returns false when field rejects the extracted value', () => {
      const field = new TaxonomyField('id1', 'Dept', 'Department', '', false, null, 'ts1', [
        { termGuid: 'g1', label: 'HR' },
      ]);
      const ef = new MetadataExtractionField(field);
      ef.extractedValue = 'Unknown Department';
      ef.confidence = 'green';

      expect(ef.canApply()).toBe(false);
    });

    it('returns true when taxonomy label matches', () => {
      const field = new TaxonomyField('id1', 'Dept', 'Department', '', false, null, 'ts1', [
        { termGuid: 'g1', label: 'HR' },
      ]);
      const ef = new MetadataExtractionField(field);
      ef.extractedValue = 'HR';
      ef.confidence = 'green';

      expect(ef.canApply()).toBe(true);
    });

    it('returns false for choice field with non-matching value', () => {
      const field = new ChoiceField('id1', 'Status', 'Status', '', false, 'Draft', ['Draft', 'Final']);
      const ef = new MetadataExtractionField(field);
      ef.extractedValue = 'Unknown';
      ef.confidence = 'green';

      expect(ef.canApply()).toBe(false);
    });
  });

  describe('resolveValueForApply', () => {
    it('returns null when extractedValue is null', () => {
      const field = new StringField('id1', 'Title', 'Title', '', false, 'val');
      const ef = new MetadataExtractionField(field);

      expect(ef.resolveValueForApply()).toBeNull();
    });

    it('delegates to field for taxonomy resolution', () => {
      const field = new TaxonomyField('id1', 'Dept', 'Department', '', false, null, 'ts1', [
        { termGuid: 'guid-abc-123', label: 'Executive Management' },
      ]);
      const ef = new MetadataExtractionField(field);
      ef.extractedValue = 'Executive Management';

      expect(ef.resolveValueForApply()).toBe('Executive Management|guid-abc-123');
    });

    it('delegates to field for multi-taxonomy resolution', () => {
      const field = new TaxonomyMultiField('id1', 'Tags', 'Tags', '', false, null, 'ts1', [
        { termGuid: 'guid-1', label: 'Finance' },
        { termGuid: 'guid-2', label: 'Legal' },
      ]);
      const ef = new MetadataExtractionField(field);
      ef.extractedValue = 'Finance, Legal';

      expect(ef.resolveValueForApply()).toBe('Finance|guid-1;#Legal|guid-2');
    });

    it('passes through string values for non-taxonomy fields', () => {
      const field = new StringField('id1', 'Title', 'Title', '', false, 'val');
      const ef = new MetadataExtractionField(field);
      ef.extractedValue = 'Some Title';

      expect(ef.resolveValueForApply()).toBe('Some Title');
    });

    it('passes through numeric values', () => {
      const field = new NumericField('id1', 'Count', 'Count', '', false, 0);
      const ef = new MetadataExtractionField(field);
      ef.extractedValue = 42;

      expect(ef.resolveValueForApply()).toBe(42);
    });
  });
});
