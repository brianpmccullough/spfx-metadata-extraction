import { FieldFactory, ISharePointFieldSchema } from './FieldFactory';
import { FieldKind } from './FieldBase';
import { StringField } from './StringField';
import { ChoiceField, MultiChoiceField } from './ChoiceField';
import { TaxonomyField, TaxonomyMultiField } from './TaxonomyField';
import { NumericField } from './NumericField';
import { BooleanField } from './BooleanField';
import { DateTimeField } from './DateTimeField';
import { UnsupportedField } from './UnsupportedField';
import type { ITaxonomyService } from '../../services/ITaxonomyService';

describe('FieldFactory', () => {
  const siteUrl = 'https://contoso.sharepoint.com/sites/TestSite';

  const makeMockTaxonomyService = (): ITaxonomyService => ({
    getTerms: jest.fn().mockResolvedValue([
      { termGuid: 'term-1', label: 'Term One' },
      { termGuid: 'term-2', label: 'Term Two' },
    ]),
  });

  const makeBaseSchema = (
    typeAsString: string,
    overrides?: Partial<ISharePointFieldSchema>
  ): ISharePointFieldSchema => ({
    Id: 'field-id',
    InternalName: 'FieldName',
    Title: 'Field Title',
    Description: 'Field description',
    TypeAsString: typeAsString,
    Required: false,
    ReadOnlyField: false,
    ...overrides,
  });

  describe('Text and Note fields', () => {
    it('creates StringField for Text type', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Text');

      const field = await factory.createField(schema, 'test value', siteUrl);

      expect(field).toBeInstanceOf(StringField);
      expect(field.fieldKind).toBe(FieldKind.String);
      expect(field.value).toBe('test value');
    });

    it('creates StringField for Note type', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Note');

      const field = await factory.createField(schema, 'multiline text', siteUrl);

      expect(field).toBeInstanceOf(StringField);
      expect(field.fieldKind).toBe(FieldKind.String);
    });

    it('handles null value', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Text');

      const field = await factory.createField(schema, null, siteUrl);

      expect(field.value).toBeNull();
    });
  });

  describe('Choice fields', () => {
    it('creates ChoiceField with choices', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Choice', {
        Choices: { results: ['A', 'B', 'C'] },
      });

      const field = await factory.createField(schema, 'B', siteUrl);

      expect(field).toBeInstanceOf(ChoiceField);
      expect(field.fieldKind).toBe(FieldKind.Choice);
      expect((field as ChoiceField).choices).toEqual(['A', 'B', 'C']);
      expect(field.value).toBe('B');
    });

    it('handles missing Choices property', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Choice');

      const field = await factory.createField(schema, null, siteUrl);

      expect((field as ChoiceField).choices).toEqual([]);
    });
  });

  describe('MultiChoice fields', () => {
    it('creates MultiChoiceField with choices', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('MultiChoice', {
        Choices: { results: ['Red', 'Green', 'Blue'] },
      });

      const field = await factory.createField(schema, ['Red', 'Blue'], siteUrl);

      expect(field).toBeInstanceOf(MultiChoiceField);
      expect(field.fieldKind).toBe(FieldKind.MultiChoice);
      expect((field as MultiChoiceField).choices).toEqual(['Red', 'Green', 'Blue']);
      expect(field.value).toEqual(['Red', 'Blue']);
    });

    it('parses SharePoint semicolon-delimited format', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('MultiChoice', {
        Choices: { results: ['Red', 'Green', 'Blue'] },
      });

      const field = await factory.createField(schema, ';#Red;#Blue;#', siteUrl);

      expect(field.value).toEqual(['Red', 'Blue']);
    });

    it('returns null for empty semicolon string', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('MultiChoice');

      const field = await factory.createField(schema, ';#;#', siteUrl);

      expect(field.value).toBeNull();
    });
  });

  describe('Taxonomy fields', () => {
    it('creates TaxonomyField and loads terms', async () => {
      const taxonomyService = makeMockTaxonomyService();
      const factory = new FieldFactory(taxonomyService);
      const schema = makeBaseSchema('TaxonomyFieldType', {
        TermSetId: 'termset-123',
      });

      const field = await factory.createField(
        schema,
        {
          termGuid: 'term-1',
          label: 'Term One',
          wssId: 10,
        },
        siteUrl
      );

      expect(field).toBeInstanceOf(TaxonomyField);
      expect(field.fieldKind).toBe(FieldKind.Taxonomy);
      expect(taxonomyService.getTerms).toHaveBeenCalledWith('termset-123', siteUrl);
      expect((field as TaxonomyField).terms).toHaveLength(2);
      expect(field.value).toEqual({ termGuid: 'term-1', label: 'Term One', wssId: 10 });
    });

    it('skips term loading when TermSetId is missing', async () => {
      const taxonomyService = makeMockTaxonomyService();
      const factory = new FieldFactory(taxonomyService);
      const schema = makeBaseSchema('TaxonomyFieldType');

      const field = await factory.createField(schema, null, siteUrl);

      expect(taxonomyService.getTerms).not.toHaveBeenCalled();
      expect((field as TaxonomyField).terms).toEqual([]);
    });
  });

  describe('TaxonomyMulti fields', () => {
    it('creates TaxonomyMultiField with multiple values', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('TaxonomyFieldTypeMulti', {
        TermSetId: 'termset-123',
      });

      const field = await factory.createField(
        schema,
        [
          { termGuid: 'term-1', label: 'Term One' },
          { termGuid: 'term-2', label: 'Term Two' },
        ],
        siteUrl
      );

      expect(field).toBeInstanceOf(TaxonomyMultiField);
      expect(field.fieldKind).toBe(FieldKind.TaxonomyMulti);
      expect(field.value).toEqual([
        { termGuid: 'term-1', label: 'Term One', wssId: undefined },
        { termGuid: 'term-2', label: 'Term Two', wssId: undefined },
      ]);
    });
  });

  describe('Numeric fields', () => {
    it('creates NumericField for Number type', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Number');

      const field = await factory.createField(schema, 42.5, siteUrl);

      expect(field).toBeInstanceOf(NumericField);
      expect(field.fieldKind).toBe(FieldKind.Numeric);
      expect(field.value).toBe(42.5);
    });

    it('creates NumericField for Currency type', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Currency');

      const field = await factory.createField(schema, 99.99, siteUrl);

      expect(field).toBeInstanceOf(NumericField);
      expect(field.fieldKind).toBe(FieldKind.Numeric);
    });
  });

  describe('Boolean fields', () => {
    it('creates BooleanField', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Boolean');

      const field = await factory.createField(schema, true, siteUrl);

      expect(field).toBeInstanceOf(BooleanField);
      expect(field.fieldKind).toBe(FieldKind.Boolean);
      expect(field.value).toBe(true);
    });
  });

  describe('DateTime fields', () => {
    it('creates DateTimeField with date only format', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('DateTime', { DisplayFormat: 0 });

      const field = await factory.createField(schema, '2024-06-15T00:00:00Z', siteUrl);

      expect(field).toBeInstanceOf(DateTimeField);
      expect(field.fieldKind).toBe(FieldKind.DateTime);
      expect((field as DateTimeField).includesTime).toBe(false);
    });

    it('creates DateTimeField with date and time format', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('DateTime', { DisplayFormat: 1 });

      const field = await factory.createField(schema, '2024-06-15T10:30:00Z', siteUrl);

      expect((field as DateTimeField).includesTime).toBe(true);
    });

    it('handles null value', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('DateTime');

      const field = await factory.createField(schema, null, siteUrl);

      expect(field.value).toBeNull();
    });
  });

  describe('Unsupported fields', () => {
    it('creates UnsupportedField for unknown types', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Lookup');

      const field = await factory.createField(schema, { Id: 1, Title: 'Item' }, siteUrl);

      expect(field).toBeInstanceOf(UnsupportedField);
      expect(field.fieldKind).toBe(FieldKind.Unsupported);
      expect((field as UnsupportedField).originalType).toBe('Lookup');
    });

    it('creates UnsupportedField for User type', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('User');

      const field = await factory.createField(schema, null, siteUrl);

      expect(field).toBeInstanceOf(UnsupportedField);
      expect((field as UnsupportedField).originalType).toBe('User');
    });
  });

  describe('base properties', () => {
    it('sets id from schema', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Text', { Id: 'unique-field-id' });

      const field = await factory.createField(schema, null, siteUrl);

      expect(field.id).toBe('unique-field-id');
    });

    it('sets internalName from schema', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Text', { InternalName: 'CustomField' });

      const field = await factory.createField(schema, null, siteUrl);

      expect(field.internalName).toBe('CustomField');
    });

    it('sets isRequired from schema', async () => {
      const factory = new FieldFactory(makeMockTaxonomyService());
      const schema = makeBaseSchema('Text', { Required: true });

      const field = await factory.createField(schema, null, siteUrl);

      expect(field.isRequired).toBe(true);
    });
  });
});
