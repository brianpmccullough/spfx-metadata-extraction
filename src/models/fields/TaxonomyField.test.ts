import { TaxonomyField, TaxonomyMultiField, ITaxonomyValue, ITerm } from './TaxonomyField';
import { FieldKind } from './FieldBase';

describe('TaxonomyField', () => {
  const terms: ITerm[] = [
    { termGuid: 'guid-1', label: 'Finance' },
    { termGuid: 'guid-2', label: 'HR' },
    { termGuid: 'guid-3', label: 'IT' },
  ];

  const makeField = (
    value: ITaxonomyValue | null,
    overrides?: Partial<{ description: string; isRequired: boolean; terms: ITerm[] }>
  ): TaxonomyField => {
    return new TaxonomyField(
      'field-id',
      'Department',
      'Department',
      overrides?.description ?? '',
      overrides?.isRequired ?? false,
      value,
      'termset-guid',
      overrides?.terms ?? terms
    );
  };

  describe('fieldKind', () => {
    it('returns FieldKind.Taxonomy', () => {
      const field = makeField({ termGuid: 'guid-1', label: 'Finance' });
      expect(field.fieldKind).toBe(FieldKind.Taxonomy);
    });
  });

  describe('formatForDisplay', () => {
    it('returns the term label when set', () => {
      const field = makeField({ termGuid: 'guid-2', label: 'HR' });
      expect(field.formatForDisplay()).toBe('HR');
    });

    it('returns "(empty)" when value is null', () => {
      const field = makeField(null);
      expect(field.formatForDisplay()).toBe('(empty)');
    });
  });

  describe('serializeForSharePoint', () => {
    it('returns SharePoint taxonomy format with metadata', () => {
      const field = makeField({ termGuid: 'guid-1', label: 'Finance', wssId: 42 });
      expect(field.serializeForSharePoint()).toEqual({
        __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
        Label: 'Finance',
        TermGuid: 'guid-1',
        WssId: 42,
      });
    });

    it('uses WssId -1 when wssId is not provided', () => {
      const field = makeField({ termGuid: 'guid-1', label: 'Finance' });
      const serialized = field.serializeForSharePoint() as { WssId: number };
      expect(serialized.WssId).toBe(-1);
    });

    it('returns null when value is null', () => {
      const field = makeField(null);
      expect(field.serializeForSharePoint()).toBeNull();
    });
  });
});

describe('TaxonomyMultiField', () => {
  const terms: ITerm[] = [
    { termGuid: 'guid-1', label: 'Policy' },
    { termGuid: 'guid-2', label: 'Procedure' },
    { termGuid: 'guid-3', label: 'Guideline' },
  ];

  const makeField = (
    value: ITaxonomyValue[] | null,
    overrides?: Partial<{ description: string; isRequired: boolean; terms: ITerm[] }>
  ): TaxonomyMultiField => {
    return new TaxonomyMultiField(
      'field-id',
      'DocumentType',
      'Document Type',
      overrides?.description ?? '',
      overrides?.isRequired ?? false,
      value,
      'termset-guid',
      overrides?.terms ?? terms
    );
  };

  describe('fieldKind', () => {
    it('returns FieldKind.TaxonomyMulti', () => {
      const field = makeField([{ termGuid: 'guid-1', label: 'Policy' }]);
      expect(field.fieldKind).toBe(FieldKind.TaxonomyMulti);
    });
  });

  describe('formatForDisplay', () => {
    it('returns comma-separated labels', () => {
      const field = makeField([
        { termGuid: 'guid-1', label: 'Policy' },
        { termGuid: 'guid-2', label: 'Procedure' },
      ]);
      expect(field.formatForDisplay()).toBe('Policy, Procedure');
    });

    it('returns "(empty)" when value is null', () => {
      const field = makeField(null);
      expect(field.formatForDisplay()).toBe('(empty)');
    });

    it('returns "(empty)" when value is empty array', () => {
      const field = makeField([]);
      expect(field.formatForDisplay()).toBe('(empty)');
    });
  });

  describe('serializeForSharePoint', () => {
    it('returns array of SharePoint taxonomy format objects', () => {
      const field = makeField([
        { termGuid: 'guid-1', label: 'Policy', wssId: 10 },
        { termGuid: 'guid-2', label: 'Procedure', wssId: 11 },
      ]);
      expect(field.serializeForSharePoint()).toEqual([
        {
          __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
          Label: 'Policy',
          TermGuid: 'guid-1',
          WssId: 10,
        },
        {
          __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
          Label: 'Procedure',
          TermGuid: 'guid-2',
          WssId: 11,
        },
      ]);
    });

    it('returns null when value is null', () => {
      const field = makeField(null);
      expect(field.serializeForSharePoint()).toBeNull();
    });

    it('returns null when value is empty array', () => {
      const field = makeField([]);
      expect(field.serializeForSharePoint()).toBeNull();
    });
  });
});
