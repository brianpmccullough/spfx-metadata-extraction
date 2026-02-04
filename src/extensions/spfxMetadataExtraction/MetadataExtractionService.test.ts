import { MetadataExtractionService } from './MetadataExtractionService';
import type { IDocumentContext } from '../../models/IDocumentContext';
import type { ISharePointRestClient } from '../../clients/ISharePointRestClient';
import type { ITaxonomyService } from '../../services/ITaxonomyService';
import { FieldKind, StringField, ChoiceField, NumericField, BooleanField, UnsupportedField, TaxonomyField, TaxonomyMultiField } from '../../models/fields';

function makeMockDocumentContext(overrides?: Partial<IDocumentContext>): IDocumentContext {
  return {
    contentTypeId: '0x0101',
    fileExtension: 'docx',
    fileName: 'report.docx',
    fileLeafRef: 'report.docx',
    fileRef: '/sites/TestSite/Documents/report.docx',
    fileSize: 204800,
    fileSizeInBytes: 204800,
    fileSizeInKiloBytes: 200,
    fileSizeInMegaBytes: 0.1953,
    fileType: 'docx',
    serverRelativeUrl: '/sites/TestSite/Documents/report.docx',
    itemId: 42,
    uniqueId: '{AAAA-BBBB-CCCC-DDDD}',
    driveId: 'driveABC123',
    driveItemId: 'itemDEF456',
    spItemUrl: 'https://graph.microsoft.com/v1.0/drives/driveABC123/items/itemDEF456',
    siteUrl: 'https://contoso.sharepoint.com/sites/TestSite',
    webUrl: 'https://contoso.sharepoint.com/sites/TestSite',
    siteId: 'site-guid-1234',
    webId: 'web-guid-5678',
    listId: 'list-guid-9012',
    ...overrides,
  };
}

interface IMockClientConfig {
  getResponses?: unknown[];
  postResponses?: unknown[];
}

function makeMockSPOClient(config: IMockClientConfig = {}): ISharePointRestClient {
  const mockGet = jest.fn();
  const mockPost = jest.fn();

  if (config.getResponses) {
    config.getResponses.forEach((response) => {
      mockGet.mockResolvedValueOnce(response);
    });
  } else {
    mockGet.mockResolvedValue({ value: [] });
  }

  if (config.postResponses) {
    config.postResponses.forEach((response) => {
      mockPost.mockResolvedValueOnce(response);
    });
  } else {
    mockPost.mockResolvedValue({ Row: [] });
  }

  return { get: mockGet, post: mockPost };
}

function makeMockTaxonomyService(): ITaxonomyService {
  return {
    getTerms: jest.fn().mockResolvedValue([]),
  };
}

describe('MetadataExtractionService', () => {
  describe('loadFields', () => {
    it('constructs the correct REST URL for field schemas', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [{ value: [] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());
      const docContext = makeMockDocumentContext();

      await service.loadFields(docContext);

      expect(spoClient.get).toHaveBeenCalledWith(
        expect.stringContaining("https://contoso.sharepoint.com/sites/TestSite/_api/web/lists(guid'list-guid-9012')/contenttypes('0x0101')/fields")
      );
    });

    it('includes required field schema properties in select', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [{ value: [] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      await service.loadFields(makeMockDocumentContext());

      const url = (spoClient.get as jest.Mock).mock.calls[0][0];
      expect(url).toContain('$select=');
      expect(url).toContain('Choices');
      expect(url).toContain('DisplayFormat');
      expect(url).toContain('TermSetId');
      expect(url).toContain('SspId');
    });

    it('uses RenderListDataAsStream to get field values', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Notes', Title: 'Notes', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
            ],
          },
        ],
        postResponses: [{ Row: [{ Notes: 'Sample note' }] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      await service.loadFields(makeMockDocumentContext());

      expect(spoClient.post).toHaveBeenCalledWith(
        expect.stringContaining('RenderListDataAsStream'),
        expect.objectContaining({
          parameters: expect.objectContaining({
            ViewXml: expect.stringContaining('<FieldRef Name="Notes"'),
          }),
        }),
        expect.objectContaining({
          Accept: 'application/json;odata=verbose',
        })
      );
    });

    it('filters by item ID using CAML query in ViewXml', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Notes', Title: 'Notes', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
            ],
          },
        ],
        postResponses: [{ Row: [{ Notes: 'Sample note' }] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      await service.loadFields(makeMockDocumentContext());

      const postBody = (spoClient.post as jest.Mock).mock.calls[0][1];
      expect(postBody.parameters.ViewXml).toContain('<FieldRef Name="ID"');
      expect(postBody.parameters.ViewXml).toContain('<Value Type="Counter">42</Value>');
    });

    it('creates StringField for Text type', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Notes', Title: 'Notes', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: 'Some notes' },
            ],
          },
        ],
        postResponses: [{ Row: [{ Notes: 'Sample note content' }] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields).toHaveLength(1);
      expect(fields[0]).toBeInstanceOf(StringField);
      expect(fields[0].fieldKind).toBe(FieldKind.String);
      expect(fields[0].value).toBe('Sample note content');
    });

    it('creates ChoiceField with choices loaded', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Status', Title: 'Status', TypeAsString: 'Choice', Required: false, ReadOnlyField: false, Description: '', Choices: { results: ['Draft', 'Review', 'Final'] } },
            ],
          },
        ],
        postResponses: [{ Row: [{ Status: 'Review' }] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields[0]).toBeInstanceOf(ChoiceField);
      expect(fields[0].fieldKind).toBe(FieldKind.Choice);
      expect((fields[0] as ChoiceField).choices).toEqual(['Draft', 'Review', 'Final']);
      expect(fields[0].value).toBe('Review');
    });

    it('creates NumericField for Number type', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Count', Title: 'Count', TypeAsString: 'Number', Required: false, ReadOnlyField: false, Description: '' },
            ],
          },
        ],
        postResponses: [{ Row: [{ Count: 42 }] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields[0]).toBeInstanceOf(NumericField);
      expect(fields[0].fieldKind).toBe(FieldKind.Numeric);
      expect(fields[0].value).toBe(42);
    });

    it('creates BooleanField for Boolean type', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Active', Title: 'Active', TypeAsString: 'Boolean', Required: false, ReadOnlyField: false, Description: '' },
            ],
          },
        ],
        postResponses: [{ Row: [{ Active: '1' }] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields[0]).toBeInstanceOf(BooleanField);
      expect(fields[0].fieldKind).toBe(FieldKind.Boolean);
      // RenderListDataAsStream returns "1" for true
      expect(fields[0].value).toBe('1');
    });

    it('creates UnsupportedField for unknown types', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Custom', Title: 'Custom', TypeAsString: 'Lookup', Required: false, ReadOnlyField: false, Description: '' },
            ],
          },
        ],
        postResponses: [{ Row: [{ Custom: 'lookup value' }] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields[0]).toBeInstanceOf(UnsupportedField);
      expect(fields[0].fieldKind).toBe(FieldKind.Unsupported);
    });

    it('returns an empty array when no fields exist', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [{ value: [] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields).toEqual([]);
    });

    it('filters out fields with internal names starting with underscore', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: '_ModerationStatus', Title: 'Moderation Status', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
              { Id: 'f2', InternalName: 'CustomField', Title: 'Custom Field', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
              { Id: 'f3', InternalName: '_UIVersionString', Title: 'UI Version', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
            ],
          },
        ],
        postResponses: [{ Row: [{ CustomField: 'value' }] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields).toHaveLength(1);
      expect(fields[0].internalName).toBe('CustomField');
    });

    it('filters out read-only fields', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'ReadOnlyField', Title: 'Read Only', TypeAsString: 'Text', Required: false, ReadOnlyField: true, Description: '' },
              { Id: 'f2', InternalName: 'EditableField', Title: 'Editable', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
            ],
          },
        ],
        postResponses: [{ Row: [{ EditableField: 'value' }] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields).toHaveLength(1);
      expect(fields[0].internalName).toBe('EditableField');
    });

    it('filters out excluded system fields', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
              { Id: 'f2', InternalName: 'Modified', Title: 'Modified', TypeAsString: 'DateTime', Required: false, ReadOnlyField: false, Description: '' },
              { Id: 'f3', InternalName: 'CustomField', Title: 'Custom', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
            ],
          },
        ],
        postResponses: [{ Row: [{ CustomField: 'value' }] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields).toHaveLength(1);
      expect(fields[0].internalName).toBe('CustomField');
    });

    it('propagates errors from the SPO client get', async () => {
      const spoClient: ISharePointRestClient = {
        get: jest.fn().mockRejectedValue(new Error('HTTP 403: Forbidden')),
        post: jest.fn(),
      };
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      await expect(service.loadFields(makeMockDocumentContext()))
        .rejects.toThrow('HTTP 403: Forbidden');
    });

    it('propagates errors from the SPO client post', async () => {
      const spoClient: ISharePointRestClient = {
        get: jest.fn().mockResolvedValue({
          value: [
            { Id: 'f1', InternalName: 'Notes', Title: 'Notes', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
          ],
        }),
        post: jest.fn().mockRejectedValue(new Error('HTTP 500: Internal Server Error')),
      };
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      await expect(service.loadFields(makeMockDocumentContext()))
        .rejects.toThrow('HTTP 500: Internal Server Error');
    });

    it('parses single taxonomy field from RenderListDataAsStream format', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Category', Title: 'Category', TypeAsString: 'TaxonomyFieldType', Required: false, ReadOnlyField: false, Description: '', TermSetId: 'ts-1', SspId: 'ssp-1' },
            ],
          },
        ],
        postResponses: [
          {
            Row: [
              {
                Category: 'Engineering',
                'Category.': '5;#Engineering|term-guid-123',
              },
            ],
          },
        ],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields[0]).toBeInstanceOf(TaxonomyField);
      expect(fields[0].value).toEqual({
        termGuid: 'term-guid-123',
        label: 'Engineering',
        wssId: 5,
      });
    });

    it('parses multi taxonomy field from RenderListDataAsStream format', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Tags', Title: 'Tags', TypeAsString: 'TaxonomyFieldTypeMulti', Required: false, ReadOnlyField: false, Description: '', TermSetId: 'ts-1', SspId: 'ssp-1' },
            ],
          },
        ],
        postResponses: [
          {
            Row: [
              {
                Tags: 'Tag One; Tag Two',
                'Tags.': '10;#Tag One|guid-1;#20;#Tag Two|guid-2',
              },
            ],
          },
        ],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields[0]).toBeInstanceOf(TaxonomyMultiField);
      expect(fields[0].value).toEqual([
        { termGuid: 'guid-1', label: 'Tag One', wssId: 10 },
        { termGuid: 'guid-2', label: 'Tag Two', wssId: 20 },
      ]);
    });

    it('handles null taxonomy field value', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Category', Title: 'Category', TypeAsString: 'TaxonomyFieldType', Required: false, ReadOnlyField: false, Description: '', TermSetId: 'ts-1', SspId: 'ssp-1' },
            ],
          },
        ],
        postResponses: [
          {
            Row: [
              {
                Category: '',
                // No 'Category.' property when empty
              },
            ],
          },
        ],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields[0]).toBeInstanceOf(TaxonomyField);
      expect(fields[0].value).toBeNull();
    });

    it('handles empty Row array from RenderListDataAsStream', async () => {
      const spoClient = makeMockSPOClient({
        getResponses: [
          {
            value: [
              { Id: 'f1', InternalName: 'Notes', Title: 'Notes', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
            ],
          },
        ],
        postResponses: [{ Row: [] }],
      });
      const service = new MetadataExtractionService(spoClient, makeMockTaxonomyService());

      const fields = await service.loadFields(makeMockDocumentContext());

      expect(fields[0]).toBeInstanceOf(StringField);
      expect(fields[0].value).toBeUndefined();
    });
  });
});
