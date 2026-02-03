import { MetadataExtractionService } from './MetadataExtractionService';
import type { IDocumentContext } from '../../models/IDocumentContext';
import type { ISharePointRestClient } from '../../clients/ISharePointRestClient';
import type { IGraphClient } from '../../clients/IGraphClient';

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

function makeMockSPOClient(response?: unknown): ISharePointRestClient {
  return {
    get: jest.fn().mockResolvedValue(response ?? { value: [] }),
  };
}

function makeMockGraphClient(): IGraphClient {
  return {
    get: jest.fn().mockResolvedValue({}),
  };
}

describe('MetadataExtractionService', () => {
  describe('getContentTypeFields', () => {
    it('constructs the correct REST URL from document context', async () => {
      const spoClient = makeMockSPOClient();
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());
      const docContext = makeMockDocumentContext();

      await service.getContentTypeFields(docContext);

      const expectedUrl =
        "https://contoso.sharepoint.com/sites/TestSite/_api/web/lists(guid'list-guid-9012')/contenttypes('0x0101')/fields?$filter=Hidden eq false";
      expect(spoClient.get).toHaveBeenCalledWith(expectedUrl);
    });

    it('maps SPO field response to IFieldInfo with camelCase properties', async () => {
      const spoClient = makeMockSPOClient({
        value: [
          {
            Id: 'field-guid-1',
            InternalName: 'CustomCategory',
            Title: 'Category',
            TypeAsString: 'Text',
            Required: true,
            ReadOnlyField: false,
            Description: 'The category of the item',
          },
          {
            Id: 'field-guid-2',
            InternalName: 'Priority',
            Title: 'Priority',
            TypeAsString: 'Choice',
            Required: false,
            ReadOnlyField: false,
            Description: '',
          },
        ],
      });
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());

      const fields = await service.getContentTypeFields(makeMockDocumentContext());

      expect(fields).toHaveLength(2);
      expect(fields[0]).toEqual({
        id: 'field-guid-1',
        internalName: 'CustomCategory',
        title: 'Category',
        typeAsString: 'Text',
        required: true,
        readOnly: false,
        description: 'The category of the item',
      });
      expect(fields[1]).toEqual({
        id: 'field-guid-2',
        internalName: 'Priority',
        title: 'Priority',
        typeAsString: 'Choice',
        required: false,
        readOnly: false,
        description: '',
      });
    });

    it('returns an empty array when no fields are returned', async () => {
      const spoClient = makeMockSPOClient({ value: [] });
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());

      const fields = await service.getContentTypeFields(makeMockDocumentContext());

      expect(fields).toEqual([]);
    });

    it('propagates errors from the SPO client', async () => {
      const spoClient: ISharePointRestClient = {
        get: jest.fn().mockRejectedValue(new Error('HTTP 403: Forbidden')),
      };
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());

      await expect(service.getContentTypeFields(makeMockDocumentContext()))
        .rejects.toThrow('HTTP 403: Forbidden');
    });

    it('uses the correct webUrl, listId, and contentTypeId from context', async () => {
      const spoClient = makeMockSPOClient();
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());
      const docContext = makeMockDocumentContext({
        webUrl: 'https://other.sharepoint.com/sites/OtherSite',
        listId: 'other-list-guid',
        contentTypeId: '0x0120',
      });

      await service.getContentTypeFields(docContext);

      const expectedUrl =
        "https://other.sharepoint.com/sites/OtherSite/_api/web/lists(guid'other-list-guid')/contenttypes('0x0120')/fields?$filter=Hidden eq false";
      expect(spoClient.get).toHaveBeenCalledWith(expectedUrl);
    });

    it('filters out fields with internal names starting with underscore', async () => {
      const spoClient = makeMockSPOClient({
        value: [
          { Id: 'f1', InternalName: '_ModerationStatus', Title: 'Moderation Status', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
          { Id: 'f2', InternalName: 'CustomField', Title: 'Custom Field', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
          { Id: 'f3', InternalName: '_UIVersionString', Title: 'UI Version', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
        ],
      });
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());

      const fields = await service.getContentTypeFields(makeMockDocumentContext());

      expect(fields).toHaveLength(1);
      expect(fields[0].internalName).toBe('CustomField');
    });
  });

  describe('loadFieldMetadata', () => {
    it('maps Text fields to string type and includes field value', async () => {
      const spoClient: ISharePointRestClient = {
        get: jest.fn()
          .mockResolvedValueOnce({
            value: [
              { Id: 'f1', InternalName: 'Notes', Title: 'Notes', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: 'Some notes' },
            ],
          })
          .mockResolvedValueOnce({ Notes: 'Sample note content' }),
      };
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());

      const result = await service.loadFieldMetadata(makeMockDocumentContext());

      expect(result).toEqual([
        { id: 'f1', internalName: 'Notes', title: 'Notes', description: 'Some notes', type: 'string', value: 'Sample note content' },
      ]);
    });

    it('maps Number fields to number type', async () => {
      const spoClient: ISharePointRestClient = {
        get: jest.fn()
          .mockResolvedValueOnce({
            value: [
              { Id: 'f1', InternalName: 'Count', Title: 'Count', TypeAsString: 'Number', Required: false, ReadOnlyField: false, Description: '' },
            ],
          })
          .mockResolvedValueOnce({ Count: 42 }),
      };
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());

      const result = await service.loadFieldMetadata(makeMockDocumentContext());

      expect(result[0].type).toBe('number');
      expect(result[0].value).toBe(42);
    });

    it('maps Currency fields to number type', async () => {
      const spoClient: ISharePointRestClient = {
        get: jest.fn()
          .mockResolvedValueOnce({
            value: [
              { Id: 'f1', InternalName: 'Price', Title: 'Price', TypeAsString: 'Currency', Required: false, ReadOnlyField: false, Description: '' },
            ],
          })
          .mockResolvedValueOnce({ Price: 99.99 }),
      };
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());

      const result = await service.loadFieldMetadata(makeMockDocumentContext());

      expect(result[0].type).toBe('number');
      expect(result[0].value).toBe(99.99);
    });

    it('maps Boolean fields to boolean type', async () => {
      const spoClient: ISharePointRestClient = {
        get: jest.fn()
          .mockResolvedValueOnce({
            value: [
              { Id: 'f1', InternalName: 'Active', Title: 'Active', TypeAsString: 'Boolean', Required: false, ReadOnlyField: false, Description: '' },
            ],
          })
          .mockResolvedValueOnce({ Active: true }),
      };
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());

      const result = await service.loadFieldMetadata(makeMockDocumentContext());

      expect(result[0].type).toBe('boolean');
      expect(result[0].value).toBe(true);
    });

    it('defaults unknown field types to string', async () => {
      const spoClient: ISharePointRestClient = {
        get: jest.fn()
          .mockResolvedValueOnce({
            value: [
              { Id: 'f1', InternalName: 'Custom', Title: 'Custom', TypeAsString: 'Lookup', Required: false, ReadOnlyField: false, Description: '' },
            ],
          })
          .mockResolvedValueOnce({ Custom: 'lookup value' }),
      };
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());

      const result = await service.loadFieldMetadata(makeMockDocumentContext());

      expect(result[0].type).toBe('string');
    });

    it('returns an empty array when no fields exist', async () => {
      const spoClient = makeMockSPOClient({ value: [] });
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());

      const result = await service.loadFieldMetadata(makeMockDocumentContext());

      expect(result).toEqual([]);
    });

    it('sets value to null when field value is undefined', async () => {
      const spoClient: ISharePointRestClient = {
        get: jest.fn()
          .mockResolvedValueOnce({
            value: [
              { Id: 'f1', InternalName: 'Notes', Title: 'Notes', TypeAsString: 'Text', Required: false, ReadOnlyField: false, Description: '' },
            ],
          })
          .mockResolvedValueOnce({}),
      };
      const service = new MetadataExtractionService(spoClient, makeMockGraphClient());

      const result = await service.loadFieldMetadata(makeMockDocumentContext());

      expect(result[0].value).toBeNull();
    });
  });
});
