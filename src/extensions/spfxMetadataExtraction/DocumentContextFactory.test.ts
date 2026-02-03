import { DocumentContextFactory, IDocumentContextInput } from './DocumentContextFactory';

function makeInput(overrides?: Partial<IDocumentContextInput>): IDocumentContextInput {
  return {
    rowValues: {
      'FileLeafRef': 'report.docx',
      'File_x0020_Type': 'docx',
      'FileRef': '/sites/TestSite/Documents/report.docx',
      'ID': '42',
      'UniqueId': '{AAAA-BBBB-CCCC-DDDD}',
      'File_x0020_Size': '204800',
      'ContentTypeId': '0x0101',
      '.spItemUrl': 'https://graph.microsoft.com/v1.0/drives/driveABC123/items/itemDEF456',
    },
    siteUrl: 'https://contoso.sharepoint.com/sites/TestSite',
    webUrl: 'https://contoso.sharepoint.com/sites/TestSite',
    siteId: 'site-guid-1234',
    webId: 'web-guid-5678',
    listId: 'list-guid-9012',
    ...overrides,
  };
}

describe('DocumentContextFactory', () => {
  let factory: DocumentContextFactory;

  beforeEach(() => {
    factory = new DocumentContextFactory();
  });

  it('maps all fields from a fully-populated input', () => {
    const result = factory.create(makeInput());

    expect(result.fileName).toBe('report.docx');
    expect(result.fileType).toBe('docx');
    expect(result.fileExtension).toBe('docx');
    expect(result.fileLeafRef).toBe('report.docx');
    expect(result.fileRef).toBe('/sites/TestSite/Documents/report.docx');
    expect(result.serverRelativeUrl).toBe('/sites/TestSite/Documents/report.docx');
    expect(result.itemId).toBe(42);
    expect(result.uniqueId).toBe('{AAAA-BBBB-CCCC-DDDD}');
    expect(result.contentTypeId).toBe('0x0101');
    expect(result.siteUrl).toBe('https://contoso.sharepoint.com/sites/TestSite');
    expect(result.webUrl).toBe('https://contoso.sharepoint.com/sites/TestSite');
    expect(result.siteId).toBe('site-guid-1234');
    expect(result.webId).toBe('web-guid-5678');
    expect(result.listId).toBe('list-guid-9012');
  });

  it('extracts driveId and driveItemId from spItemUrl', () => {
    const result = factory.create(makeInput());

    expect(result.spItemUrl).toBe('https://graph.microsoft.com/v1.0/drives/driveABC123/items/itemDEF456');
    expect(result.driveId).toBe('driveABC123');
    expect(result.driveItemId).toBe('itemDEF456');
  });

  it('computes file size conversions correctly', () => {
    const result = factory.create(makeInput());

    expect(result.fileSize).toBe(204800);
    expect(result.fileSizeInBytes).toBe(204800);
    expect(result.fileSizeInKiloBytes).toBe(200);
    expect(result.fileSizeInMegaBytes).toBeCloseTo(0.1953, 4);
  });

  it('defaults to empty strings and zero when row values are missing', () => {
    const result = factory.create(makeInput({ rowValues: {} }));

    expect(result.fileName).toBe('');
    expect(result.fileExtension).toBe('');
    expect(result.serverRelativeUrl).toBe('');
    expect(result.itemId).toBe(0);
    expect(result.uniqueId).toBe('');
    expect(result.fileSize).toBe(0);
    expect(result.fileSizeInBytes).toBe(0);
    expect(result.fileSizeInKiloBytes).toBe(0);
    expect(result.fileSizeInMegaBytes).toBe(0);
    expect(result.driveId).toBe('');
    expect(result.driveItemId).toBe('');
  });

  it('handles missing .spItemUrl field gracefully', () => {
    const result = factory.create(makeInput({
      rowValues: { 'FileLeafRef': 'test.pdf', 'File_x0020_Type': 'pdf' },
    }));

    expect(result.spItemUrl).toBe('');
    expect(result.driveId).toBe('');
    expect(result.driveItemId).toBe('');
  });

  describe('spItemUrl segment parsing', () => {
    it('is case-insensitive for the segment name', () => {
      const url = 'https://graph.microsoft.com/v1.0/drives/DriveID/items/ItemID';
      const result = factory.create(makeInput({
        rowValues: { '.spItemUrl': url },
      }));

      expect(result.driveId).toBe('DriveID');
      expect(result.driveItemId).toBe('ItemID');
    });

    it('preserves original casing of extracted values', () => {
      const result = factory.create(makeInput({
        rowValues: { '.spItemUrl': 'https://graph.microsoft.com/v1.0/drives/AbCdEf/items/GhIjKl' },
      }));

      expect(result.driveId).toBe('AbCdEf');
      expect(result.driveItemId).toBe('GhIjKl');
    });

    it('strips query string before parsing', () => {
      const result = factory.create(makeInput({
        rowValues: { '.spItemUrl': 'https://graph.microsoft.com/v1.0/drives/DriveID/items/ItemID?$select=id' },
      }));

      expect(result.driveId).toBe('DriveID');
      expect(result.driveItemId).toBe('ItemID');
    });

    it('handles segment as the last path component', () => {
      const result = factory.create(makeInput({
        rowValues: { '.spItemUrl': 'https://graph.microsoft.com/v1.0/drives/OnlyDrive' },
      }));

      expect(result.driveId).toBe('OnlyDrive');
      expect(result.driveItemId).toBe('');
    });
  });
});
