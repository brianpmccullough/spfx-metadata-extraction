import type { IDocumentContext } from '../models/IDocumentContext';

export function makeMockDocumentContext(overrides?: Partial<IDocumentContext>): IDocumentContext {
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
