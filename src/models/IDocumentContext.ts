export interface IDocumentContext {

  contentTypeId: string;

  /** File extension without leading dot (e.g., "docx") */
  fileExtension: string;
  /** File name including extension (e.g., "report.docx") */
  fileName: string;

  fileLeafRef: string;

  fileRef: string;

  fileSize: number;

  fileSizeInBytes: number;

  fileSizeInKiloBytes: number;

  fileSizeInMegaBytes: number;

  fileType: string;


  /** Server-relative path to the file (e.g., "/sites/MySite/Documents/report.docx") */
  serverRelativeUrl: string;

  /** SharePoint list item integer ID */
  itemId: number;
  /** Item unique identifier (GUID) */
  uniqueId: string;

  driveId: string;

  driveItemId: string;

  spItemUrl: string;

  /** Site collection absolute URL */
  siteUrl: string;
  /** Web absolute URL */
  webUrl: string;
  /** Site collection GUID */
  siteId: string;
  /** Web GUID */
  webId: string;
  /** List GUID */
  listId: string;
}