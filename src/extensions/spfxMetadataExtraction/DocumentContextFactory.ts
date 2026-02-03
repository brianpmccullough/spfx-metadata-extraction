import { ListViewCommandSetContext, RowAccessor } from '@microsoft/sp-listview-extensibility';
import type { IDocumentContext } from '../../models/IDocumentContext';

const ROW_FIELDS: string[] = [
  'FileLeafRef',
  'File_x0020_Type',
  'FileRef',
  'ID',
  'UniqueId',
  'File_x0020_Size',
  'ContentTypeId',
  '.spItemUrl'
];

export interface IDocumentContextInput {
  rowValues: Record<string, string>;
  siteUrl: string;
  webUrl: string;
  siteId: string;
  webId: string;
  listId: string;
}

export class DocumentContextFactory {

  public createFromContext(context: ListViewCommandSetContext, row: RowAccessor): IDocumentContext {
    return this.create(this._extractInput(context, row));
  }

  public create(input: IDocumentContextInput): IDocumentContext {
    const val = (key: string): string => input.rowValues[key] || '';
    const fileSize = parseInt(val('File_x0020_Size') || '0', 10);

    return {
      fileName: val('FileLeafRef'),
      fileType: val('File_x0020_Type'),
      fileExtension: val('File_x0020_Type'),
      fileLeafRef: val('FileLeafRef'),
      fileRef: val('FileRef'),
      serverRelativeUrl: val('FileRef'),
      itemId: parseInt(val('ID') || '0', 10),
      uniqueId: val('UniqueId'),
      fileSize,
      fileSizeInBytes: fileSize,
      fileSizeInKiloBytes: fileSize / 1024,
      fileSizeInMegaBytes: fileSize / 1024 / 1024,
      siteUrl: input.siteUrl,
      webUrl: input.webUrl,
      siteId: input.siteId,
      webId: input.webId,
      listId: input.listId,
      contentTypeId: val('ContentTypeId'),
      spItemUrl: val('.spItemUrl'),
      driveId: this._getSpItemUrlSegmentValue('drives', val('.spItemUrl')),
      driveItemId: this._getSpItemUrlSegmentValue('items', val('.spItemUrl')),
    };
  }

  private _extractInput(context: ListViewCommandSetContext, row: RowAccessor): IDocumentContextInput {
    const pageContext = context.pageContext;
    const listView = context.listView;

    const rowValues: Record<string, string> = {};
    for (const field of ROW_FIELDS) {
      const value = row.getValueByName(field);
      rowValues[field] = value !== null && value !== undefined ? String(value) : '';
    }

    return {
      rowValues,
      siteUrl: pageContext.site.absoluteUrl,
      webUrl: pageContext.web.absoluteUrl,
      siteId: pageContext.site.id.toString(),
      webId: pageContext.web.id.toString(),
      listId: listView.list?.guid.toString() || '',
    };
  }

  private _getSpItemUrlSegmentValue(segment: string, spItemUrl: string): string {
    const urlWithoutQueryString = spItemUrl.split('?')[0];
    const lowerCaseUrl = urlWithoutQueryString.toLowerCase();
    const segmentPattern = `/${segment.toLowerCase()}/`;
    const segmentIndex = lowerCaseUrl.indexOf(segmentPattern);
    if (segmentIndex === -1) return '';

    const startIndex = segmentIndex + segmentPattern.length;
    const urlAfterSegment = urlWithoutQueryString.substring(startIndex);
    const nextSlashIndex = urlAfterSegment.indexOf('/');
    return nextSlashIndex === -1 ? urlAfterSegment : urlAfterSegment.substring(0, nextSlashIndex);
  }
}
