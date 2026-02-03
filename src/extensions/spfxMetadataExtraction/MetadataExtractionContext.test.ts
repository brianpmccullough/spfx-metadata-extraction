import { MetadataExtractionContext } from './MetadataExtractionContext';
import type { ListViewCommandSetContext, RowAccessor } from '@microsoft/sp-listview-extensibility';

function makeRow(fileType: string): RowAccessor {
  return {
    getValueByName: (name: string) => {
      if (name === 'File_x0020_Type') return fileType;
      return '';
    },
  } as unknown as RowAccessor;
}

function makeContext(rows: RowAccessor[]): ListViewCommandSetContext {
  return {
    pageContext: {
      site: { absoluteUrl: 'https://contoso.sharepoint.com', id: { toString: () => 'site-id' } },
      web: { absoluteUrl: 'https://contoso.sharepoint.com/web', id: { toString: () => 'web-id' } },
    },
    listView: {
      selectedRows: rows,
      list: { guid: { toString: () => 'list-id' } },
    },
  } as unknown as ListViewCommandSetContext;
}

const defaultAllowed = ['.pdf', '.doc', '.docx'];

describe('MetadataExtractionContext', () => {
  describe('isValidFileType', () => {
    it('returns false for a pptx file', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('pptx')]),
        defaultAllowed
      );
      expect(ctx.isValidFileType).toBe(false);
    });

    it('returns false for a ppt file', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('ppt')]),
        defaultAllowed
      );
      expect(ctx.isValidFileType).toBe(false);
    });

    it('returns false for an xlsx file', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('xlsx')]),
        defaultAllowed
      );
      expect(ctx.isValidFileType).toBe(false);
    });

    it('returns false for an xls file', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('xls')]),
        defaultAllowed
      );
      expect(ctx.isValidFileType).toBe(false);
    });

    it('returns true for a pdf file', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('pdf')]),
        defaultAllowed
      );
      expect(ctx.isValidFileType).toBe(true);
    });

    it('returns true for a doc file', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('doc')]),
        defaultAllowed
      );
      expect(ctx.isValidFileType).toBe(true);
    });

    it('returns true for a docx file', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('docx')]),
        defaultAllowed
      );
      expect(ctx.isValidFileType).toBe(true);
    });
  });

  describe('isVisible', () => {
    it('returns false when a single pptx file is selected', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('pptx')]),
        defaultAllowed
      );
      expect(ctx.isVisible).toBe(false);
    });

    it('returns false when a single xlsx file is selected', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('xlsx')]),
        defaultAllowed
      );
      expect(ctx.isVisible).toBe(false);
    });

    it('returns true when a single pdf file is selected', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('pdf')]),
        defaultAllowed
      );
      expect(ctx.isVisible).toBe(true);
    });
  });

  describe('canExecute', () => {
    it('returns false for a pptx file', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('pptx')]),
        defaultAllowed
      );
      expect(ctx.canExecute).toBe(false);
    });

    it('returns false for an xlsx file', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('xlsx')]),
        defaultAllowed
      );
      expect(ctx.canExecute).toBe(false);
    });

    it('returns true for a pdf file', () => {
      const ctx = new MetadataExtractionContext(
        makeContext([makeRow('pdf')]),
        defaultAllowed
      );
      expect(ctx.canExecute).toBe(true);
    });
  });
});
