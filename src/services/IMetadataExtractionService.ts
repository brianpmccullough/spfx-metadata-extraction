import type { IDocumentContext } from '../models/IDocumentContext';
import type { FieldBase } from '../models/fields';

export interface IMetadataExtractionService {
  loadFields(documentContext: IDocumentContext): Promise<FieldBase[]>;
  applyFieldValues(
    documentContext: IDocumentContext,
    fields: Array<{ internalName: string; value: string | number | boolean | null }>
  ): Promise<void>;
}
