import type { IDocumentContext } from '../../models/IDocumentContext';
import type { FieldBase } from '../../models/fields';

export interface IMetadataExtractionService {
  loadFields(documentContext: IDocumentContext): Promise<FieldBase[]>;
}
