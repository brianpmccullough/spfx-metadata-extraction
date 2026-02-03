import type { IDocumentContext } from '../../models/IDocumentContext';
import type { IFieldMetadata } from '../../models/IFieldMetadata';

export interface IMetadataExtractionService {
  loadFieldMetadata(documentContext: IDocumentContext): Promise<IFieldMetadata[]>;
}
