import type { IDocumentContext } from '../../models/IDocumentContext';
import type { IFieldInfo } from '../../models/IFieldInfo';
import type { FieldType, FieldValue, IFieldMetadata } from '../../models/IFieldMetadata';
import type { ISharePointRestClient } from '../../clients/ISharePointRestClient';
import type { IGraphClient } from '../../clients/IGraphClient';
import type { IMetadataExtractionService } from './IMetadataExtractionService';
import { ISharePointRestCollectionResponse } from '../../clients/ISharePointRestCollectionResponse';

interface ISharePointFieldResponse {
  Id: string;
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Required: boolean;
  ReadOnlyField: boolean;
  Description: string;
}

export class MetadataExtractionService implements IMetadataExtractionService {
  constructor(
    private readonly _spoClient: ISharePointRestClient,
    private readonly _graphClient: IGraphClient
  ) {}

  private static readonly _typeMap: Record<string, FieldType> = {
    'Text': 'string',
    'Note': 'string',
    'Choice': 'string',
    'MultiChoice': 'string',
    'URL': 'string',
    'Number': 'number',
    'Currency': 'number',
    'Boolean': 'boolean',
  };

  public async loadFieldMetadata(documentContext: IDocumentContext): Promise<IFieldMetadata[]> {
    const fields = await this.getContentTypeFields(documentContext);
    const fieldValues = await this.getItemFieldValues(documentContext, fields);

    return fields.map((field) => ({
      id: field.id,
      internalName: field.internalName,
      title: field.title,
      description: field.description,
      type: MetadataExtractionService._typeMap[field.typeAsString] ?? 'string',
      value: this.normalizeFieldValue(fieldValues[field.internalName]),
    }));
  }

  private async getItemFieldValues(
    documentContext: IDocumentContext,
    fields: IFieldInfo[]
  ): Promise<Record<string, FieldValue>> {
    if (fields.length === 0) {
      return {};
    }

    const selectFields = fields.map((f) => f.internalName).join(',');
    const url = `${documentContext.webUrl}/_api/web/lists(guid'${documentContext.listId}')/items(${documentContext.itemId})?$select=${selectFields}`;
    const response = await this._spoClient.get<Record<string, FieldValue>>(url);
    return response;
  }

  private normalizeFieldValue(value: unknown): FieldValue {
    if (value === undefined || value === null) {
      return null;
    }
    if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
      return value;
    }
    // Handle complex field types (e.g., lookup, user) by converting to string
    return String(value);
  }

  public async getContentTypeFields(documentContext: IDocumentContext): Promise<IFieldInfo[]> {
    const excludedFieldInternalNames = [
      'ContentType',
      'Title',
      'LinkFilename',
      'FileLeafRef',
      'Modified',
      'Created',
      'Author',
      'Created_x0020_By',
      'Editor',
      'Modified_x0020_By',
      'RatingCount',
      'AverageRating',
      'LikesCount',
    ];

    const url = `${documentContext.webUrl}/_api/web/lists(guid'${documentContext.listId}')/contenttypes('${documentContext.contentTypeId}')/fields?$filter=Hidden eq false`;
    const response = await this._spoClient.get<ISharePointRestCollectionResponse<ISharePointFieldResponse>>(url);
    return response.value
      .filter((field) => !excludedFieldInternalNames.includes(field.InternalName) 
                          && !field.ReadOnlyField 
                          && !field.InternalName.startsWith('_'))
      .map((field) => ({
      id: field.Id,
      internalName: field.InternalName,
      title: field.Title,
      typeAsString: field.TypeAsString,
      required: field.Required,
      readOnly: field.ReadOnlyField,
      description: field.Description,
    }));
  }
}
