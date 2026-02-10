import type { IDocumentContext } from '../models/IDocumentContext';
import type { ISharePointRestClient, ISharePointRestCollectionResponse } from '../clients';
import type { IMetadataExtractionService } from './IMetadataExtractionService';
import type { ITaxonomyService } from './ITaxonomyService';
import { TaxonomyService } from './TaxonomyService';
import { FieldBase, FieldFactory, ISharePointFieldSchema, type ITaxonomyValue } from '../models/fields';

/**
 * Extended SharePoint field response including properties needed for
 * Choice, MultiChoice, DateTime, and Taxonomy field types.
 */
interface ISharePointFieldResponse extends ISharePointFieldSchema {
  Hidden: boolean;
}

export class MetadataExtractionService implements IMetadataExtractionService {
  private readonly _fieldFactory: FieldFactory;

  private static readonly _excludedFieldInternalNames = [
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

  constructor(
    private readonly _spoClient: ISharePointRestClient,
    taxonomyService?: ITaxonomyService
  ) {
    const taxService = taxonomyService ?? new TaxonomyService(_spoClient);
    this._fieldFactory = new FieldFactory(taxService);
  }

  public async applyFieldValues(
    documentContext: IDocumentContext,
    fields: Array<{ internalName: string; value: string | number | boolean | null }>
  ): Promise<void> {
    const { webUrl, listId, itemId } = documentContext;
    const url = `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/ValidateUpdateListItem()`;

    const formValues = fields.map((f) => ({
      FieldName: f.internalName,
      FieldValue: String(f.value),
    }));

    await this._spoClient.post(url, { formValues });
  }

  public async loadFields(documentContext: IDocumentContext): Promise<FieldBase[]> {
    const schemas = await this.getFieldSchemas(documentContext);
    const values = await this.getFieldValues(documentContext, schemas);

    // Create field instances using the factory
    const fields = await Promise.all(
      schemas.map((schema) =>
        this._fieldFactory.createField(
          schema,
          values[schema.InternalName],
          documentContext.webUrl
        )
      )
    );

    return fields;
  }

  private async getFieldSchemas(
    documentContext: IDocumentContext
  ): Promise<ISharePointFieldSchema[]> {
    // Include additional properties needed for field type handling
    const selectFields = [
      'Id',
      'InternalName',
      'Title',
      'Description',
      'TypeAsString',
      'Required',
      'ReadOnlyField',
      'Choices',
      'DisplayFormat',
      'TermSetId',
    ].join(',');

    const url = `${documentContext.webUrl}/_api/web/lists(guid'${documentContext.listId}')/contenttypes('${documentContext.contentTypeId}')/fields?$filter=Hidden eq false&$select=${selectFields}`;
    const response = await this._spoClient.get<
      ISharePointRestCollectionResponse<ISharePointFieldResponse>
    >(url);

    return response.value.filter(
      (field) =>
        !MetadataExtractionService._excludedFieldInternalNames.includes(
          field.InternalName
        ) &&
        !field.ReadOnlyField &&
        !field.InternalName.startsWith('_')
    );
  }

  /**
   * Retrieves field values using RenderListDataAsStream API.
   * This API properly resolves taxonomy field values with labels and term GUIDs.
   */
  private async getFieldValues(
    documentContext: IDocumentContext,
    schemas: ISharePointFieldSchema[]
  ): Promise<Record<string, unknown>> {
    if (schemas.length === 0) {
      return {};
    }

    // Build ViewXml with fields and CAML query to filter to specific item
    const viewFields = schemas
      .map((s) => `<FieldRef Name="${s.InternalName}" />`)
      .join('');

    const viewXml = `<View Scope="RecursiveAll">
      <ViewFields>${viewFields}</ViewFields>
      <Query>
       <Where><Eq><FieldRef Name="ID" /><Value Type="Counter">${documentContext.itemId}</Value></Eq></Where>
      </Query>
      <RowLimit Paged="TRUE">1</RowLimit>
    </View>`;

    const url = `${documentContext.webUrl}/_api/web/lists(guid'${documentContext.listId}')/RenderListDataAsStream`;

    interface IRenderListDataResponse {
      Row: Array<Record<string, unknown>>;
    }

    const response = await this._spoClient.post<IRenderListDataResponse>(
      url,
      {
        parameters: {
          RenderOptions: 2,
          ViewXml: viewXml,
        },
      },
      {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata'
      }
    );

    if (!response.Row || response.Row.length === 0) {
      return {};
    }

    const row = response.Row[0];

    // Transform RenderListDataAsStream response to match expected format
    // Taxonomy fields come back with raw value in "FieldName." property
    return this.transformStreamResponse(row, schemas);
  }

  /**
   * Transforms RenderListDataAsStream response to standard field value format.
   * With odata=nometadata, taxonomy fields return as arrays of {Label, TermID} objects.
   */
  private transformStreamResponse(
    row: Record<string, unknown>,
    schemas: ISharePointFieldSchema[]
  ): Record<string, unknown> {
    const result: Record<string, unknown> = {};

    for (const schema of schemas) {
      const fieldName = schema.InternalName;

      if (schema.TypeAsString === 'TaxonomyFieldType') {
        // Single taxonomy: value is an object or array with single item
        result[fieldName] = this.parseTaxonomyValue(row[fieldName]);
      } else if (schema.TypeAsString === 'TaxonomyFieldTypeMulti') {
        // Multi taxonomy: value is an array of {Label, TermID} objects
        result[fieldName] = this.parseTaxonomyMultiValue(row[fieldName]);
      } else {
        // Non-taxonomy fields: use value directly
        result[fieldName] = row[fieldName];
      }
    }

    return result;
  }

  /**
   * Parses single taxonomy value from RenderListDataAsStream response.
   * Normalizes { Label, TermID } API format directly to ITaxonomyValue.
   */
  private parseTaxonomyValue(value: unknown): ITaxonomyValue | null {
    if (!value) {
      return null;
    }

    // Handle array format (take first item)
    if (Array.isArray(value)) {
      if (value.length === 0) {
        return null;
      }
      const item = value[0] as { Label?: string; TermID?: string };
      if (item.Label && item.TermID) {
        return { label: item.Label, termGuid: item.TermID };
      }
      return null;
    }

    // Handle object format
    const obj = value as { Label?: string; TermID?: string };
    if (obj.Label && obj.TermID) {
      return { label: obj.Label, termGuid: obj.TermID };
    }

    return null;
  }

  /**
   * Parses multi-value taxonomy from RenderListDataAsStream response.
   * Normalizes Array of { Label, TermID } API format directly to ITaxonomyValue[].
   */
  private parseTaxonomyMultiValue(value: unknown): ITaxonomyValue[] | null {
    if (!value || !Array.isArray(value) || value.length === 0) {
      return null;
    }

    const results: ITaxonomyValue[] = [];

    for (const item of value) {
      const parsed = this.parseTaxonomyValue(item);
      if (parsed) {
        results.push(parsed);
      }
    }

    return results.length > 0 ? results : null;
  }
}
