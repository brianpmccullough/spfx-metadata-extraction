import type { FieldBase } from '../fields';
import { FieldKind } from '../fields';

/**
 * Simple field types for LLM metadata extraction.
 * These map to how the LLM should interpret and extract the value.
 */
export enum MetadataExtractionFieldType {
  String = 'string',
  Number = 'number',
  Boolean = 'boolean',
}

/**
 * Wraps a SharePoint field with LLM extraction-specific configuration.
 * Provides editable extraction type and description for LLM guidance.
 */
export class MetadataExtractionField {
  public extractionType: MetadataExtractionFieldType;
  public description: string;

  constructor(public readonly field: FieldBase) {
    this.extractionType = this.inferExtractionType();
    this.description = field.description;
  }

  /**
   * Infers the LLM extraction type based on the SharePoint field kind.
   * - Numeric fields → Number
   * - Boolean fields → Boolean
   * - Everything else → String
   */
  private inferExtractionType(): MetadataExtractionFieldType {
    switch (this.field.fieldKind) {
      case FieldKind.Numeric:
        return MetadataExtractionFieldType.Number;
      case FieldKind.Boolean:
        return MetadataExtractionFieldType.Boolean;
      default:
        return MetadataExtractionFieldType.String;
    }
  }
}
