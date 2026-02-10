import type { FieldBase } from '../fields';
import type { IDocumentContext } from '../IDocumentContext';
import {
  FieldKind,
  ChoiceField,
  MultiChoiceField,
  TaxonomyField,
  TaxonomyMultiField,
  DateTimeField,
} from '../fields';
/**
 * Confidence level for an extraction result.
 */
export type ExtractionConfidence = 'red' | 'yellow' | 'green';

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
  public extractedValue: string | number | boolean | null = null;
  public confidence: ExtractionConfidence | null = null;

  constructor(public readonly field: FieldBase) {
    this.extractionType = this.inferExtractionType();
    this.description = this.buildDescription();
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

  /**
   * Builds a description combining the field's SharePoint description
   * with type-specific extraction hints for the LLM.
   */
  private buildDescription(): string {
    const parts: string[] = [];

    if (this.field.description) {
      parts.push(this.field.description);
    }

    const hint = this.getExtractionHint();
    if (hint) {
      parts.push(hint);
    }

    return parts.join('. ');
  }

  /**
   * Returns type-specific extraction guidance for the LLM.
   */
  private getExtractionHint(): string {
    const { field } = this;

    switch (field.fieldKind) {
      case FieldKind.DateTime: {
        const dtField = field as DateTimeField;
        return dtField.includesTime
          ? 'Return as ISO 8601 date and time (e.g. 2025-01-15T14:30:00Z)'
          : 'Return as ISO 8601 date (e.g. 2025-01-15)';
      }
      case FieldKind.Choice: {
        const choiceField = field as ChoiceField;
        const choices = choiceField.choices.map((c) => `"${c}"`).join(', ');
        return `Value must be one of: [${choices}]`;
      }
      case FieldKind.MultiChoice: {
        const mcField = field as MultiChoiceField;
        const choices = mcField.choices.map((c) => `"${c}"`).join(', ');
        return `Select one or more from: [${choices}]`;
      }
      case FieldKind.Taxonomy: {
        const taxField = field as TaxonomyField;
        const terms = taxField.terms.map((t) => `"${t.label}"`).join(', ');
        return `Value must be one of: [${terms}]`;
      }
      case FieldKind.TaxonomyMulti: {
        const taxMultiField = field as TaxonomyMultiField;
        const terms = taxMultiField.terms.map((t) => `"${t.label}"`).join(', ');
        return `Select one or more from: [${terms}]`;
      }
      case FieldKind.Boolean:
        return 'Return as true or false';
      case FieldKind.Numeric:
        return 'Return as a number';
      default:
        return '';
    }
  }

  /**
   * Creates a shallow clone with all mutable state copied.
   * The underlying field reference is shared (it's immutable).
   */
  public clone(): MetadataExtractionField {
    const cloned = new MetadataExtractionField(this.field);
    cloned.extractionType = this.extractionType;
    cloned.description = this.description;
    cloned.extractedValue = this.extractedValue;
    cloned.confidence = this.confidence;
    return cloned;
  }

  /**
   * Returns schema representation for LLM extraction API.
   */
  public toSchema(): IFieldSchema {
    return {
      internalName: this.field.internalName,
      title: this.field.title,
      description: this.description,
      dataType: this.extractionType,
    };
  }

  /**
   * Builds an extraction request from document context and fields.
   */
  public static buildExtractionRequest(
    documentContext: IDocumentContext,
    fields: MetadataExtractionField[]
  ): { document: { driveId: string; driveItemId: string }; fields: IFieldSchema[] } {
    return {
      document: {
        driveId: documentContext.driveId,
        driveItemId: documentContext.driveItemId,
      },
      fields: fields.map((f) => f.toSchema()),
    };
  }
}

/**
 * Schema representation of a field for LLM extraction.
 */
export interface IFieldSchema {
  /** Internal name used for mapping results back to SharePoint fields */
  internalName: string;
  /** The name/title of the field to extract */
  title: string;
  /** Instructions for the LLM on how to extract content for this field */
  description: string;
  /** The expected data type of the extracted value */
  dataType: MetadataExtractionFieldType;
}
