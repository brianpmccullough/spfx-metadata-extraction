import type { ITaxonomyService } from '../../services/ITaxonomyService';
import { FieldBase } from './FieldBase';
import { StringField } from './StringField';
import { ChoiceField, MultiChoiceField } from './ChoiceField';
import { TaxonomyField, TaxonomyMultiField, ITaxonomyValue } from './TaxonomyField';
import { NumericField } from './NumericField';
import { BooleanField } from './BooleanField';
import { DateTimeField } from './DateTimeField';
import { UnsupportedField } from './UnsupportedField';

/**
 * SharePoint field schema returned by the REST API.
 */
export interface ISharePointFieldSchema {
  Id: string;
  InternalName: string;
  Title: string;
  Description: string;
  TypeAsString: string;
  Required: boolean;
  ReadOnlyField: boolean;
  // Choice fields (OData verbose: { results: string[] }, nometadata: string[])
  Choices?: { results: string[] } | string[];
  // DateTime fields (0 = DateOnly, 1 = DateTime)
  DisplayFormat?: number;
  // Taxonomy fields
  TermSetId?: string;
  SspId?: string;
}

/**
 * Factory interface for creating field instances.
 */
export interface IFieldFactory {
  createField(
    schema: ISharePointFieldSchema,
    value: unknown,
    siteUrl: string
  ): Promise<FieldBase>;
}

/**
 * Factory for creating field instances from SharePoint field schema.
 */
export class FieldFactory implements IFieldFactory {
  constructor(private readonly _taxonomyService: ITaxonomyService) {}

  public async createField(
    schema: ISharePointFieldSchema,
    value: unknown,
    siteUrl: string
  ): Promise<FieldBase> {
    const baseArgs = {
      id: schema.Id,
      internalName: schema.InternalName,
      title: schema.Title,
      description: schema.Description,
      isRequired: schema.Required,
    };

    switch (schema.TypeAsString) {
      case 'Text':
      case 'Note':
        return new StringField(
          baseArgs.id,
          baseArgs.internalName,
          baseArgs.title,
          baseArgs.description,
          baseArgs.isRequired,
          value as string | null
        );

      case 'Choice':
        return new ChoiceField(
          baseArgs.id,
          baseArgs.internalName,
          baseArgs.title,
          baseArgs.description,
          baseArgs.isRequired,
          value as string | null,
          this.parseChoices(schema.Choices)
        );

      case 'MultiChoice':
        return new MultiChoiceField(
          baseArgs.id,
          baseArgs.internalName,
          baseArgs.title,
          baseArgs.description,
          baseArgs.isRequired,
          this.parseMultiChoiceValue(value),
          this.parseChoices(schema.Choices)
        );

      case 'TaxonomyFieldType':
        return this.createTaxonomyField(schema, baseArgs, value, siteUrl);

      case 'TaxonomyFieldTypeMulti':
        return this.createTaxonomyMultiField(schema, baseArgs, value, siteUrl);

      case 'Number':
      case 'Currency':
        return new NumericField(
          baseArgs.id,
          baseArgs.internalName,
          baseArgs.title,
          baseArgs.description,
          baseArgs.isRequired,
          value as number | null
        );

      case 'Boolean':
        return new BooleanField(
          baseArgs.id,
          baseArgs.internalName,
          baseArgs.title,
          baseArgs.description,
          baseArgs.isRequired,
          value as boolean | null
        );

      case 'DateTime':
        return new DateTimeField(
          baseArgs.id,
          baseArgs.internalName,
          baseArgs.title,
          baseArgs.description,
          baseArgs.isRequired,
          value ? new Date(value as string) : null,
          schema.DisplayFormat === 1
        );

      default:
        return new UnsupportedField(
          baseArgs.id,
          baseArgs.internalName,
          baseArgs.title,
          baseArgs.description,
          baseArgs.isRequired,
          value,
          schema.TypeAsString
        );
    }
  }

  private async createTaxonomyField(
    schema: ISharePointFieldSchema,
    baseArgs: {
      id: string;
      internalName: string;
      title: string;
      description: string;
      isRequired: boolean;
    },
    value: unknown,
    siteUrl: string
  ): Promise<TaxonomyField> {
    const termSetId = schema.TermSetId ?? '';
    const sspId = schema.SspId ?? '';
    const terms =
      termSetId && sspId
        ? await this._taxonomyService.getTerms(termSetId, sspId, siteUrl)
        : [];

    return new TaxonomyField(
      baseArgs.id,
      baseArgs.internalName,
      baseArgs.title,
      baseArgs.description,
      baseArgs.isRequired,
      this.parseTaxonomyValue(value),
      termSetId,
      sspId,
      terms
    );
  }

  private async createTaxonomyMultiField(
    schema: ISharePointFieldSchema,
    baseArgs: {
      id: string;
      internalName: string;
      title: string;
      description: string;
      isRequired: boolean;
    },
    value: unknown,
    siteUrl: string
  ): Promise<TaxonomyMultiField> {
    const termSetId = schema.TermSetId ?? '';
    const sspId = schema.SspId ?? '';
    const terms =
      termSetId && sspId
        ? await this._taxonomyService.getTerms(termSetId, sspId, siteUrl)
        : [];

    return new TaxonomyMultiField(
      baseArgs.id,
      baseArgs.internalName,
      baseArgs.title,
      baseArgs.description,
      baseArgs.isRequired,
      this.parseTaxonomyMultiValue(value),
      termSetId,
      sspId,
      terms
    );
  }

  private parseChoices(choices: ISharePointFieldSchema['Choices']): string[] {
    if (!choices) {
      return [];
    }
    // OData nometadata returns string[], verbose returns { results: string[] }
    if (Array.isArray(choices)) {
      return choices;
    }
    return choices.results ?? [];
  }

  private parseMultiChoiceValue(value: unknown): string[] | null {
    if (value === null || value === undefined) {
      return null;
    }
    if (Array.isArray(value)) {
      return value as string[];
    }
    // SharePoint REST API returns ";#value1;#value2;#" format
    if (typeof value === 'string') {
      const parts = value.split(';#').filter((v) => v.length > 0);
      return parts.length > 0 ? parts : null;
    }
    return null;
  }

  private parseTaxonomyValue(value: unknown): ITaxonomyValue | null {
    if (value === null || value === undefined) {
      return null;
    }
    const v = value as { TermGuid?: string; Label?: string; WssId?: number };
    if (v.TermGuid && v.Label) {
      return { termGuid: v.TermGuid, label: v.Label, wssId: v.WssId };
    }
    return null;
  }

  private parseTaxonomyMultiValue(value: unknown): ITaxonomyValue[] | null {
    if (value === null || value === undefined) {
      return null;
    }
    if (Array.isArray(value)) {
      const parsed = value
        .map((v) => this.parseTaxonomyValue(v))
        .filter((v): v is ITaxonomyValue => v !== null);
      return parsed.length > 0 ? parsed : null;
    }
    return null;
  }
}
