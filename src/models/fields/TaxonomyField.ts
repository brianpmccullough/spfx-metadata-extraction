import { FieldBase, FieldKind } from './FieldBase';

/**
 * Represents a taxonomy term value with GUID for write-back support.
 */
export interface ITaxonomyValue {
  termGuid: string;
  label: string;
  wssId?: number;
}

/**
 * Represents an available term in a term set.
 */
export interface ITerm {
  termGuid: string;
  label: string;
}

/**
 * Represents a single-select managed metadata (Taxonomy) field from SharePoint.
 */
export class TaxonomyField extends FieldBase {
  public readonly fieldKind = FieldKind.Taxonomy;

  constructor(
    id: string,
    internalName: string,
    title: string,
    description: string,
    isRequired: boolean,
    public value: ITaxonomyValue | null,
    public readonly termSetId: string,
    public readonly sspId: string,
    public readonly terms: ITerm[]
  ) {
    super(id, internalName, title, description, isRequired);
  }

  public formatForDisplay(): string {
    return this.value?.label ?? '(empty)';
  }

  public generateLlmPrompt(): string {
    const required = this.isRequired ? ' (required)' : '';
    const desc = this.description ? ` - ${this.description}` : '';
    const validTerms = `Valid terms: [${this.terms.map((t) => `"${t.label}"`).join(', ')}]`;
    return `"${this.title}"${required}: A managed metadata field${desc}. ${validTerms}. Current value: ${this.formatForDisplay()}`;
  }

  public serializeForSharePoint(): object | null {
    if (!this.value) {
      return null;
    }
    return {
      __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
      Label: this.value.label,
      TermGuid: this.value.termGuid,
      WssId: this.value.wssId ?? -1,
    };
  }
}

/**
 * Represents a multi-select managed metadata (Taxonomy) field from SharePoint.
 */
export class TaxonomyMultiField extends FieldBase {
  public readonly fieldKind = FieldKind.TaxonomyMulti;

  constructor(
    id: string,
    internalName: string,
    title: string,
    description: string,
    isRequired: boolean,
    public value: ITaxonomyValue[] | null,
    public readonly termSetId: string,
    public readonly sspId: string,
    public readonly terms: ITerm[]
  ) {
    super(id, internalName, title, description, isRequired);
  }

  public formatForDisplay(): string {
    if (!this.value || this.value.length === 0) {
      return '(empty)';
    }
    return this.value.map((v) => v.label).join(', ');
  }

  public generateLlmPrompt(): string {
    const required = this.isRequired ? ' (required)' : '';
    const desc = this.description ? ` - ${this.description}` : '';
    const validTerms = `Valid terms: [${this.terms.map((t) => `"${t.label}"`).join(', ')}]`;
    return `"${this.title}"${required}: A multi-value managed metadata field (select one or more)${desc}. ${validTerms}. Current value: ${this.formatForDisplay()}`;
  }

  public serializeForSharePoint(): object[] | null {
    if (!this.value || this.value.length === 0) {
      return null;
    }
    return this.value.map((v) => ({
      __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
      Label: v.label,
      TermGuid: v.termGuid,
      WssId: v.wssId ?? -1,
    }));
  }
}
