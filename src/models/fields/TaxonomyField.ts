/* eslint-disable @rushstack/no-new-null -- SharePoint REST API uses null for empty field values */
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
    public readonly terms: ITerm[]
  ) {
    super(id, internalName, title, description, isRequired);
  }

  public formatForDisplay(): string {
    return this.value?.label ?? '(empty)';
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

  public isValidExtractedValue(value: string | number | boolean): boolean {
    const label = String(value).trim();
    return this.terms.some((t) => t.label.toLowerCase() === label.toLowerCase());
  }

  public resolveValueForApply(value: string | number | boolean): string {
    const label = String(value).trim();
    const match = this.terms.find(
      (t) => t.label.toLowerCase() === label.toLowerCase()
    );
    if (match) {
      return `${match.label}|${match.termGuid}`;
    }
    return label;
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

  public isValidExtractedValue(value: string | number | boolean): boolean {
    const labels = String(value)
      .split(',')
      .map((l) => l.trim())
      .filter((l) => l.length > 0);
    return labels.every((label) =>
      this.terms.some((t) => t.label.toLowerCase() === label.toLowerCase())
    );
  }

  public resolveValueForApply(value: string | number | boolean): string {
    const labels = String(value)
      .split(',')
      .map((l) => l.trim())
      .filter((l) => l.length > 0);
    const resolved = labels.map((label) => {
      const match = this.terms.find(
        (t) => t.label.toLowerCase() === label.toLowerCase()
      );
      if (match) {
        return `${match.label}|${match.termGuid}`;
      }
      return label;
    });
    return resolved.join(';#');
  }
}
