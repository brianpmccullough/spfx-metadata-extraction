/* eslint-disable @rushstack/no-new-null -- SharePoint REST API uses null for empty field values */
import { FieldBase, FieldKind } from './FieldBase';

/**
 * Represents Number and Currency field types from SharePoint.
 * Both store numeric values.
 */
export class NumericField extends FieldBase {
  public readonly fieldKind = FieldKind.Numeric;

  constructor(
    id: string,
    internalName: string,
    title: string,
    description: string,
    isRequired: boolean,
    public value: number | null
  ) {
    super(id, internalName, title, description, isRequired);
  }

  public formatForDisplay(): string {
    if (this.value === null) {
      return '(empty)';
    }
    return String(this.value);
  }

  public serializeForSharePoint(): number | null {
    return this.value;
  }

  public isValidExtractedValue(_value: string | number | boolean): boolean {
    return true;
  }

  public resolveValueForApply(value: string | number | boolean): string | number | boolean {
    return value;
  }
}
