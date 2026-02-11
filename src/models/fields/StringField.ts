/* eslint-disable @rushstack/no-new-null -- SharePoint REST API uses null for empty field values */
import { FieldBase, FieldKind } from './FieldBase';

/**
 * Represents Text and Note field types from SharePoint.
 * Text fields have a configurable MaxLength (up to 255).
 * Note fields are capped at 255 unless unlimited length is enabled.
 */
export class StringField extends FieldBase {
  public readonly fieldKind = FieldKind.String;

  constructor(
    id: string,
    internalName: string,
    title: string,
    description: string,
    isRequired: boolean,
    public value: string | null,
    public readonly maxLength: number | null = null
  ) {
    super(id, internalName, title, description, isRequired);
  }

  public formatForDisplay(): string {
    return this.value ?? '(empty)';
  }

  public serializeForSharePoint(): string | null {
    return this.value;
  }

  public isValidExtractedValue(value: string | number | boolean): boolean {
    if (this.maxLength !== null && String(value).length > this.maxLength) {
      return false;
    }
    return true;
  }

  public resolveValueForApply(value: string | number | boolean): string | number | boolean {
    return value;
  }
}
