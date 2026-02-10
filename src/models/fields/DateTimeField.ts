import { FieldBase, FieldKind } from './FieldBase';

/**
 * Represents a DateTime field from SharePoint.
 * The includesTime property indicates whether the field should include time or just date.
 */
export class DateTimeField extends FieldBase {
  public readonly fieldKind = FieldKind.DateTime;

  constructor(
    id: string,
    internalName: string,
    title: string,
    description: string,
    isRequired: boolean,
    public value: Date | null,
    public readonly includesTime: boolean
  ) {
    super(id, internalName, title, description, isRequired);
  }

  public formatForDisplay(): string {
    if (!this.value) {
      return '(empty)';
    }
    return this.includesTime
      ? this.value.toLocaleString()
      : this.value.toLocaleDateString();
  }

  public serializeForSharePoint(): string | null {
    return this.value?.toISOString() ?? null;
  }
}
