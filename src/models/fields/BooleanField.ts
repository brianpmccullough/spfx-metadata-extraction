import { FieldBase, FieldKind } from './FieldBase';

/**
 * Represents a Boolean (Yes/No) field from SharePoint.
 */
export class BooleanField extends FieldBase {
  public readonly fieldKind = FieldKind.Boolean;

  constructor(
    id: string,
    internalName: string,
    title: string,
    description: string,
    isRequired: boolean,
    public value: boolean | null
  ) {
    super(id, internalName, title, description, isRequired);
  }

  public formatForDisplay(): string {
    if (this.value === null) {
      return '(empty)';
    }
    return this.value ? 'Yes' : 'No';
  }

  public serializeForSharePoint(): boolean | null {
    return this.value;
  }
}
