import { FieldBase, FieldKind } from './FieldBase';

/**
 * Represents Text and Note field types from SharePoint.
 * Both store simple string values with no validation constraints.
 */
export class StringField extends FieldBase {
  public readonly fieldKind = FieldKind.String;

  constructor(
    id: string,
    internalName: string,
    title: string,
    description: string,
    isRequired: boolean,
    public value: string | null
  ) {
    super(id, internalName, title, description, isRequired);
  }

  public formatForDisplay(): string {
    return this.value ?? '(empty)';
  }

  public serializeForSharePoint(): string | null {
    return this.value;
  }
}
