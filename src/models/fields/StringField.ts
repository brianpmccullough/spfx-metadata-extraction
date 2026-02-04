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

  public generateLlmPrompt(): string {
    const required = this.isRequired ? ' (required)' : '';
    const desc = this.description ? ` - ${this.description}` : '';
    return `"${this.title}"${required}: A text field${desc}. Current value: ${this.formatForDisplay()}`;
  }

  public serializeForSharePoint(): string | null {
    return this.value;
  }
}
