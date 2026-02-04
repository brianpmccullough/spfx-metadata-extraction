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

  public generateLlmPrompt(): string {
    const required = this.isRequired ? ' (required)' : '';
    const desc = this.description ? ` - ${this.description}` : '';
    return `"${this.title}"${required}: A yes/no field${desc}. Current value: ${this.formatForDisplay()}`;
  }

  public serializeForSharePoint(): boolean | null {
    return this.value;
  }
}
