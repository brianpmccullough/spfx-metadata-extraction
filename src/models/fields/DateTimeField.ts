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

  public generateLlmPrompt(): string {
    const required = this.isRequired ? ' (required)' : '';
    const desc = this.description ? ` - ${this.description}` : '';
    const format = this.includesTime ? 'date and time' : 'date only';
    return `"${this.title}"${required}: A ${format} field${desc}. Current value: ${this.formatForDisplay()}`;
  }

  public serializeForSharePoint(): string | null {
    return this.value?.toISOString() ?? null;
  }
}
