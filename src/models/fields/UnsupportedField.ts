import { FieldBase, FieldKind } from './FieldBase';

/**
 * Represents a SharePoint field type that is not currently supported.
 * Captures the original value and type for debugging purposes.
 */
export class UnsupportedField extends FieldBase {
  public readonly fieldKind = FieldKind.Unsupported;

  constructor(
    id: string,
    internalName: string,
    title: string,
    description: string,
    isRequired: boolean,
    public value: unknown,
    public readonly originalType: string
  ) {
    super(id, internalName, title, description, isRequired);
  }

  public formatForDisplay(): string {
    if (this.value === null || this.value === undefined) {
      return '(empty)';
    }
    return `[${this.originalType}]`;
  }

  public generateLlmPrompt(): string {
    return `"${this.title}": Unsupported field type (${this.originalType}). This field cannot be extracted.`;
  }

  public serializeForSharePoint(): null {
    // Unsupported fields should not be written back
    return null;
  }
}
