import { FieldBase, FieldKind } from './FieldBase';

/**
 * Represents a single-select Choice field from SharePoint.
 */
export class ChoiceField extends FieldBase {
  public readonly fieldKind = FieldKind.Choice;

  constructor(
    id: string,
    internalName: string,
    title: string,
    description: string,
    isRequired: boolean,
    public value: string | null,
    public readonly choices: string[]
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

/**
 * Represents a multi-select Choice field from SharePoint.
 */
export class MultiChoiceField extends FieldBase {
  public readonly fieldKind = FieldKind.MultiChoice;

  constructor(
    id: string,
    internalName: string,
    title: string,
    description: string,
    isRequired: boolean,
    public value: string[] | null,
    public readonly choices: string[]
  ) {
    super(id, internalName, title, description, isRequired);
  }

  public formatForDisplay(): string {
    if (!this.value || this.value.length === 0) {
      return '(empty)';
    }
    return this.value.join(', ');
  }

  public serializeForSharePoint(): string | null {
    if (!this.value || this.value.length === 0) {
      return null;
    }
    // SharePoint expects ";#value1;#value2;#" format for multi-choice
    return `;#${this.value.join(';#')};#`;
  }
}
