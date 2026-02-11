/**
 * Discriminator for field type narrowing in consuming code.
 */
export enum FieldKind {
  String = 'String',
  Choice = 'Choice',
  MultiChoice = 'MultiChoice',
  Taxonomy = 'Taxonomy',
  TaxonomyMulti = 'TaxonomyMulti',
  Numeric = 'Numeric',
  Boolean = 'Boolean',
  DateTime = 'DateTime',
  Unsupported = 'Unsupported',
}

/**
 * Abstract base class for all field types.
 * Provides common properties and defines the contract for type-specific behavior.
 */
export abstract class FieldBase {
  constructor(
    public readonly id: string,
    public readonly internalName: string,
    public readonly title: string,
    public readonly description: string,
    public readonly isRequired: boolean
  ) {}

  /** Current value of the field. Type varies by subclass. */
  public abstract value: unknown;

  /** Discriminator for type narrowing. */
  public abstract readonly fieldKind: FieldKind;

  /** Format the current value for display in UI. */
  public abstract formatForDisplay(): string;

  /** Serialize the value for SharePoint REST API write-back. */
  public abstract serializeForSharePoint(): unknown;

  /** Whether an LLM-extracted value is valid for this field's constraints. */
  public abstract isValidExtractedValue(value: string | number | boolean): boolean;

  /** Format an LLM-extracted value for ValidateUpdateListItem write-back. */
  public abstract resolveValueForApply(value: string | number | boolean): string | number | boolean;
}
