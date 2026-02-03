export const FIELD_TYPES = ['string', 'number', 'boolean'] as const;

export type FieldType = typeof FIELD_TYPES[number];

export type FieldValue = string | number | boolean | null;

export interface IFieldMetadata {
  id: string;
  internalName: string;
  title: string;
  description: string;
  type: FieldType;
  value: FieldValue;
}
