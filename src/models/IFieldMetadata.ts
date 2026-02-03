export type FieldType = 'string' | 'number' | 'boolean';

export interface IFieldMetadata {
  id: string;
  title: string;
  description: string;
  type: FieldType;
}
