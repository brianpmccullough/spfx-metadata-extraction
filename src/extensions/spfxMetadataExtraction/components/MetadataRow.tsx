import * as React from 'react';
import { Dropdown, IDropdownOption, Label, TextField } from '@fluentui/react';
import type { FieldType, IFieldMetadata } from '../../../models/IFieldMetadata';

export interface IMetadataRowProps {
  field: IFieldMetadata;
  onDescriptionChange: (id: string, description: string) => void;
  onTypeChange: (id: string, type: FieldType) => void;
}

const typeOptions: IDropdownOption[] = [
  { key: 'string', text: 'string' },
  { key: 'number', text: 'number' },
  { key: 'boolean', text: 'boolean' },
];

export const MetadataRow: React.FC<IMetadataRowProps> = ({ field, onDescriptionChange, onTypeChange }) => {
  const handleDescriptionChange = React.useCallback(
    (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
      onDescriptionChange(field.id, newValue ?? '');
    },
    [field.id, onDescriptionChange]
  );

  const handleTypeChange = React.useCallback(
    (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
      if (option) {
        onTypeChange(field.id, option.key as FieldType);
      }
    },
    [field.id, onTypeChange]
  );

  return (
    <div style={{ display: 'flex', alignItems: 'flex-end', gap: 12, marginBottom: 8 }}>
      <Label style={{ width: 160, flexShrink: 0 }}>{field.title}</Label>
      <TextField
        style={{ flex: 1 }}
        value={field.description}
        onChange={handleDescriptionChange}
        placeholder="Enter description"
      />
      <Dropdown
        style={{ width: 120 }}
        selectedKey={field.type}
        options={typeOptions}
        onChange={handleTypeChange}
      />
    </div>
  );
};
