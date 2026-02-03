import * as React from 'react';
import { Dropdown, IDropdownOption, Label, Text, TextField } from '@fluentui/react';
import { FIELD_TYPES, type FieldType, type IFieldMetadata } from '../../../models/IFieldMetadata';

export interface IMetadataRowProps {
  field: IFieldMetadata;
  onDescriptionChange: (id: string, description: string) => void;
  onTypeChange: (id: string, type: FieldType) => void;
}

const typeOptions: IDropdownOption[] = FIELD_TYPES.map((t) => ({ key: t, text: t }));

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

  const formatValue = (value: string | number | boolean | null): string => {
    if (value === null || value === undefined) {
      return '(empty)';
    }
    if (typeof value === 'boolean') {
      return value ? 'Yes' : 'No';
    }
    return String(value);
  };

  return (
    <div style={{ display: 'flex', alignItems: 'flex-start', gap: 12, marginBottom: 8 }}>
      <Label style={{ width: 140, flexShrink: 0, paddingTop: 6 }}>{field.title}</Label>
      <Dropdown
        style={{ width: 80 }}
        selectedKey={field.type}
        options={typeOptions}
        onChange={handleTypeChange}
      />
      <TextField
        style={{ flex: 1, minWidth: 200 }}
        value={field.description}
        onChange={handleDescriptionChange}
        placeholder="Enter description"
        multiline
        rows={2}
        resizable={false}
      />
      <Text
        style={{
          width: 180,
          flexShrink: 0,
          padding: '6px 0',
          color: field.value === null ? '#888' : undefined,
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap',
        }}
      >
        {formatValue(field.value)}
      </Text>
    </div>
  );
};
