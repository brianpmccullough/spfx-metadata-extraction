import * as React from 'react';
import { Dropdown, Label, Text, TextField, IDropdownOption } from '@fluentui/react';
import {
  MetadataExtractionField,
  MetadataExtractionFieldType,
} from '../../../models/extraction';

export interface IMetadataRowProps {
  extractionField: MetadataExtractionField;
  onExtractionTypeChange: (newType: MetadataExtractionFieldType) => void;
  onDescriptionChange: (newDescription: string) => void;
}

const extractionTypeOptions: IDropdownOption[] = [
  { key: MetadataExtractionFieldType.String, text: MetadataExtractionFieldType.String },
  { key: MetadataExtractionFieldType.Number, text: MetadataExtractionFieldType.Number },
  { key: MetadataExtractionFieldType.Boolean, text: MetadataExtractionFieldType.Boolean },
];

export const MetadataRow: React.FC<IMetadataRowProps> = ({
  extractionField,
  onExtractionTypeChange,
  onDescriptionChange,
}) => {
  const { field } = extractionField;

  const handleTypeChange = React.useCallback(
    (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
      if (option) {
        onExtractionTypeChange(option.key as MetadataExtractionFieldType);
      }
    },
    [onExtractionTypeChange]
  );

  const handleDescriptionChange = React.useCallback(
    (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
      onDescriptionChange(newValue ?? '');
    },
    [onDescriptionChange]
  );

  return (
    <div style={{ display: 'flex', alignItems: 'flex-start', gap: 12, marginBottom: 12, paddingBottom: 12, borderBottom: '1px solid #edebe9' }}>
      <Label style={{ width: 120, flexShrink: 0, paddingTop: 6 }}>{field.title}</Label>
      <Dropdown
        selectedKey={extractionField.extractionType}
        options={extractionTypeOptions}
        onChange={handleTypeChange}
        styles={{ root: { width: 100, flexShrink: 0 } }}
      />
      <TextField
        value={extractionField.description}
        onChange={handleDescriptionChange}
        multiline
        rows={2}
        styles={{ root: { flex: 1, minWidth: 200 } }}
        placeholder="Description for LLM"
      />
      <Text
        style={{
          width: 150,
          flexShrink: 0,
          padding: '6px 0',
          color: field.value === null ? '#888' : undefined,
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap',
        }}
      >
        {field.formatForDisplay()}
      </Text>
    </div>
  );
};
