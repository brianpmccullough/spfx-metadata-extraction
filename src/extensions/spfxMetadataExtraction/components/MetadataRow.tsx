import * as React from 'react';
import { Checkbox, Dropdown, Label, Text, TextField, IDropdownOption } from '@fluentui/react';
import {
  MetadataExtractionField,
  MetadataExtractionFieldType,
} from '../../../models/extraction';

export interface IMetadataRowProps {
  extractionField: MetadataExtractionField;
  onExtractionTypeChange: (newType: MetadataExtractionFieldType) => void;
  onDescriptionChange: (newDescription: string) => void;
  applyChecked: boolean;
  onApplyCheckedChange: (checked: boolean) => void;
  isApplyEnabled: boolean;
}

const confidenceStyles: Record<string, { background: string; text: string }> = {
  green: { background: '#e6f4ea', text: '#107c10' },
  yellow: { background: '#fff4ce', text: '#8a6914' },
  red: { background: '#fde7e9', text: '#a4262c' },
  none: { background: '#f3f2f1', text: '#888' },
};

const extractionTypeOptions: IDropdownOption[] = [
  { key: MetadataExtractionFieldType.String, text: MetadataExtractionFieldType.String },
  { key: MetadataExtractionFieldType.Number, text: MetadataExtractionFieldType.Number },
  { key: MetadataExtractionFieldType.Boolean, text: MetadataExtractionFieldType.Boolean },
];

export const MetadataRow: React.FC<IMetadataRowProps> = ({
  extractionField,
  onExtractionTypeChange,
  onDescriptionChange,
  applyChecked,
  onApplyCheckedChange,
  isApplyEnabled,
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

  const handleApplyCheckedChange = React.useCallback(
    (_ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
      onApplyCheckedChange(!!checked);
    },
    [onApplyCheckedChange]
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
        styles={{ root: { flex: 1, minWidth: 150 } }}
        placeholder="Description for LLM"
      />
      <Text
        style={{
          width: 120,
          flexShrink: 0,
          padding: '6px 0',
          color: field.value === null ? '#888' : undefined,
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap',
        }}
        title={field.formatForDisplay()}
      >
        {field.formatForDisplay()}
      </Text>
      {/* Extracted Value - read only, styled by confidence */}
      <Text
        style={{
          width: 150,
          flexShrink: 0,
          padding: '6px 8px',
          backgroundColor: confidenceStyles[extractionField.confidence ?? 'none'].background,
          borderRadius: 4,
          color: confidenceStyles[extractionField.confidence ?? 'none'].text,
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap',
          fontWeight: extractionField.extractedValue !== null ? 500 : 400,
        }}
        title={extractionField.extractedValue !== null ? String(extractionField.extractedValue) : '(not extracted)'}
      >
        {extractionField.extractedValue !== null ? String(extractionField.extractedValue) : '(not extracted)'}
      </Text>
      <div style={{ width: 60, flexShrink: 0, display: 'flex', justifyContent: 'center', paddingTop: 6 }}>
        <Checkbox
          checked={applyChecked}
          onChange={handleApplyCheckedChange}
          disabled={!isApplyEnabled}
        />
      </div>
    </div>
  );
};
