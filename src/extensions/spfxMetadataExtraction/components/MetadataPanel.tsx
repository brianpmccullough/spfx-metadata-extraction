import * as React from 'react';
import {
  DefaultButton,
  IconButton,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack,
} from '@fluentui/react';
import type { FieldType, IFieldMetadata } from '../../../models/IFieldMetadata';
import { MetadataRow } from './MetadataRow';

export interface IMetadataPanelProps {
  loadFields: () => Promise<IFieldMetadata[]>;
  onDismiss: () => void;
  onSave: (fields: IFieldMetadata[]) => void;
}

export const MetadataPanel: React.FC<IMetadataPanelProps> = ({ loadFields, onDismiss, onSave }) => {
  const [editableFields, setEditableFields] = React.useState<IFieldMetadata[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string>();

  React.useEffect(() => {
    let cancelled = false;
    setIsLoading(true);
    setError(undefined);

    loadFields()
      .then((fields) => {
        if (!cancelled) {
          setEditableFields(fields);
          setIsLoading(false);
        }
      })
      .catch((err) => {
        if (!cancelled) {
          setError(err.message || 'Failed to load fields');
          setIsLoading(false);
        }
      });

    return () => { cancelled = true; };
  }, [loadFields]);

  const handleDescriptionChange = React.useCallback((id: string, description: string): void => {
    setEditableFields(prev => prev.map(f => f.id === id ? { ...f, description } : f));
  }, []);

  const handleTypeChange = React.useCallback((id: string, type: FieldType): void => {
    setEditableFields(prev => prev.map(f => f.id === id ? { ...f, type } : f));
  }, []);

  const handleSave = React.useCallback((): void => {
    onSave(editableFields);
  }, [editableFields, onSave]);

  if (isLoading) {
    return (
      <Stack style={{ padding: 24 }} horizontalAlign="center">
        <Spinner size={SpinnerSize.large} label="Loading fields..." />
      </Stack>
    );
  }

  if (error) {
    return (
      <Stack style={{ padding: 24 }}>
        <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }} style={{ marginTop: 16 }}>
          <DefaultButton text="Close" onClick={onDismiss} />
        </Stack>
      </Stack>
    );
  }

  return (
    <Stack style={{ padding: 24, maxWidth: 720 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 16 }}>
        <h2 style={{ margin: 0 }}>Field Metadata</h2>
        <IconButton
          iconProps={{ iconName: 'Cancel' }}
          ariaLabel="Close"
          onClick={onDismiss}
        />
      </Stack>

      <div style={{ maxHeight: 400, overflowY: 'auto', marginBottom: 16 }}>
        <div style={{ display: 'flex', gap: 12, marginBottom: 8, fontWeight: 600, fontSize: 12, color: '#605e5c' }}>
          <span style={{ width: 140, flexShrink: 0 }}>Field</span>
          <span style={{ width: 80 }}>Type</span>
          <span style={{ flex: 1, minWidth: 200 }}>Description</span>
          <span style={{ width: 180, flexShrink: 0 }}>Current Value</span>
        </div>
        {editableFields.map(field => (
          <MetadataRow
            key={field.id}
            field={field}
            onDescriptionChange={handleDescriptionChange}
            onTypeChange={handleTypeChange}
          />
        ))}
      </div>

      <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
        <PrimaryButton text="Save" onClick={handleSave} />
        <DefaultButton text="Cancel" onClick={onDismiss} />
      </Stack>
    </Stack>
  );
};
