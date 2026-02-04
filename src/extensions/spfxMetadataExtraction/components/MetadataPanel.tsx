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
import type { FieldBase } from '../../../models/fields';
import {
  MetadataExtractionField,
  MetadataExtractionFieldType,
} from '../../../models/extraction';
import { MetadataRow } from './MetadataRow';

export interface IMetadataPanelProps {
  loadFields: () => Promise<FieldBase[]>;
  onDismiss: () => void;
  onSave: (extractionFields: MetadataExtractionField[]) => void;
}

export const MetadataPanel: React.FC<IMetadataPanelProps> = ({ loadFields, onDismiss, onSave }) => {
  const [extractionFields, setExtractionFields] = React.useState<MetadataExtractionField[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string>();

  React.useEffect(() => {
    let cancelled = false;
    setIsLoading(true);
    setError(undefined);

    loadFields()
      .then((loadedFields) => {
        if (!cancelled) {
          const wrapped = loadedFields.map((f) => new MetadataExtractionField(f));
          setExtractionFields(wrapped);
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

  const handleExtractionTypeChange = React.useCallback(
    (index: number, newType: MetadataExtractionFieldType): void => {
      setExtractionFields((prev) => {
        const updated = [...prev];
        updated[index].extractionType = newType;
        return updated;
      });
    },
    []
  );

  const handleDescriptionChange = React.useCallback(
    (index: number, newDescription: string): void => {
      setExtractionFields((prev) => {
        const updated = [...prev];
        updated[index].description = newDescription;
        return updated;
      });
    },
    []
  );

  const handleSave = React.useCallback((): void => {
    onSave(extractionFields);
  }, [extractionFields, onSave]);

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
    <Stack style={{ padding: 24, maxWidth: 800 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 16 }}>
        <h2 style={{ margin: 0 }}>Field Metadata</h2>
        <IconButton
          iconProps={{ iconName: 'Cancel' }}
          ariaLabel="Close"
          onClick={onDismiss}
        />
      </Stack>

      <div style={{ maxHeight: 500, overflowY: 'auto', marginBottom: 16 }}>
        {extractionFields.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.info}>
            No fields available for metadata extraction.
          </MessageBar>
        ) : (
          <>
            <div style={{ display: 'flex', gap: 12, marginBottom: 8, fontWeight: 600, fontSize: 12, color: '#605e5c' }}>
              <span style={{ width: 120, flexShrink: 0 }}>Field</span>
              <span style={{ width: 100, flexShrink: 0 }}>Type</span>
              <span style={{ flex: 1, minWidth: 200 }}>Description</span>
              <span style={{ width: 150, flexShrink: 0 }}>Current Value</span>
            </div>
            {extractionFields.map((extractionField, index) => (
              <MetadataRow
                key={extractionField.field.id}
                extractionField={extractionField}
                onExtractionTypeChange={(newType) => handleExtractionTypeChange(index, newType)}
                onDescriptionChange={(newDesc) => handleDescriptionChange(index, newDesc)}
              />
            ))}
          </>
        )}
      </div>

      <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
        <PrimaryButton text="Save" onClick={handleSave} />
        <DefaultButton text="Cancel" onClick={onDismiss} />
      </Stack>
    </Stack>
  );
};
