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
import type { IDocumentContext } from '../../../models/IDocumentContext';
import type { ILlmExtractionService } from '../../../services';
import { buildExtractionRequest } from '../../../services';
import { MetadataRow } from './MetadataRow';

export interface IMetadataPanelProps {
  loadFields: () => Promise<FieldBase[]>;
  documentContext: IDocumentContext;
  llmService: ILlmExtractionService;
  onDismiss: () => void;
  onSave: (extractionFields: MetadataExtractionField[]) => void;
}

export const MetadataPanel: React.FC<IMetadataPanelProps> = ({
  loadFields,
  documentContext,
  llmService,
  onDismiss,
  onSave,
}) => {
  const [extractionFields, setExtractionFields] = React.useState<MetadataExtractionField[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [isExtracting, setIsExtracting] = React.useState(false);
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

  const handleExtract = React.useCallback(async (): Promise<void> => {
    setIsExtracting(true);
    setError(undefined);

    try {
      // Build schema from extraction fields
      const schema = extractionFields.map((ef) => ef.toSchema());
      const request = buildExtractionRequest(documentContext, schema);

      // Call LLM extraction service
      const response = await llmService.extract(request);

      // Update extraction fields with results and notify parent
      // Use functional update to ensure we have latest state
      setExtractionFields((prev) => {
        const updated = prev.map((ef) => {
          const result = response.results.find((r) => r.fieldName === ef.field.title);
          if (result) {
            ef.extractedValue = result.value;
            ef.confidence = result.confidence;
          }
          return ef;
        });
        onSave(updated);
        return updated;
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Extraction failed');
    } finally {
      setIsExtracting(false);
    }
  }, [extractionFields, documentContext, llmService, onSave]);

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
    <Stack style={{ padding: 24, width: 900, maxWidth: '95vw', height: '80vh', maxHeight: 700 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 16, flexShrink: 0 }}>
        <h2 style={{ margin: 0 }}>Field Metadata</h2>
        <IconButton
          iconProps={{ iconName: 'Cancel' }}
          ariaLabel="Close"
          onClick={onDismiss}
          disabled={isExtracting}
        />
      </Stack>

      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setError(undefined)}
          style={{ marginBottom: 16, flexShrink: 0 }}
        >
          {error}
        </MessageBar>
      )}

      <div style={{ flex: 1, overflow: 'auto', marginBottom: 16 }}>
        {extractionFields.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.info}>
            No fields available for metadata extraction.
          </MessageBar>
        ) : (
          <>
            {/* Column headers */}
            <div style={{ display: 'flex', gap: 12, marginBottom: 8, fontWeight: 600, fontSize: 12, color: '#605e5c', position: 'sticky', top: 0, backgroundColor: 'white', paddingBottom: 8, borderBottom: '1px solid #edebe9' }}>
              <span style={{ width: 120, flexShrink: 0 }}>Field</span>
              <span style={{ width: 100, flexShrink: 0 }}>Type</span>
              <span style={{ flex: 1, minWidth: 150 }}>Description</span>
              <span style={{ width: 120, flexShrink: 0 }}>Current Value</span>
              <span style={{ width: 150, flexShrink: 0 }}>Extracted Value</span>
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

      <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }} style={{ flexShrink: 0 }}>
        <PrimaryButton
          text={isExtracting ? 'Extracting...' : 'Extract'}
          onClick={handleExtract}
          disabled={isExtracting || extractionFields.length === 0}
          iconProps={isExtracting ? undefined : { iconName: 'KeyPhraseExtraction' }}
        />
        <DefaultButton text="Cancel" onClick={onDismiss} disabled={isExtracting} />
      </Stack>
    </Stack>
  );
};
