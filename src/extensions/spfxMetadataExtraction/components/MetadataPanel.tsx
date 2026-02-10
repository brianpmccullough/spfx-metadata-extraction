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
import { MetadataRow } from './MetadataRow';

export interface IMetadataPanelProps {
  loadFields: () => Promise<FieldBase[]>;
  documentContext: IDocumentContext;
  llmService: ILlmExtractionService;
  onDismiss: () => void;
  onApply: (fields: Array<{ internalName: string; value: string | number | boolean | null }>) => Promise<void>;
}

export const MetadataPanel: React.FC<IMetadataPanelProps> = ({
  loadFields,
  documentContext,
  llmService,
  onDismiss,
  onApply,
}) => {
  const [extractionFields, setExtractionFields] = React.useState<MetadataExtractionField[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [isExtracting, setIsExtracting] = React.useState(false);
  const [isApplying, setIsApplying] = React.useState(false);
  const [hasExtracted, setHasExtracted] = React.useState(false);
  const [applyChecked, setApplyChecked] = React.useState<Set<string>>(new Set());
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
      setExtractionFields((prev) =>
        prev.map((ef, i) => {
          if (i !== index) return ef;
          const cloned = ef.clone();
          cloned.extractionType = newType;
          return cloned;
        })
      );
    },
    []
  );

  const handleDescriptionChange = React.useCallback(
    (index: number, newDescription: string): void => {
      setExtractionFields((prev) =>
        prev.map((ef, i) => {
          if (i !== index) return ef;
          const cloned = ef.clone();
          cloned.description = newDescription;
          return cloned;
        })
      );
    },
    []
  );

  const handleExtract = React.useCallback(async (): Promise<void> => {
    setIsExtracting(true);
    setError(undefined);
    setApplyChecked(new Set());

    // Clear previous extraction results
    setExtractionFields((prev) =>
      prev.map((ef) => {
        const cloned = ef.clone();
        cloned.extractedValue = null;
        cloned.confidence = null;
        return cloned;
      })
    );

    try {
      // Build extraction request from fields and document context
      const request = MetadataExtractionField.buildExtractionRequest(documentContext, extractionFields);

      // Call LLM extraction service
      const response = await llmService.extract(request);

      // Update extraction fields with results and notify parent
      // Use functional update to ensure we have latest state
      setExtractionFields((prev) => {
        const updated = prev.map((ef) => {
          const result = response.results.find((r) => r.fieldName === ef.field.title);
          if (result) {
            const cloned = ef.clone();
            cloned.extractedValue = result.value;
            cloned.confidence = result.confidence;
            return cloned;
          }
          return ef;
        });

        // Auto-check green confidence fields
        const autoChecked = new Set<string>();
        for (const ef of updated) {
          if (ef.confidence === 'green') {
            autoChecked.add(ef.field.id);
          }
        }
        setApplyChecked(autoChecked);

        return updated;
      });

      setHasExtracted(true);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Extraction failed');
    } finally {
      setIsExtracting(false);
    }
  }, [extractionFields, documentContext, llmService]);

  const handleApplyCheckedChange = React.useCallback(
    (fieldId: string, checked: boolean): void => {
      setApplyChecked((prev) => {
        const next = new Set(prev);
        if (checked) {
          next.add(fieldId);
        } else {
          next.delete(fieldId);
        }
        return next;
      });
    },
    []
  );

  const handleApply = React.useCallback(async (): Promise<void> => {
    setIsApplying(true);
    setError(undefined);

    try {
      const fieldsToApply = extractionFields
        .filter((ef) => applyChecked.has(ef.field.id))
        .map((ef) => ({
          internalName: ef.field.internalName,
          value: ef.extractedValue,
        }));

      await onApply(fieldsToApply);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Apply failed');
    } finally {
      setIsApplying(false);
    }
  }, [extractionFields, applyChecked, onApply]);

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
              <span style={{ width: 60, flexShrink: 0, textAlign: 'center' }}>Apply</span>
            </div>
            {extractionFields.map((extractionField, index) => (
              <MetadataRow
                key={extractionField.field.id}
                extractionField={extractionField}
                onExtractionTypeChange={(newType) => handleExtractionTypeChange(index, newType)}
                onDescriptionChange={(newDesc) => handleDescriptionChange(index, newDesc)}
                applyChecked={applyChecked.has(extractionField.field.id)}
                onApplyCheckedChange={(checked) => handleApplyCheckedChange(extractionField.field.id, checked)}
                isApplyEnabled={hasExtracted && extractionField.confidence !== 'red' && extractionField.confidence !== null}
              />
            ))}
          </>
        )}
      </div>

      <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }} style={{ flexShrink: 0 }}>
        <PrimaryButton
          text={isExtracting ? 'Extracting...' : 'Extract'}
          onClick={handleExtract}
          disabled={isExtracting || isApplying || extractionFields.length === 0}
          iconProps={isExtracting ? undefined : { iconName: 'KeyPhraseExtraction' }}
        />
        <PrimaryButton
          text={isApplying ? 'Applying...' : 'Apply'}
          onClick={handleApply}
          disabled={applyChecked.size === 0 || isApplying || isExtracting}
          iconProps={isApplying ? undefined : { iconName: 'Accept' }}
        />
        <DefaultButton text="Cancel" onClick={onDismiss} disabled={isExtracting || isApplying} />
      </Stack>
    </Stack>
  );
};
