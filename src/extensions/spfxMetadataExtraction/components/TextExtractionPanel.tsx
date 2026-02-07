import * as React from 'react';
import {
  DefaultButton,
  IconButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Stack,
} from '@fluentui/react';
import type { IDocumentContext } from '../../../models/IDocumentContext';
import type { ITextExtractionService } from '../../../services';

export interface ITextExtractionPanelProps {
  documentContext: IDocumentContext;
  textExtractionService: ITextExtractionService;
  onDismiss: () => void;
}

export const TextExtractionPanel: React.FC<ITextExtractionPanelProps> = ({
  documentContext,
  textExtractionService,
  onDismiss,
}) => {
  const [extractedText, setExtractedText] = React.useState<string>('');
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string>();

  React.useEffect(() => {
    let cancelled = false;
    setIsLoading(true);
    setError(undefined);

    const origin = new URL(documentContext.webUrl).origin;
    const absoluteUrl = `${origin}${documentContext.serverRelativeUrl}`;

    textExtractionService
      .extractText(absoluteUrl)
      .then((response) => {
        if (!cancelled) {
          setExtractedText(response.content);
          setIsLoading(false);
        }
      })
      .catch((err) => {
        if (!cancelled) {
          setError(err.message || 'Failed to extract text');
          setIsLoading(false);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [documentContext, textExtractionService]);

  if (isLoading) {
    return (
      <Stack style={{ padding: 24 }} horizontalAlign="center">
        <Spinner size={SpinnerSize.large} label="Extracting text..." />
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
    <Stack style={{ padding: 24, width: 800, maxWidth: '95vw', height: '80vh', maxHeight: 700 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 16, flexShrink: 0 }}>
        <h2 style={{ margin: 0 }}>Extracted Text: {documentContext.fileName}</h2>
        <IconButton
          iconProps={{ iconName: 'Cancel' }}
          ariaLabel="Close"
          onClick={onDismiss}
        />
      </Stack>

      <div
        style={{
          flex: 1,
          overflow: 'auto',
          marginBottom: 16,
          padding: 16,
          backgroundColor: '#faf9f8',
          border: '1px solid #edebe9',
          borderRadius: 4,
          fontFamily: 'monospace',
          fontSize: 13,
          lineHeight: 1.5,
          whiteSpace: 'pre-wrap',
          wordBreak: 'break-word',
        }}
      >
        {extractedText || 'No text content extracted.'}
      </div>

      <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }} style={{ flexShrink: 0 }}>
        <DefaultButton text="Close" onClick={onDismiss} />
      </Stack>
    </Stack>
  );
};
