import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import type { FieldBase } from '../../models/fields';
import type { MetadataExtractionField } from '../../models/extraction';
import type { IDocumentContext } from '../../models/IDocumentContext';
import type { IMetadataExtractionService } from './IMetadataExtractionService';
import type { ILlmExtractionService } from '../../services';
import { MetadataPanel } from './components/MetadataPanel';

export class MetadataDialog extends BaseDialog {
  constructor(
    private readonly _service: IMetadataExtractionService,
    private readonly _documentContext: IDocumentContext,
    private readonly _llmService: ILlmExtractionService
  ) {
    super();
  }

  protected render(): void {
    const loadFields = (): Promise<FieldBase[]> =>
      this._service.loadFields(this._documentContext);

    const handleSave = (extractionFields: MetadataExtractionField[]): void => {
      // Log extracted values for debugging (don't close dialog - user can review and close)
      console.log('Extraction complete:', extractionFields.map((ef) => ({
        field: ef.field.title,
        extractedValue: ef.extractedValue,
        currentValue: ef.field.formatForDisplay(),
      })));
    };

    const handleApply = async (
      fields: Array<{ internalName: string; value: string | number | boolean | null }>
    ): Promise<void> => {
      await this._service.applyFieldValues(this._documentContext, fields);
    };

    ReactDOM.render(
      <MetadataPanel
        loadFields={loadFields}
        documentContext={this._documentContext}
        llmService={this._llmService}
        onDismiss={() => this.close()}
        onSave={handleSave}
        onApply={handleApply}
      />,
      this.domElement
    );
  }

  protected onAfterClose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
  }
}
