import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import type { FieldBase } from '../../models/fields';
import type { MetadataExtractionField } from '../../models/extraction';
import type { IDocumentContext } from '../../models/IDocumentContext';
import type { IMetadataExtractionService } from './IMetadataExtractionService';
import type { ILlmExtractionService } from '../../services';
import type { ISharePointRestClient } from '../../clients/ISharePointRestClient';
import { MetadataPanel } from './components/MetadataPanel';

export class MetadataDialog extends BaseDialog {
  constructor(
    private readonly _service: IMetadataExtractionService,
    private readonly _documentContext: IDocumentContext,
    private readonly _llmService: ILlmExtractionService,
    private readonly _spoClient: ISharePointRestClient
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
      const { webUrl, listId, itemId } = this._documentContext;
      const url = `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/ValidateUpdateListItem()`;

      const formValues = fields.map((f) => ({
        FieldName: f.internalName,
        FieldValue: String(f.value),
      }));

      await this._spoClient.post(url, { formValues });
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
