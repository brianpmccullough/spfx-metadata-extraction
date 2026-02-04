import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import type { FieldBase } from '../../models/fields';
import type { MetadataExtractionField } from '../../models/extraction';
import type { IDocumentContext } from '../../models/IDocumentContext';
import type { IMetadataExtractionService } from './IMetadataExtractionService';
import { MetadataPanel } from './components/MetadataPanel';

export class MetadataDialog extends BaseDialog {
  constructor(
    private readonly _service: IMetadataExtractionService,
    private readonly _documentContext: IDocumentContext
  ) {
    super();
  }

  protected render(): void {
    const loadFields = (): Promise<FieldBase[]> =>
      this._service.loadFields(this._documentContext);

    const handleSave = (extractionFields: MetadataExtractionField[]): void => {
      // TODO: call LLM extraction service with extractionFields
      console.log('Saved extraction fields:', extractionFields);
      this.close().catch(() => { /* close error */ });
    };

    ReactDOM.render(
      <MetadataPanel
        loadFields={loadFields}
        onDismiss={() => this.close()}
        onSave={handleSave}
      />,
      this.domElement
    );
  }

  protected onAfterClose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
  }
}
