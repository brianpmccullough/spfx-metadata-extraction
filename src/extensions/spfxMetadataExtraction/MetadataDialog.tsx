import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import type { IFieldMetadata } from '../../models/IFieldMetadata';
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
    const loadFields = (): Promise<IFieldMetadata[]> =>
      this._service.loadFieldMetadata(this._documentContext);

    const handleSave = (fields: IFieldMetadata[]): void => {
      // TODO: call this._service.saveFieldMetadata(this._documentContext, fields)
      console.log('Saved field metadata:', fields);
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
