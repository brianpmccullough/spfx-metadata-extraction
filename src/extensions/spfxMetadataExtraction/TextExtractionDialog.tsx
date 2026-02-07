import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import type { IDocumentContext } from '../../models/IDocumentContext';
import type { ITextExtractionService } from '../../services';
import { TextExtractionPanel } from './components/TextExtractionPanel';

export class TextExtractionDialog extends BaseDialog {
  constructor(
    private readonly _textExtractionService: ITextExtractionService,
    private readonly _documentContext: IDocumentContext
  ) {
    super();
  }

  protected render(): void {
    ReactDOM.render(
      <TextExtractionPanel
        documentContext={this._documentContext}
        textExtractionService={this._textExtractionService}
        onDismiss={() => this.close()}
      />,
      this.domElement
    );
  }

  protected onAfterClose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
  }
}
