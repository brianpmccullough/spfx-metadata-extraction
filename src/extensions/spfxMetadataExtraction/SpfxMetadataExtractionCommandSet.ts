import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { MetadataExtractionContext } from './MetadataExtractionContext';
import { SharePointRestClient } from '../../clients/SharePointRestClient';
import { MetadataExtractionService } from './MetadataExtractionService';
import { MetadataDialog } from './MetadataDialog';
import { TextExtractionDialog } from './TextExtractionDialog';
import { TextExtractionService, LlmExtractionService } from '../../services';
import type { ILlmExtractionService } from '../../services';
import { AadHttpClientWrapper } from '../../clients/AadHttpClientWrapper';

export interface ISpfxMetadataExtractionCommandSetProperties {
  allowedFileTypes: string[];
}

enum Commands {
  Extract = 'Extract',
  ExtractText = 'ExtractText'
}

const LOG_SOURCE: string = 'SpfxMetadataExtractionCommandSet';

export default class SpfxMetadataExtractionCommandSet extends BaseListViewCommandSet<ISpfxMetadataExtractionCommandSetProperties> {

  private _metadataExtractionService!: MetadataExtractionService;
  private _textExtractionService!: TextExtractionService;
  private _llmExtractionService!: ILlmExtractionService;
  private _sharePointRestClient!: SharePointRestClient;
  private _allowedFileTypes!: string[];

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SpfxMetadataExtractionCommandSet');

    const extractCommand: Command = this.tryGetCommand(Commands.Extract);
    extractCommand.visible = false;

    const extractTextCommand: Command = this.tryGetCommand(Commands.ExtractText);
    extractTextCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    this._allowedFileTypes = this.properties.allowedFileTypes ?? ['.pdf', '.doc', '.docx'];
    this._sharePointRestClient = new SharePointRestClient(this.context.spHttpClient);
    this._metadataExtractionService = new MetadataExtractionService(this._sharePointRestClient);

    // Create AAD HTTP client for extraction APIs with Entra ID authentication
    const aadHttpClient = await this.context.aadHttpClientFactory.getClient('d93c7720-43a9-4924-99c5-68464eb75b20');
    const aadHttpClientWrapper = new AadHttpClientWrapper(aadHttpClient);
    this._textExtractionService = new TextExtractionService(aadHttpClientWrapper);
    this._llmExtractionService = new LlmExtractionService(aadHttpClientWrapper);
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case Commands.Extract: {
        const context = new MetadataExtractionContext(this.context, this._allowedFileTypes);
        if (!context.canExecute) {
          return;
        }

        const dialog = new MetadataDialog(this._metadataExtractionService, context.documentContext!, this._llmExtractionService);
        dialog.show().catch((error) => {
          Log.error(LOG_SOURCE, error);
        });
        break;
      }
      case Commands.ExtractText: {
        const context = new MetadataExtractionContext(this.context, this._allowedFileTypes);
        if (!context.canExecute) {
          return;
        }

        const dialog = new TextExtractionDialog(this._textExtractionService, context.documentContext!);
        dialog.show().catch((error) => {
          Log.error(LOG_SOURCE, error);
        });
        break;
      }
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (_args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const context = new MetadataExtractionContext(this.context, this._allowedFileTypes);

    const extractCommand: Command = this.tryGetCommand(Commands.Extract);
    if (extractCommand) {
      extractCommand.visible = context.isVisible;
    }

    const extractTextCommand: Command = this.tryGetCommand(Commands.ExtractText);
    if (extractTextCommand) {
      extractTextCommand.visible = context.isVisible;
    }

    this.raiseOnChange();
  }
}
