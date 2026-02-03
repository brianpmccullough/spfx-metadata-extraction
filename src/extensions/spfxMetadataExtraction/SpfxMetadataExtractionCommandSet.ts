import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { MetadataExtractionContext } from './MetadataExtractionContext';
import { SharePointRestClient } from '../../clients/SharePointRestClient';
import { GraphClient } from '../../clients/GraphClient';
import { MetadataExtractionService } from './MetadataExtractionService';
import { MetadataDialog } from './MetadataDialog';

export interface ISpfxMetadataExtractionCommandSetProperties {
  allowedFileTypes: string[];
}

enum Commands {
  Extract = 'Extract'
}

const LOG_SOURCE: string = 'SpfxMetadataExtractionCommandSet';

export default class SpfxMetadataExtractionCommandSet extends BaseListViewCommandSet<ISpfxMetadataExtractionCommandSetProperties> {

  private _metadataExtractionService!: MetadataExtractionService;
  private _allowedFileTypes!: string[];

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SpfxMetadataExtractionCommandSet');

    const command: Command = this.tryGetCommand(Commands.Extract);
    command.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    this._allowedFileTypes = this.properties.allowedFileTypes ?? ['.pdf', '.doc', '.docx'];
    const sharePointRestClient = new SharePointRestClient(this.context.spHttpClient);
    const msGraphClient = await this.context.msGraphClientFactory.getClient('3');
    const graphClient = new GraphClient(msGraphClient);
    this._metadataExtractionService = new MetadataExtractionService(sharePointRestClient, graphClient);
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case Commands.Extract: {
        const context = new MetadataExtractionContext(this.context, this._allowedFileTypes);
        if (!context.canExecute) {
          return;
        }

        const dialog = new MetadataDialog(this._metadataExtractionService, context.documentContext!);
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

    const command: Command = this.tryGetCommand(Commands.Extract);
    if (command) {
      const context = new MetadataExtractionContext(this.context, this._allowedFileTypes);
      command.visible = context.isVisible;
    }

    this.raiseOnChange();
  }
}
