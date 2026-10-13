import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'DocActionsCommandSetStrings';

/** No ClientSideComponentProperties yet; kept for the BaseListViewCommandSet generic. */
export interface IDocActionsCommandSetProperties {}

const LOG_SOURCE: string = 'DocActionsCommandSet';

/** Command ids — must match the keys in DocActionsCommandSet.manifest.json. */
const enum CommandId {
  Preview = 'PREVIEW',
  Merge = 'MERGE',
  EmailPdf = 'EMAIL_PDF',
  EmailVersion = 'EMAIL_VERSION'
}

const MIN_MERGE_SELECTION: number = 2;

/** Document actions for the advocate-firm POC: preview, merge to PDF (with index page), email variants. */
export default class DocActionsCommandSet extends BaseListViewCommandSet<IDocActionsCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized DocActionsCommandSet');

    this._updateCommandVisibility();
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case CommandId.Preview:
        this._notAvailableYet(strings.Preview);
        break;
      case CommandId.Merge:
        this._notAvailableYet(strings.Merge);
        break;
      case CommandId.EmailPdf:
        this._notAvailableYet(strings.EmailPdf);
        break;
      case CommandId.EmailVersion:
        this._notAvailableYet(strings.EmailVersion);
        break;
      default:
        throw new Error(`Unknown command: ${event.itemId}`);
    }
  }

  private _onListViewStateChanged = (): void => {
    this._updateCommandVisibility();
    this.raiseOnChange();
  };

  /** Single-document actions need exactly one row; merge needs two or more. */
  private _updateCommandVisibility(): void {
    const selectedCount: number = this.context.listView.selectedRows?.length ?? 0;

    this._setVisible(CommandId.Preview, selectedCount === 1);
    this._setVisible(CommandId.EmailPdf, selectedCount === 1);
    this._setVisible(CommandId.EmailVersion, selectedCount === 1);
    this._setVisible(CommandId.Merge, selectedCount >= MIN_MERGE_SELECTION);
  }

  private _setVisible(id: CommandId, visible: boolean): void {
    const command: Command | undefined = this.tryGetCommand(id);
    if (command) {
      command.visible = visible;
    }
  }

  // Scaffold placeholder per command; replaced feature-by-feature (see Tasks/advdocs-poc/task.md).
  private _notAvailableYet(feature: string): void {
    Dialog.alert(strings.NotAvailableYet.replace('{0}', feature)).catch((error: Error) => {
      Log.error(LOG_SOURCE, error);
    });
  }
}
