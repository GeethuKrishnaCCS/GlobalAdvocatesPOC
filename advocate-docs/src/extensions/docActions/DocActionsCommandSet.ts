import * as React from 'react';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'DocActionsCommandSetStrings';
import { PanelHost } from './utils/panelHost';
import { toSelectedDocuments } from './utils/documentMapper';
import { format } from './utils/format';
import { PreviewDialog } from './components/PreviewDialog';
import { BundleBuilderDrawer } from './components/BundleBuilderDrawer';
import { DocumentQueryService } from './services/DocumentQueryService';

/** ClientSideComponentProperties: URL names (not titles) of the matter site libraries. */
export interface IDocActionsCommandSetProperties {
  documentsLibrary?: string;
  bundlesLibrary?: string;
}

const LOG_SOURCE: string = 'DocActionsCommandSet';

/** Command ids — must match the keys in DocActionsCommandSet.manifest.json. */
const enum CommandId {
  Preview = 'PREVIEW',
  CreateBundle = 'CREATE_BUNDLE',
  EmailVersion = 'EMAIL_VERSION'
}

/** Library kinds the commands care about; anything else shows no commands. */
const enum LibraryKind {
  Documents,
  Bundles,
  Other
}

const DEFAULT_DOCUMENTS_LIBRARY: string = 'Documents';
const DEFAULT_BUNDLES_LIBRARY: string = 'Bundles';

/** Document actions for the advocate-firm POC, all opened from the library views. */
export default class DocActionsCommandSet extends BaseListViewCommandSet<IDocActionsCommandSetProperties> {
  private readonly _panels: PanelHost = new PanelHost();

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized DocActionsCommandSet');
    this._updateCommandVisibility();
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case CommandId.Preview:
        this._openPreview();
        break;
      case CommandId.CreateBundle:
        this._openBundleBuilder();
        break;
      case CommandId.EmailVersion:
        this._notAvailableYet(strings.EmailVersion);
        break;
      default:
        throw new Error(`Unknown command: ${event.itemId}`);
    }
  }

  protected onDispose(): void {
    this._panels.close();
    super.onDispose();
  }

  private get _webAbsoluteUrl(): string {
    return this.context.pageContext.web.absoluteUrl;
  }

  private _openPreview(): void {
    const documents = toSelectedDocuments(this.context.listView.selectedRows ?? []);
    this._panels.open(close => React.createElement(PreviewDialog, { webAbsoluteUrl: this._webAbsoluteUrl, documents, onClose: close }));
  }

  private _openBundleBuilder(): void {
    const documents = toSelectedDocuments(this.context.listView.selectedRows ?? []);
    const loadedRowCount: number = this.context.listView.rows.length;
    const list = this.context.listView.list;
    Log.info(LOG_SOURCE, `Bundle builder: ${documents.length} selected of ${loadedRowCount} loaded rows`);
    const queryService: DocumentQueryService = new DocumentQueryService(
      this.context.spHttpClient, this._webAbsoluteUrl, list?.serverRelativeUrl ?? '');
    this._panels.open(close => React.createElement(BundleBuilderDrawer, {
      webAbsoluteUrl: this._webAbsoluteUrl, documents, loadedRowCount, queryService,
      libraryTitle: list?.title ?? '', onClose: close
    }));
  }

  private _onListViewStateChanged = (): void => {
    this._updateCommandVisibility();
    this.raiseOnChange();
  };

  /** Which matter library this view shows, by URL name (never title). */
  private _libraryKind(): LibraryKind {
    const url: string = (this.context.listView.list?.serverRelativeUrl ?? '').toLowerCase();
    const leaf: string = url.split('/').pop() ?? '';
    if (leaf === (this.properties.documentsLibrary ?? DEFAULT_DOCUMENTS_LIBRARY).toLowerCase()) { return LibraryKind.Documents; }
    if (leaf === (this.properties.bundlesLibrary ?? DEFAULT_BUNDLES_LIBRARY).toLowerCase()) { return LibraryKind.Bundles; }
    return LibraryKind.Other;
  }

  private _updateCommandVisibility(): void {
    const selectedCount: number = this.context.listView.selectedRows?.length ?? 0;
    const kind: LibraryKind = this._libraryKind();
    const inMatterLibrary: boolean = kind !== LibraryKind.Other;

    this._setVisible(CommandId.Preview, inMatterLibrary && selectedCount >= 1);
    // No selection needed: documents can be added by folder, filter or search inside the builder.
    this._setVisible(CommandId.CreateBundle, kind === LibraryKind.Documents);
    this._setVisible(CommandId.EmailVersion, inMatterLibrary && selectedCount === 1);
  }

  private _setVisible(id: CommandId, visible: boolean): void {
    const command: Command | undefined = this.tryGetCommand(id);
    if (command) {
      command.visible = visible;
    }
  }

  // Placeholder until the email slice lands (task.md slice 4).
  private _notAvailableYet(feature: string): void {
    Dialog.alert(format(strings.NotAvailableYet, feature)).catch((error: Error) => {
      Log.error(LOG_SOURCE, error);
    });
  }
}
