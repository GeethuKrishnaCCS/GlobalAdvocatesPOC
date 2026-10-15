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
import { BundleService } from './services/BundleService';
import { BundlePickerDialog } from './components/BundlePickerDialog';
import type { IBundle } from './models/IBundle';
import type { ISelectedDocument } from './models/ISelectedDocument';

/** ClientSideComponentProperties: URL names (not titles) of the matter site libraries and lists. */
export interface IDocActionsCommandSetProperties {
  documentsLibrary?: string;
  bundlesLibrary?: string;
  /** Under /Lists/. */
  bundleRegistryList?: string;
  /** Under /Lists/. */
  bundleItemsList?: string;
}

const LOG_SOURCE: string = 'DocActionsCommandSet';

/** Command ids — must match the keys in DocActionsCommandSet.manifest.json. */
const enum CommandId {
  Preview = 'PREVIEW',
  CreateBundle = 'CREATE_BUNDLE',
  OpenBundle = 'OPEN_BUNDLE',
  AddToBundle = 'ADD_TO_BUNDLE',
  Send = 'SEND',
  SendAsPdf = 'SEND_AS_PDF',
  SendVersion = 'SEND_VERSION',
  SendVersionAs = 'SEND_VERSION_AS'
}

/** Library kinds the commands care about; anything else shows no commands. */
const enum LibraryKind {
  Documents,
  Bundles,
  Other
}

const DEFAULT_DOCUMENTS_LIBRARY: string = 'Documents';
const DEFAULT_BUNDLES_LIBRARY: string = 'Bundles';
const DEFAULT_BUNDLE_REGISTRY_LIST: string = 'BundleRegistry';
const DEFAULT_BUNDLE_ITEMS_LIST: string = 'BundleItems';

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
        this._openBundleBuilder(toSelectedDocuments(this.context.listView.selectedRows ?? []));
        break;
      case CommandId.OpenBundle:
        this._openBundlePicker('open', []);
        break;
      case CommandId.AddToBundle:
        this._openBundlePicker('add', toSelectedDocuments(this.context.listView.selectedRows ?? []));
        break;
      case CommandId.Send:
        this._notAvailableYet(strings.Send);
        break;
      case CommandId.SendAsPdf:
        this._notAvailableYet(strings.SendAsPdf);
        break;
      case CommandId.SendVersion:
        this._notAvailableYet(strings.SendVersion);
        break;
      case CommandId.SendVersionAs:
        this._notAvailableYet(strings.SendVersionAs);
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

  private get _webServerRelativeUrl(): string {
    return this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, '');
  }

  private _documentsLibraryUrl(): string {
    return `${this._webServerRelativeUrl}/${this.properties.documentsLibrary ?? DEFAULT_DOCUMENTS_LIBRARY}`;
  }

  private _bundleService(): BundleService {
    const lists: string = `${this._webServerRelativeUrl}/Lists`;
    return new BundleService(this.context.spHttpClient, this._webAbsoluteUrl,
      `${lists}/${this.properties.bundleRegistryList ?? DEFAULT_BUNDLE_REGISTRY_LIST}`,
      `${lists}/${this.properties.bundleItemsList ?? DEFAULT_BUNDLE_ITEMS_LIST}`);
  }

  /** New bundle from a selection, or an existing saved bundle when `bundle` is given. */
  private _openBundleBuilder(documents: ISelectedDocument[], bundle?: IBundle): void {
    const queryService: DocumentQueryService = new DocumentQueryService(
      this.context.spHttpClient, this._webAbsoluteUrl, this._documentsLibraryUrl());
    const libraryTitle: string = this._libraryKind() === LibraryKind.Documents ? this.context.listView.list?.title ?? '' : '';
    this._panels.open(close => React.createElement(BundleBuilderDrawer, {
      webAbsoluteUrl: this._webAbsoluteUrl, documents, bundle, queryService,
      bundleService: this._bundleService(), libraryTitle: libraryTitle || (this.properties.documentsLibrary ?? DEFAULT_DOCUMENTS_LIBRARY),
      onClose: close
    }));
  }

  private _openBundlePicker(mode: 'open' | 'add', documents: ISelectedDocument[]): void {
    this._panels.open(close => React.createElement(BundlePickerDialog, {
      service: this._bundleService(), mode, documents,
      onOpen: (bundle: IBundle) => this._openBundleBuilder([], bundle),
      onClose: close
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
    this._setVisible(CommandId.OpenBundle, inMatterLibrary);
    this._setVisible(CommandId.AddToBundle, kind === LibraryKind.Documents && selectedCount >= 1);
    // Current version: one or more documents in one email. Previous version: exactly one document.
    this._setVisible(CommandId.Send, inMatterLibrary && selectedCount >= 1);
    this._setVisible(CommandId.SendAsPdf, inMatterLibrary && selectedCount >= 1);
    this._setVisible(CommandId.SendVersion, inMatterLibrary && selectedCount === 1);
    this._setVisible(CommandId.SendVersionAs, inMatterLibrary && selectedCount === 1);
  }

  private _setVisible(id: CommandId, visible: boolean): void {
    const command: Command | undefined = this.tryGetCommand(id);
    if (command) {
      command.visible = visible;
    }
  }

  // Placeholder until the send slice lands (task.md slice 4; needs Graph Mail.ReadWrite approval).
  private _notAvailableYet(feature: string): void {
    Dialog.alert(format(strings.NotAvailableYet, feature)).catch((error: Error) => {
      Log.error(LOG_SOURCE, error);
    });
  }
}

