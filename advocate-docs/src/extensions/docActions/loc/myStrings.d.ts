declare interface IDocActionsCommandSetStrings {
  Preview: string;
  CreateBundle: string;
  NotAvailableYet: string;

  Close: string;
  Next: string;
  Previous: string;
  PreviewPosition: string;
  PreviewFrameTitle: string;
  PreviewLoading: string;
  NothingToPreview: string;

  BundleBuilderTitle: string;
  SelectionSummary: string;
  DocumentCount: string;
  BuildBundle: string;
  BuildNotConnected: string;
  DocumentActions: string;
  MoveUp: string;
  MoveDown: string;
  RemoveFromBundle: string;
  NoDocuments: string;

  AddDocumentsButton: string;
  AddedSummary: string;
  ModeFolder: string;
  ModeFilter: string;
  ModeSearch: string;
  FolderPathLabel: string;
  FolderWithCount: string;
  NoSubfolders: string;
  IncludeSubfolders: string;
  FilterDocType: string;
  FilterDateFrom: string;
  FilterDateTo: string;
  FilterFileTypes: string;
  SearchLabel: string;
  SearchHint: string;
  FindDocuments: string;
  BackToBundle: string;
  FoundSummary: string;
  ResultTruncated: string;
  AddSelected: string;
  SelectAllResults: string;
  FinderResultsLabel: string;
  PickDocument: string;
  InBundle: string;
  SendGroup: string;
  BundleGroup: string;
  Send: string;
  SendAsPdf: string;
  SendVersion: string;
  SendVersionAs: string;
  AddToBundle: string;
  OpenBundle: string;
  EditBundleTitle: string;
  BundleName: string;
  SaveBundle: string;
  BundleSaved: string;
  DiscardTitle: string;
  DiscardBody: string;
  Discard: string;
  KeepEditing: string;
  Cancel: string;
  OpenBundleTitle: string;
  AddToBundleTitle: string;
  LoadingBundles: string;
  NoSavedBundles: string;
  SavedBundles: string;
  BundleOptionMeta: string;
  AddToBundleResult: string;
  AddToBundleConfirm: string;
}

declare module 'DocActionsCommandSetStrings' {
  const strings: IDocActionsCommandSetStrings;
  export = strings;
}


