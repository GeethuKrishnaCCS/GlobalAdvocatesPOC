declare interface IDocActionsCommandSetStrings {
  Preview: string;
  CreateBundle: string;
  EmailVersion: string;
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
  AddDocuments: string;
}

declare module 'DocActionsCommandSetStrings' {
  const strings: IDocActionsCommandSetStrings;
  export = strings;
}
