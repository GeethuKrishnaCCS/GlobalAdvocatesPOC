import * as React from 'react';
import type { DocumentQueryService } from '../services/DocumentQueryService';
import type { IDocumentFilter, IDocumentQueryResult, ILibraryFolder } from '../models/IDocumentQuery';

export type FinderMode = 'folder' | 'filter' | 'search';

export const EMPTY_FILTER: IDocumentFilter = { docType: '', dateFrom: '', dateTo: '', fileTypes: [] };

export interface IDocumentFinder {
  state: {
    mode: FinderMode;
    folderPath: string[];
    subfolders: ILibraryFolder[];
    includeSubfolders: boolean;
    filter: IDocumentFilter;
    searchText: string;
    result: IDocumentQueryResult | undefined;
    isLoading: boolean;
    error: string | undefined;
  };
  actions: {
    setMode: (mode: FinderMode) => void;
    openFolder: (folder: ILibraryFolder) => void;
    goToFolderLevel: (depth: number) => void;
    setIncludeSubfolders: (value: boolean) => void;
    setFilter: (filter: IDocumentFilter) => void;
    setSearchText: (text: string) => void;
    find: () => Promise<void>;
    clearResult: () => void;
  };
  computed: { currentFolderUrl: string; canFind: boolean };
}

/** Finds documents to add to a bundle by folder, metadata filter or search. */
export function useDocumentFinder(service: DocumentQueryService): IDocumentFinder {
  const [mode, setModeState] = React.useState<FinderMode>('folder');
  const [folderPath, setFolderPath] = React.useState<string[]>([]);
  const [subfolders, setSubfolders] = React.useState<ILibraryFolder[]>([]);
  const [includeSubfolders, setIncludeSubfolders] = React.useState<boolean>(true);
  const [filter, setFilter] = React.useState<IDocumentFilter>(EMPTY_FILTER);
  const [searchText, setSearchText] = React.useState<string>('');
  const [result, setResult] = React.useState<IDocumentQueryResult | undefined>(undefined);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  const currentFolderUrl: string = React.useMemo(
    () => [service.libraryServerRelativeUrl, ...folderPath].join('/'),
    [service, folderPath]
  );

  // Subfolders of the folder being browsed.
  React.useEffect(() => {
    let cancelled: boolean = false;
    service.getSubfolders(currentFolderUrl)
      .then(folders => { if (!cancelled) { setSubfolders(folders); } })
      .catch((e: Error) => { if (!cancelled) { setError(e.message); } });
    return () => { cancelled = true; };
  }, [service, currentFolderUrl]);

  const clearResult = React.useCallback((): void => { setResult(undefined); setError(undefined); }, []);

  const setMode = React.useCallback((next: FinderMode): void => { setModeState(next); clearResult(); }, [clearResult]);
  const openFolder = React.useCallback((folder: ILibraryFolder): void => {
    setFolderPath(path => [...path, folder.name]);
    clearResult();
  }, [clearResult]);
  const goToFolderLevel = React.useCallback((depth: number): void => {
    setFolderPath(path => path.slice(0, depth));
    clearResult();
  }, [clearResult]);

  const canFind: boolean = mode === 'folder' || (mode === 'search' ? searchText.trim().length > 0 : true);

  const find = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setError(undefined);
    try {
      const next: IDocumentQueryResult =
        mode === 'folder' ? await service.getFolderDocuments(currentFolderUrl, includeSubfolders)
          : mode === 'filter' ? await service.getFilteredDocuments(filter)
            : await service.searchDocuments(searchText);
      setResult(next);
    } catch (e) {
      setResult(undefined);
      setError((e as Error).message);
    } finally {
      setIsLoading(false);
    }
  }, [mode, service, currentFolderUrl, includeSubfolders, filter, searchText]);

  return {
    state: { mode, folderPath, subfolders, includeSubfolders, filter, searchText, result, isLoading, error },
    actions: { setMode, openFolder, goToFolderLevel, setIncludeSubfolders, setFilter, setSearchText, find, clearResult },
    computed: { currentFolderUrl, canFind }
  };
}
