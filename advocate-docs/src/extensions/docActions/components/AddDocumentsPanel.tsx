import * as React from 'react';
import {
  Body1, Button, Caption1, Field, MessageBar, MessageBarBody, SearchBox, Spinner, Tab, TabList, makeStyles, tokens
} from '@fluentui/react-components';
import * as strings from 'DocActionsCommandSetStrings';
import type { DocumentQueryService } from '../services/DocumentQueryService';
import { MAX_RESULTS } from '../services/DocumentQueryService';
import { useDocumentFinder, type FinderMode } from '../hooks/useDocumentFinder';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import { FolderBrowser } from './FolderBrowser';
import { FilterFields } from './FilterFields';
import { format } from '../utils/format';

export interface IAddDocumentsPanelProps {
  service: DocumentQueryService;
  libraryTitle: string;
  /** Unique ids already in the bundle, to count what is new. */
  existingIds: ReadonlySet<string>;
  onAdd: (docs: ReadonlyArray<ISelectedDocument>) => void;
  onBack: () => void;
}

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, minHeight: 0, overflowY: 'auto' },
  result: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS },
  actions: { display: 'flex', gap: tokens.spacingHorizontalS, flexWrap: 'wrap' }
});

/** Adds documents to the bundle by folder, metadata filter or search — beyond what the view selection can reach. */
export const AddDocumentsPanel: React.FC<IAddDocumentsPanelProps> = ({ service, libraryTitle, existingIds, onAdd, onBack }) => {
  const styles = useStyles();
  const finder = useDocumentFinder(service);
  const { state, actions, computed } = finder;

  const fresh: ISelectedDocument[] = React.useMemo(
    () => (state.result?.documents ?? []).filter(d => !existingIds.has(d.uniqueId.toLowerCase())),
    [state.result, existingIds]
  );

  const onFind = (): void => { actions.find().catch(() => undefined); };

  return (
    <div className={styles.root} data-testid="add-documents">
      <TabList selectedValue={state.mode} onTabSelect={(_, d) => actions.setMode(d.value as FinderMode)} data-testid="add-documents-mode">
        <Tab value="folder" data-testid="add-documents-mode-folder">{strings.ModeFolder}</Tab>
        <Tab value="filter" data-testid="add-documents-mode-filter">{strings.ModeFilter}</Tab>
        <Tab value="search" data-testid="add-documents-mode-search">{strings.ModeSearch}</Tab>
      </TabList>

      {state.mode === 'folder' && <FolderBrowser finder={finder} libraryTitle={libraryTitle} />}
      {state.mode === 'filter' && <FilterFields finder={finder} />}
      {state.mode === 'search' && (
        <Field label={strings.SearchLabel} hint={strings.SearchHint}>
          <SearchBox value={state.searchText} onChange={(_, d) => actions.setSearchText(d.value)}
            onKeyDown={e => { if (e.key === 'Enter' && computed.canFind) { onFind(); } }} data-testid="add-documents-search" />
        </Field>
      )}

      <div className={styles.actions}>
        <Button appearance="primary" disabled={!computed.canFind || state.isLoading} onClick={onFind}
          icon={state.isLoading ? <Spinner size="tiny" /> : undefined} data-testid="add-documents-find">{strings.FindDocuments}</Button>
        <Button appearance="secondary" onClick={onBack} data-testid="add-documents-back">{strings.BackToBundle}</Button>
      </div>

      {state.error && (
        <MessageBar intent="error" layout="multiline" data-testid="add-documents-error"><MessageBarBody>{state.error}</MessageBarBody></MessageBar>
      )}
      {state.result && (
        <div className={styles.result} data-testid="add-documents-result">
          <Body1>{format(strings.FoundSummary, state.result.documents.length, fresh.length)}</Body1>
          {state.result.truncated && <Caption1 data-testid="add-documents-truncated">{format(strings.ResultTruncated, MAX_RESULTS)}</Caption1>}
          <div className={styles.actions}>
            <Button appearance="primary" disabled={fresh.length === 0} onClick={() => onAdd(fresh)}
              data-testid="add-documents-add">{format(strings.AddDocuments, fresh.length)}</Button>
          </div>
        </div>
      )}
    </div>
  );
};
