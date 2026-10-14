import * as React from 'react';
import {
  Body1, Button, Caption1, Field, MessageBar, MessageBarBody, SearchBox, Spinner, Tab, TabList, makeStyles, tokens
} from '@fluentui/react-components';
import * as strings from 'DocActionsCommandSetStrings';
import type { DocumentQueryService } from '../services/DocumentQueryService';
import { MAX_RESULTS } from '../services/DocumentQueryService';
import { useDocumentFinder, type FinderMode } from '../hooks/useDocumentFinder';
import { useResultSelection } from '../hooks/useResultSelection';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import { FolderBrowser } from './FolderBrowser';
import { FilterFields } from './FilterFields';
import { FinderResultList } from './FinderResultList';
import { format } from '../utils/format';

export interface IAddDocumentsPanelProps {
  service: DocumentQueryService;
  libraryTitle: string;
  /** Unique ids (lower case) already in the bundle. */
  existingIds: ReadonlySet<string>;
  previewId: string | undefined;
  onPreview: (doc: ISelectedDocument) => void;
  onAdd: (docs: ReadonlyArray<ISelectedDocument>) => void;
  onBack: () => void;
}

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, minHeight: 0, overflowY: 'auto' },
  result: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, minHeight: 0 },
  actions: { display: 'flex', gap: tokens.spacingHorizontalS, flexWrap: 'wrap' }
});

/** Finds documents by folder, filter or search; the user picks individual results (or all) to add. */
export const AddDocumentsPanel: React.FC<IAddDocumentsPanelProps> = props => {
  const { service, libraryTitle, existingIds, previewId, onPreview, onAdd, onBack } = props;
  const styles = useStyles();
  const finder = useDocumentFinder(service);
  const { state, actions, computed } = finder;
  const selection = useResultSelection(state.result?.documents, existingIds);

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
          <Body1>{format(strings.FoundSummary, state.result.documents.length, selection.computed.selectable.length)}</Body1>
          {state.result.truncated && <Caption1 data-testid="add-documents-truncated">{format(strings.ResultTruncated, MAX_RESULTS)}</Caption1>}
          {state.result.documents.length > 0 && (
            <FinderResultList results={state.result.documents} existingIds={existingIds} selection={selection}
              previewId={previewId} onPreview={onPreview} />
          )}
          <div className={styles.actions}>
            <Button appearance="primary" disabled={selection.computed.selected.length === 0}
              onClick={() => onAdd(selection.computed.selected)} data-testid="add-documents-add">
              {format(strings.AddSelected, selection.computed.selected.length)}
            </Button>
          </div>
        </div>
      )}
    </div>
  );
};
