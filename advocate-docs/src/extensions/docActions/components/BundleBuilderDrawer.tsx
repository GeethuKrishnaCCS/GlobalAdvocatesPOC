import * as React from 'react';
import {
  Button, Caption1, DrawerBody, DrawerFooter, DrawerHeader, DrawerHeaderTitle, MessageBar, MessageBarBody,
  OverlayDrawer, Tooltip, makeStyles, tokens
} from '@fluentui/react-components';
import { Add20Regular, Dismiss20Regular } from '@fluentui/react-icons';
import * as strings from 'DocActionsCommandSetStrings';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import type { DocumentQueryService } from '../services/DocumentQueryService';
import { useBundleDraft } from '../hooks/useBundleDraft';
import { BundleDocumentList } from './BundleDocumentList';
import { DocumentPreviewPane } from './DocumentPreviewPane';
import { AddDocumentsPanel } from './AddDocumentsPanel';
import { format } from '../utils/format';

export interface IBundleBuilderDrawerProps {
  webAbsoluteUrl: string;
  documents: ReadonlyArray<ISelectedDocument>;
  /** Rows currently loaded in the list view — selection can never exceed this. */
  loadedRowCount: number;
  queryService: DocumentQueryService;
  libraryTitle: string;
  onClose: () => void;
}

const useStyles = makeStyles({
  drawer: { width: '92vw' },
  body: {
    display: 'grid',
    gridTemplateColumns: 'minmax(18rem, 1fr) 2fr',
    gap: tokens.spacingHorizontalL,
    height: '100%',
    minHeight: 0
  },
  side: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, minHeight: 0 },
  sideHeader: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: tokens.spacingHorizontalS },
  footer: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }
});

/** Bundle builder: ordered member list (or the add-documents finder) on the left, full preview on the right. */
export const BundleBuilderDrawer: React.FC<IBundleBuilderDrawerProps> = props => {
  const { webAbsoluteUrl, documents, loadedRowCount, queryService, libraryTitle, onClose } = props;
  const styles = useStyles();
  const draft = useBundleDraft(documents);
  const [isAdding, setIsAdding] = React.useState<boolean>(documents.length === 0);
  const [lastAdded, setLastAdded] = React.useState<number | undefined>(undefined);

  const existingIds: ReadonlySet<string> = React.useMemo(
    () => new Set(draft.state.documents.map(d => d.uniqueId.toLowerCase())),
    [draft.state.documents]
  );

  const onAdd = (docs: ReadonlyArray<ISelectedDocument>): void => {
    setLastAdded(draft.actions.addMany(docs));
    setIsAdding(false);
  };

  return (
    <OverlayDrawer open position="end" className={styles.drawer} onOpenChange={(_, d) => { if (!d.open) { onClose(); } }}
      data-testid="bundle-builder">
      <DrawerHeader>
        <DrawerHeaderTitle action={
          <Tooltip content={strings.Close} relationship="label">
            <Button appearance="subtle" icon={<Dismiss20Regular />} onClick={onClose} data-testid="bundle-builder-close" />
          </Tooltip>
        }>
          {strings.BundleBuilderTitle}
        </DrawerHeaderTitle>
      </DrawerHeader>
      <DrawerBody className={styles.body}>
        {isAdding
          ? <AddDocumentsPanel service={queryService} libraryTitle={libraryTitle} existingIds={existingIds}
              onAdd={onAdd} onBack={() => setIsAdding(false)} />
          : (
            <div className={styles.side}>
              <div className={styles.sideHeader}>
                <Caption1 data-testid="bundle-builder-selection-summary">
                  {format(strings.SelectionSummary, documents.length, loadedRowCount)}
                </Caption1>
                <Button appearance="secondary" icon={<Add20Regular />} onClick={() => { setLastAdded(undefined); setIsAdding(true); }}
                  data-testid="bundle-builder-add">{strings.AddDocumentsButton}</Button>
              </div>
              {lastAdded !== undefined && (
                <MessageBar intent="success" data-testid="bundle-builder-added">
                  <MessageBarBody>{format(strings.AddedSummary, lastAdded)}</MessageBarBody>
                </MessageBar>
              )}
              {draft.computed.count === 0
                ? <Caption1 data-testid="bundle-builder-empty">{strings.NoDocuments}</Caption1>
                : <BundleDocumentList draft={draft} />}
            </div>
          )}
        <DocumentPreviewPane webAbsoluteUrl={webAbsoluteUrl} document={draft.computed.highlighted} />
      </DrawerBody>
      <DrawerFooter className={styles.footer}>
        <Tooltip content={strings.BuildNotConnected} relationship="description">
          <Button appearance="primary" disabledFocusable data-testid="bundle-builder-build">{strings.BuildBundle}</Button>
        </Tooltip>
        <Caption1 data-testid="bundle-builder-count">{format(strings.DocumentCount, draft.computed.count)}</Caption1>
      </DrawerFooter>
    </OverlayDrawer>
  );
};
