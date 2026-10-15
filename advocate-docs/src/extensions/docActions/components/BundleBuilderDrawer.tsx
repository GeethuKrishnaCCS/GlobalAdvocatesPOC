import * as React from 'react';
import {
  Button, Caption1, DrawerBody, DrawerFooter, DrawerHeader, DrawerHeaderTitle, Field, Input, MessageBar, MessageBarBody,
  OverlayDrawer, Spinner, Tooltip, makeStyles, tokens
} from '@fluentui/react-components';
import { Add20Regular, Dismiss20Regular, Save20Regular } from '@fluentui/react-icons';
import * as strings from 'DocActionsCommandSetStrings';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import type { IBundle } from '../models/IBundle';
import type { DocumentQueryService } from '../services/DocumentQueryService';
import type { BundleService } from '../services/BundleService';
import { useBundleDraft } from '../hooks/useBundleDraft';
import { BundleDocumentList } from './BundleDocumentList';
import { DocumentPreviewPane } from './DocumentPreviewPane';
import { AddDocumentsPanel } from './AddDocumentsPanel';
import { ConfirmDialog } from './common/ConfirmDialog';
import { format } from '../utils/format';

export interface IBundleBuilderDrawerProps {
  webAbsoluteUrl: string;
  /** Selection to start a new bundle from (ignored when `bundle` is given). */
  documents: ReadonlyArray<ISelectedDocument>;
  /** Saved bundle to edit. */
  bundle?: IBundle;
  queryService: DocumentQueryService;
  bundleService: BundleService;
  libraryTitle: string;
  onClose: () => void;
}

const useStyles = makeStyles({
  drawer: { width: '92vw' },
  body: { display: 'grid', gridTemplateColumns: 'minmax(18rem, 1fr) 2fr', gap: tokens.spacingHorizontalL, height: '100%', minHeight: 0 },
  side: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, minHeight: 0 },
  sideHeader: { display: 'flex', alignItems: 'flex-end', gap: tokens.spacingHorizontalS },
  title: { flexGrow: 1 },
  footer: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }
});

/** Bundle builder: create or edit a bundle — ordered members (or the finder) left, full preview right. */
export const BundleBuilderDrawer: React.FC<IBundleBuilderDrawerProps> = props => {
  const { webAbsoluteUrl, documents, bundle, queryService, bundleService, libraryTitle, onClose } = props;
  const styles = useStyles();
  const draft = useBundleDraft(bundleService, documents, bundle);
  const [isAdding, setIsAdding] = React.useState<boolean>(!bundle && documents.length === 0);
  const [notice, setNotice] = React.useState<string | undefined>(undefined);
  const [finderPreview, setFinderPreview] = React.useState<ISelectedDocument | undefined>(undefined);
  const [confirmClose, setConfirmClose] = React.useState<boolean>(false);
  const previewDocument: ISelectedDocument | undefined = isAdding ? finderPreview : draft.computed.highlighted;

  const existingIds: ReadonlySet<string> = React.useMemo(
    () => new Set(draft.state.documents.map(d => d.uniqueId)), [draft.state.documents]);

  const requestClose = (): void => { if (draft.computed.isDirty) { setConfirmClose(true); } else { onClose(); } };
  const onAdd = (docs: ReadonlyArray<ISelectedDocument>): void => {
    setNotice(format(strings.AddedSummary, draft.actions.addMany(docs)));
    setFinderPreview(undefined);
    setIsAdding(false);
  };
  const onSave = (): void => {
    draft.actions.save()
      .then(ok => { if (ok) { setNotice(format(strings.BundleSaved, draft.state.title.trim())); } })
      .catch(() => undefined);
  };

  return (
    <OverlayDrawer open position="end" className={styles.drawer} onOpenChange={(_, d) => { if (!d.open) { requestClose(); } }}
      data-testid="bundle-builder">
      <DrawerHeader>
        <DrawerHeaderTitle action={
          <Tooltip content={strings.Close} relationship="label">
            <Button appearance="subtle" icon={<Dismiss20Regular />} onClick={requestClose} data-testid="bundle-builder-close" />
          </Tooltip>
        }>
          {bundle ? strings.EditBundleTitle : strings.BundleBuilderTitle}
        </DrawerHeaderTitle>
      </DrawerHeader>
      <DrawerBody className={styles.body}>
        {isAdding
          ? <AddDocumentsPanel service={queryService} libraryTitle={libraryTitle} existingIds={existingIds}
              previewId={finderPreview?.uniqueId} onPreview={setFinderPreview}
              onAdd={onAdd} onBack={() => { setFinderPreview(undefined); setIsAdding(false); }} />
          : (
            <div className={styles.side}>
              <div className={styles.sideHeader}>
                <Field label={strings.BundleName} required className={styles.title}>
                  <Input value={draft.state.title} onChange={(_, d) => draft.actions.setTitle(d.value)} data-testid="bundle-builder-name" />
                </Field>
                <Button appearance="secondary" icon={<Add20Regular />} onClick={() => { setNotice(undefined); setIsAdding(true); }}
                  data-testid="bundle-builder-add">{strings.AddDocumentsButton}</Button>
              </div>
              {notice && <MessageBar intent="success" data-testid="bundle-builder-notice"><MessageBarBody>{notice}</MessageBarBody></MessageBar>}
              {draft.state.saveError && (
                <MessageBar intent="error" layout="multiline" data-testid="bundle-builder-save-error">
                  <MessageBarBody>{draft.state.saveError}</MessageBarBody>
                </MessageBar>
              )}
              {draft.computed.count === 0
                ? <Caption1 data-testid="bundle-builder-empty">{strings.NoDocuments}</Caption1>
                : <BundleDocumentList draft={draft} />}
            </div>
          )}
        <DocumentPreviewPane webAbsoluteUrl={webAbsoluteUrl} document={previewDocument} />
      </DrawerBody>
      <DrawerFooter className={styles.footer}>
        <Button appearance="primary" icon={draft.state.isSaving ? <Spinner size="tiny" /> : <Save20Regular />}
          disabled={!draft.computed.canSave || draft.state.isSaving} onClick={onSave} data-testid="bundle-builder-save">
          {strings.SaveBundle}
        </Button>
        <Tooltip content={strings.BuildNotConnected} relationship="description">
          <Button appearance="secondary" disabledFocusable data-testid="bundle-builder-build">{strings.BuildBundle}</Button>
        </Tooltip>
        <Caption1 data-testid="bundle-builder-count">{format(strings.DocumentCount, draft.computed.count)}</Caption1>
      </DrawerFooter>
      <ConfirmDialog open={confirmClose} title={strings.DiscardTitle} body={strings.DiscardBody}
        confirmLabel={strings.Discard} cancelLabel={strings.KeepEditing}
        onConfirm={() => { setConfirmClose(false); onClose(); }} onCancel={() => setConfirmClose(false)} testId="bundle-builder-discard" />
    </OverlayDrawer>
  );
};
