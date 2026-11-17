import * as React from 'react';
import {
  Body1Strong, Button, Caption1, Dialog, DialogBody, DialogSurface, DialogTitle, Tooltip, makeStyles, tokens
} from '@fluentui/react-components';
import { ChevronLeft20Regular, ChevronRight20Regular, Dismiss20Regular } from '@fluentui/react-icons';
import * as strings from 'DocActionsCommandSetStrings';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import { useDocumentNavigator } from '../hooks/useDocumentNavigator';
import { DocumentPreviewPane } from './DocumentPreviewPane';
import { format } from '../utils/format';

export interface IPreviewDialogProps {
  webAbsoluteUrl: string;
  documents: ReadonlyArray<ISelectedDocument>;
  onClose: () => void;
}

const useStyles = makeStyles({
  surface: { width: '92vw', maxWidth: '92vw', height: '92vh', display: 'flex', padding: tokens.spacingVerticalL },
  body: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, width: '100%', height: '100%' },
  header: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS },
  name: { flexGrow: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  pane: { flexGrow: 1, minHeight: 0 }
});

/** Full-screen preview of one or more selected documents with next / previous. */
export const PreviewDialog: React.FC<IPreviewDialogProps> = ({ webAbsoluteUrl, documents, onClose }) => {
  const styles = useStyles();
  const { actions, computed } = useDocumentNavigator(documents);

  const onKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === 'ArrowRight' && computed.hasNext) { actions.next(); }
    if (e.key === 'ArrowLeft' && computed.hasPrevious) { actions.previous(); }
  };

  return (
    <Dialog open modalType="modal" onOpenChange={(_, d) => { if (!d.open) { onClose(); } }}>
      <DialogSurface className={styles.surface} onKeyDown={onKeyDown} data-testid="preview-dialog">
        <DialogBody className={styles.body}>
          <DialogTitle as="div" className={styles.header}>
            <Body1Strong className={styles.name} data-testid="preview-dialog-name">{computed.current?.name}</Body1Strong>
            <Caption1 data-testid="preview-dialog-position">{format(strings.PreviewPosition, computed.position, computed.total)}</Caption1>
            <Tooltip content={strings.Previous} relationship="label">
              <Button appearance="subtle" icon={<ChevronLeft20Regular />} disabled={!computed.hasPrevious}
                onClick={actions.previous} data-testid="preview-dialog-previous" />
            </Tooltip>
            <Tooltip content={strings.Next} relationship="label">
              <Button appearance="subtle" icon={<ChevronRight20Regular />} disabled={!computed.hasNext}
                onClick={actions.next} data-testid="preview-dialog-next" />
            </Tooltip>
            <Tooltip content={strings.Close} relationship="label">
              <Button appearance="subtle" icon={<Dismiss20Regular />} onClick={onClose} data-testid="preview-dialog-close" />
            </Tooltip>
          </DialogTitle>
          <div className={styles.pane}>
            <DocumentPreviewPane webAbsoluteUrl={webAbsoluteUrl} document={computed.current} />
          </div>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
