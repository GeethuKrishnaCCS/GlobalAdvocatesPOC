import * as React from 'react';
import {
  Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Link, MessageBar, MessageBarBody, Spinner,
  makeStyles, tokens
} from '@fluentui/react-components';
import * as strings from 'DocActionsCommandSetStrings';
import type { BundleApiService } from '../services/BundleApiService';
import { usePrint } from '../hooks/usePrint';
import { format } from '../utils/format';

export interface IPrintDialogProps {
  api: BundleApiService;
  itemIds: number[];
  onClose: () => void;
}

const useStyles = makeStyles({
  content: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM },
  // Off-screen, not display:none: some browsers won't print a frame that isn't rendered.
  frame: { position: 'fixed', width: '1px', height: '1px', left: '-10000px', top: '0', border: 'none' }
});

/** Prepares the selected documents as one PDF and opens the browser's print dialog. */
export const PrintDialog: React.FC<IPrintDialogProps> = ({ api, itemIds, onClose }) => {
  const styles = useStyles();
  const { state, actions, frameRef } = usePrint(api, itemIds);

  return (
    <Dialog open modalType="modal" onOpenChange={(_, d) => { if (!d.open) { onClose(); } }}>
      <DialogSurface data-testid="print-dialog">
        <DialogBody>
          <DialogTitle>{strings.Print}</DialogTitle>
          <DialogContent className={styles.content}>
            {state.isLoading && (
              <Spinner size="small" label={format(strings.PrintPreparing, itemIds.length)} data-testid="print-dialog-preparing" />
            )}
            {state.error && (
              <MessageBar intent="error" data-testid="print-dialog-error"><MessageBarBody>{state.error}</MessageBarBody></MessageBar>
            )}
            {state.pdfUrl && (
              <MessageBar intent="success" data-testid="print-dialog-ready">
                <MessageBarBody>
                  {strings.PrintReady}{' '}
                  <Link href={state.pdfUrl} target="_blank" rel="noopener" data-testid="print-dialog-open">{strings.PrintOpenPdf}</Link>
                </MessageBarBody>
              </MessageBar>
            )}
            {state.pdfUrl && (
              <iframe ref={frameRef} src={state.pdfUrl} title={strings.Print} className={styles.frame}
                onLoad={actions.onFrameLoad} data-testid="print-dialog-frame" />
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!state.pdfUrl} onClick={actions.print} data-testid="print-dialog-print">{strings.PrintAgain}</Button>
            <Button onClick={onClose} data-testid="print-dialog-close">{strings.Close}</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
