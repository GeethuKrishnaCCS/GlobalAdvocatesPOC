import * as React from 'react';
import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle } from '@fluentui/react-components';

export interface IConfirmDialogProps {
  open: boolean;
  title: string;
  body: string;
  confirmLabel: string;
  cancelLabel: string;
  onConfirm: () => void;
  onCancel: () => void;
  testId: string;
}

/** Alert-style confirmation; Cancel takes initial focus so Enter never confirms by accident. */
export const ConfirmDialog: React.FC<IConfirmDialogProps> = ({ open, title, body, confirmLabel, cancelLabel, onConfirm, onCancel, testId }) => (
  <Dialog open={open} modalType="alert" onOpenChange={(_, d) => { if (!d.open) { onCancel(); } }}>
    <DialogSurface data-testid={testId}>
      <DialogBody>
        <DialogTitle>{title}</DialogTitle>
        <DialogContent>{body}</DialogContent>
        <DialogActions>
          <Button appearance="secondary" onClick={onCancel} autoFocus data-testid={`${testId}-cancel`}>{cancelLabel}</Button>
          <Button appearance="primary" onClick={onConfirm} data-testid={`${testId}-confirm`}>{confirmLabel}</Button>
        </DialogActions>
      </DialogBody>
    </DialogSurface>
  </Dialog>
);
