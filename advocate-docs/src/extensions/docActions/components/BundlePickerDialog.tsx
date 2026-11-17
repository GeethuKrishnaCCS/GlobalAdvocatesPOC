import * as React from 'react';
import {
  Body1, Button, Caption1, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle,
  MessageBar, MessageBarBody, Radio, RadioGroup, Skeleton, SkeletonItem, Spinner, makeStyles, tokens
} from '@fluentui/react-components';
import * as strings from 'DocActionsCommandSetStrings';
import type { BundleService } from '../services/BundleService';
import type { IBundle } from '../models/IBundle';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import { useBundlePicker } from '../hooks/useBundlePicker';
import { format } from '../utils/format';

export interface IBundlePickerDialogProps {
  service: BundleService;
  /** 'open' loads the chosen bundle for editing; 'add' appends `documents` to it. */
  mode: 'open' | 'add';
  documents?: ReadonlyArray<ISelectedDocument>;
  onOpen?: (bundle: IBundle) => void;
  onClose: () => void;
}

const useStyles = makeStyles({
  list: { maxHeight: '50vh', overflowY: 'auto' },
  option: { display: 'flex', flexDirection: 'column' },
  content: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM }
});

/** Pick a saved bundle to open, or to add the selected documents to. */
export const BundlePickerDialog: React.FC<IBundlePickerDialogProps> = ({ service, mode, documents = [], onOpen, onClose }) => {
  const styles = useStyles();
  const { state, actions, computed } = useBundlePicker(service);
  const done: boolean = mode === 'add' && state.addResult !== undefined;

  const onPrimary = (): void => {
    if (mode === 'open') {
      actions.open().then(b => { if (b && onOpen) { onOpen(b); } }).catch(() => undefined);
    } else {
      actions.add(documents).catch(() => undefined);
    }
  };

  return (
    <Dialog open modalType="modal" onOpenChange={(_, d) => { if (!d.open && !state.isBusy) { onClose(); } }}>
      <DialogSurface data-testid="bundle-picker">
        <DialogBody>
          <DialogTitle>{mode === 'open' ? strings.OpenBundleTitle : format(strings.AddToBundleTitle, documents.length)}</DialogTitle>
          <DialogContent className={styles.content}>
            {state.bundles === undefined && (
              <Skeleton aria-label={strings.LoadingBundles} data-testid="bundle-picker-loading">
                <SkeletonItem size={16} /><SkeletonItem size={16} /><SkeletonItem size={16} />
              </Skeleton>
            )}
            {state.bundles?.length === 0 && !state.error && <Body1 data-testid="bundle-picker-empty">{strings.NoSavedBundles}</Body1>}
            {state.bundles && state.bundles.length > 0 && !done && (
              <RadioGroup className={styles.list} value={state.selectedId ?? ''} onChange={(_, d) => actions.select(d.value)}
                aria-label={strings.SavedBundles} data-testid="bundle-picker-list">
                {state.bundles.map(b => (
                  <Radio key={b.bundleId} value={b.bundleId} data-testid="bundle-picker-option" label={
                    <span className={styles.option}>
                      <Body1>{b.title}</Body1>
                      <Caption1>{format(strings.BundleOptionMeta, b.memberCount, b.status)}</Caption1>
                    </span>
                  } />
                ))}
              </RadioGroup>
            )}
            {done && state.addResult && (
              <MessageBar intent="success" data-testid="bundle-picker-result">
                <MessageBarBody>{format(strings.AddToBundleResult, state.addResult.added, computed.selected?.title ?? '', state.addResult.skipped)}</MessageBarBody>
              </MessageBar>
            )}
            {state.error && (
              <MessageBar intent="error" layout="multiline" data-testid="bundle-picker-error"><MessageBarBody>{state.error}</MessageBarBody></MessageBar>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" disabled={state.isBusy} onClick={onClose} data-testid="bundle-picker-cancel">
              {done ? strings.Close : strings.Cancel}
            </Button>
            {!done && (
              <Button appearance="primary" disabled={!computed.selected || state.isBusy} onClick={onPrimary}
                icon={state.isBusy ? <Spinner size="tiny" /> : undefined} data-testid="bundle-picker-confirm">
                {mode === 'open' ? strings.OpenBundle : format(strings.AddToBundleConfirm, documents.length)}
              </Button>
            )}
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
