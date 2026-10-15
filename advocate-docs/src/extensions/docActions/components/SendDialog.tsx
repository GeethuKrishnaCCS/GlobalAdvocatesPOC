import * as React from 'react';
import {
  Badge, Body1, Button, Caption1, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field,
  MessageBar, MessageBarBody, Radio, RadioGroup, Skeleton, SkeletonItem, Spinner, makeStyles, tokens
} from '@fluentui/react-components';
import { Open20Regular } from '@fluentui/react-icons';
import * as strings from 'DocActionsCommandSetStrings';
import type { GraphFileService } from '../services/GraphFileService';
import type { MailDraftService } from '../services/MailDraftService';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import type { ISendRequest, SendStepStatus } from '../models/ISend';
import { useSendDraft } from '../hooks/useSendDraft';
import { format } from '../utils/format';

export interface ISendDialogProps {
  files: GraphFileService;
  mail: MailDraftService;
  documents: ReadonlyArray<ISelectedDocument>;
  request: ISendRequest;
  webOrigin: string;
  locale: string;
  onClose: () => void;
}

const STATUS_TEXT: Record<SendStepStatus, string> = {
  waiting: strings.SendStatusWaiting,
  working: strings.SendStatusWorking,
  attached: strings.SendStatusAttached,
  linked: strings.SendStatusLinked,
  failed: strings.SendStatusFailed
};
const STATUS_COLOR: Record<SendStepStatus, 'informative' | 'brand' | 'success' | 'warning' | 'danger'> = {
  waiting: 'informative', working: 'brand', attached: 'success', linked: 'warning', failed: 'danger'
};

const useStyles = makeStyles({
  content: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM },
  versions: { maxHeight: '40vh', overflowY: 'auto' },
  steps: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, margin: 0, padding: 0, listStyleType: 'none' },
  step: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: tokens.spacingHorizontalS },
  stepName: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }
});

/** Creates an Outlook draft with the documents attached — current or a chosen previous version, native or PDF. */
export const SendDialog: React.FC<ISendDialogProps> = ({ files, mail, documents, request, webOrigin, locale, onClose }) => {
  const styles = useStyles();
  const { state, actions, computed } = useSendDraft(files, mail, documents, request, webOrigin);
  const dateFormat: Intl.DateTimeFormat = React.useMemo(
    () => new Intl.DateTimeFormat(locale, { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }), [locale]);

  const title: string = request.mode === 'version'
    ? (request.asPdf ? strings.SendVersionAs : strings.SendVersion)
    : (request.asPdf ? strings.SendAsPdf : strings.Send);

  return (
    <Dialog open modalType="modal" onOpenChange={(_, d) => { if (!d.open && !state.isRunning) { onClose(); } }}>
      <DialogSurface data-testid="send-dialog">
        <DialogBody>
          <DialogTitle>{title}</DialogTitle>
          <DialogContent className={styles.content}>
            {request.mode === 'version' && (
              <Field label={format(strings.SendPickVersion, documents[0]?.name ?? '')}>
                {state.versions === undefined
                  ? <Skeleton aria-label={strings.SendLoadingVersions}><SkeletonItem size={16} /><SkeletonItem size={16} /></Skeleton>
                  : (
                    <RadioGroup className={styles.versions} value={state.selectedVersionId ?? ''} disabled={state.isRunning || !!state.result}
                      onChange={(_, d) => actions.selectVersion(d.value)} data-testid="send-dialog-versions">
                      {state.versions.map(v => (
                        <Radio key={v.id} value={v.id} data-testid="send-dialog-version"
                          label={format(strings.SendVersionOption, v.label, dateFormat.format(new Date(v.modified)), v.modifiedBy)
                            + (v.isCurrent ? ` ${strings.SendCurrentVersion}` : '')} />
                      ))}
                    </RadioGroup>
                  )}
              </Field>
            )}
            <ul className={styles.steps} aria-label={strings.SendProgressLabel} data-testid="send-dialog-steps">
              {state.steps.map((s, i) => (
                <li key={i} className={styles.step} data-testid="send-dialog-step">
                  <Body1 className={styles.stepName} title={s.detail ?? s.name}>{s.name}</Body1>
                  <Badge appearance="tint" color={STATUS_COLOR[s.status]} title={s.detail}>{STATUS_TEXT[s.status]}</Badge>
                </li>
              ))}
            </ul>
            {state.result && (
              <MessageBar intent={state.result.failed ? 'warning' : 'success'} layout="multiline" data-testid="send-dialog-result">
                <MessageBarBody>
                  {format(strings.SendResult, state.result.attached, state.result.linked, state.result.failed)}
                  <Caption1 block>{strings.SendDesktopNote}</Caption1>
                </MessageBarBody>
              </MessageBar>
            )}
            {state.error && (
              <MessageBar intent="error" layout="multiline" data-testid="send-dialog-error"><MessageBarBody>{state.error}</MessageBarBody></MessageBar>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" disabled={state.isRunning} onClick={onClose} data-testid="send-dialog-close">
              {state.result ? strings.Close : strings.Cancel}
            </Button>
            {state.result
              ? (
                <Button appearance="primary" icon={<Open20Regular />} onClick={() => window.open(state.result?.webLink, '_blank', 'noopener')}
                  data-testid="send-dialog-open">{strings.SendOpenOutlook}</Button>
              )
              : (
                <Button appearance="primary" disabled={!computed.canRun} onClick={() => { actions.run().catch(() => undefined); }}
                  icon={state.isRunning ? <Spinner size="tiny" /> : undefined} data-testid="send-dialog-create">{strings.SendCreateDraft}</Button>
              )}
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
