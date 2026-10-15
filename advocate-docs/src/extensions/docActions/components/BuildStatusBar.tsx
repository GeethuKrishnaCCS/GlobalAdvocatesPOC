import * as React from 'react';
import { Link, MessageBar, MessageBarBody, ProgressBar, makeStyles, tokens } from '@fluentui/react-components';
import * as strings from 'DocActionsCommandSetStrings';
import type { IBundleBuild } from '../hooks/useBundleBuild';
import { format } from '../utils/format';

export interface IBuildStatusBarProps {
  build: IBundleBuild;
}

const STAGE_TEXT: Record<string, string> = {
  loading: strings.BuildStageLoading,
  converting: strings.BuildStageConverting,
  assembling: strings.BuildStageAssembling,
  publishing: strings.BuildStagePublishing
};

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS }
});

/** Progress while a bundle builds; link to the PDF when done; the error if it failed. */
export const BuildStatusBar: React.FC<IBuildStatusBarProps> = ({ build }) => {
  const styles = useStyles();
  const { status, isRunning, error } = build.state;

  if (error) {
    return (
      <MessageBar intent="error" layout="multiline" data-testid="build-status-error">
        <MessageBarBody>{format(strings.BuildFailed, error)}</MessageBarBody>
      </MessageBar>
    );
  }
  if (status?.status === 'Completed' && status.result) {
    const r = status.result;
    return (
      <MessageBar intent={r.placeholders ? 'warning' : 'success'} layout="multiline" data-testid="build-status-done">
        <MessageBarBody>
          {format(strings.BuildDone, r.version, r.pageCount, r.documents, r.placeholders)}{' '}
          <Link href={r.pdfUrl} target="_blank" rel="noopener" data-testid="build-status-open">{strings.BuildOpenPdf}</Link>
        </MessageBarBody>
      </MessageBar>
    );
  }
  if (!isRunning) {
    return null;
  }
  const p = status?.progress;
  const label: string = p && p.total
    ? format(strings.BuildProgress, STAGE_TEXT[p.stage] ?? p.stage, p.done, p.total)
    : strings.BuildStarting;
  return (
    <div className={styles.root} data-testid="build-status-running">
      <MessageBar intent="info"><MessageBarBody>{label}</MessageBarBody></MessageBar>
      <ProgressBar value={p && p.total ? p.done / p.total : undefined} />
    </div>
  );
};
