import * as React from "react";
import { Button, Caption1, Link, Spinner, Subtitle2, makeStyles, tokens } from "@fluentui/react-components";
import { CheckmarkCircle20Regular, ErrorCircle20Regular } from "@fluentui/react-icons";
import type { IDropItem } from "../../../shared/hooks/useDropToFile";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    padding: tokens.spacingHorizontalS,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
  },
  header: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: tokens.spacingHorizontalS },
  row: { display: "flex", alignItems: "flex-start", gap: tokens.spacingHorizontalS },
  icon: { flexShrink: 0 },
  filed: { color: tokens.colorStatusSuccessForeground1 },
  failed: { color: tokens.colorStatusDangerForeground1 },
  text: { display: "flex", flexDirection: "column", minWidth: 0, flexGrow: 1, overflowWrap: "anywhere" },
  note: { color: tokens.colorNeutralForeground3 },
  error: { color: tokens.colorStatusDangerForeground1 },
});

export interface IDropQueueProps {
  items: IDropItem[];
  filedCount: number;
  failedCount: number;
  isFiling: boolean;
  onRetry: (id: number) => void;
  onClearFinished: () => void;
}

function summary({ items, filedCount, failedCount, isFiling }: IDropQueueProps): string {
  if (isFiling) return `Filing ${items.length - filedCount - failedCount} of ${items.length}…`;
  if (failedCount) return `${filedCount} filed, ${failedCount} failed`;
  return `${filedCount} filed. The emails stay in your mailbox.`;
}

/** Progress and result of dropped emails, newest first; failed ones offer Retry. */
export const DropQueue: React.FC<IDropQueueProps> = (props) => {
  const { items, filedCount, isFiling, onRetry, onClearFinished } = props;
  const styles = useStyles();
  if (!items.length) return null;

  return (
    <div className={styles.root} data-testid="drop-queue">
      <div className={styles.header}>
        <Subtitle2 as="h2" role="status" data-testid="drop-queue-summary">
          {summary(props)}
        </Subtitle2>
        {filedCount > 0 && !isFiling && (
          <Button size="small" appearance="subtle" onClick={onClearFinished} data-testid="drop-queue-clear">
            Clear
          </Button>
        )}
      </div>
      <div className={styles.root} role="list" aria-label="Dropped emails" data-testid="drop-queue-list">
        {items.map((item) => (
          <div key={item.id} role="listitem" className={styles.row} data-testid={`drop-queue-row-${item.id}`}>
            <span className={styles.icon} data-testid={`drop-queue-status-${item.id}`} data-status={item.status}>
              {item.status === "filed" && <CheckmarkCircle20Regular className={styles.filed} aria-label="Filed" />}
              {item.status === "failed" && <ErrorCircle20Regular className={styles.failed} aria-label="Failed" />}
              {(item.status === "filing" || item.status === "waiting") && (
                <Spinner size="tiny" aria-label={item.status === "filing" ? "Filing" : "Waiting"} />
              )}
            </span>
            <span className={styles.text}>
              {item.webUrl ? (
                <Link href={item.webUrl} target="_blank" rel="noreferrer" data-testid={`drop-queue-link-${item.id}`}>
                  {item.name}
                </Link>
              ) : (
                <span data-testid={`drop-queue-name-${item.id}`}>{item.name}</span>
              )}
              <Caption1 className={styles.note} data-testid={`drop-queue-target-${item.id}`}>
                → {item.target}
              </Caption1>
              {item.error && (
                <Caption1 className={styles.error} data-testid={`drop-queue-error-${item.id}`}>
                  {item.error}
                </Caption1>
              )}
            </span>
            {item.status === "failed" && (
              <Button size="small" onClick={() => onRetry(item.id)} data-testid={`drop-queue-retry-${item.id}`}>
                Retry
              </Button>
            )}
          </div>
        ))}
      </div>
    </div>
  );
};
