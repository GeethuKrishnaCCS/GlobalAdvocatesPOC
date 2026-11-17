import * as React from "react";
import { Body1, Caption1, Link, Subtitle2, makeStyles, tokens } from "@fluentui/react-components";
import { Mail20Regular } from "@fluentui/react-icons";
import { isEmailFile } from "../../../shared/services/mailDrop";
import type { IMatterFile } from "../../../shared/services/matterService";
import { testIdPart } from "../../../shared/testId";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
  row: { display: "flex", alignItems: "flex-start", gap: tokens.spacingHorizontalS },
  icon: { flexShrink: 0, color: tokens.colorNeutralForeground3 },
  text: { display: "flex", flexDirection: "column", minWidth: 0 },
  // Long subjects wrap instead of widening the narrow pane.
  name: { overflowWrap: "anywhere" },
  note: { color: tokens.colorNeutralForeground3 },
});

const KB = 1024;

/** Emails (.eml / .msg) already filed in the open matter folder, by anyone; each opens from SharePoint in a new tab. */
export const FiledEmailList: React.FC<{ files: IMatterFile[] }> = ({ files }) => {
  const styles = useStyles();
  const emails = React.useMemo(() => files.filter((f) => isEmailFile(f.name)), [files]);

  return (
    <div className={styles.root} data-testid="filed-emails">
      <Subtitle2 as="h2" data-testid="filed-emails-heading">
        Filed emails ({emails.length})
      </Subtitle2>
      {emails.length === 0 && (
        <Body1 data-testid="filed-emails-empty">No emails in this folder yet. Drag emails here to file them.</Body1>
      )}
      <div className={styles.root} role="list" aria-label="Filed emails" data-testid="filed-emails-list">
        {emails.map((f) => (
          <div key={f.id} role="listitem" className={styles.row} data-testid={`filed-emails-row-${testIdPart(f.name)}`}>
            <Mail20Regular className={styles.icon} aria-hidden data-testid={`filed-emails-icon-${testIdPart(f.name)}`} />
            <span className={styles.text}>
              <Link
                href={f.webUrl}
                target="_blank"
                rel="noreferrer"
                className={styles.name}
                data-testid={`filed-emails-link-${testIdPart(f.name)}`}
              >
                {f.name}
              </Link>
              <Caption1 className={styles.note} data-testid={`filed-emails-size-${testIdPart(f.name)}`}>
                {Math.max(1, Math.round(f.size / KB))} KB
              </Caption1>
            </span>
          </div>
        ))}
      </div>
    </div>
  );
};
