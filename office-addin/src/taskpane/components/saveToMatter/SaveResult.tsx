import * as React from "react";
import { Body1, Button, Caption1, Link, makeStyles, tokens } from "@fluentui/react-components";
import { CheckmarkCircle24Filled } from "@fluentui/react-icons";
import type { ISaveResult } from "../../../shared/hooks/useSaveToMatter";
import type { ISavedFile } from "../../../shared/services/saveToMatterService";

/* global Office */

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  header: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS },
  icon: { color: tokens.colorPaletteGreenForeground1 },
  file: { display: "flex", flexDirection: "column" },
});

/** Desktop task panes would open target=_blank inside the add-in's own webview; use the system browser there. */
function openInBrowser(event: React.MouseEvent, url: string): void {
  if (!Office.context.ui?.openBrowserWindow) return;
  event.preventDefault();
  Office.context.ui.openBrowserWindow(url);
}

const SavedFile: React.FC<{ file: ISavedFile; testId: string }> = ({ file, testId }) => {
  const styles = useStyles();
  return (
    <div className={styles.file} data-testid={testId}>
      <Link
        href={file.webUrl}
        target="_blank"
        rel="noreferrer"
        onClick={(e) => openInBrowser(e, file.webUrl)}
        data-testid={`${testId}-link`}
      >
        {file.name}
      </Link>
      <Caption1 data-testid={`${testId}-version`}>Version {file.version}</Caption1>
    </div>
  );
};

export interface ISaveResultProps {
  result: ISaveResult;
  locationLabel: string;
  onSaveAnother: () => void;
}

/** Done: what was saved where, with links and version labels. */
export const SaveResult: React.FC<ISaveResultProps> = ({ result, locationLabel, onSaveAnother }) => {
  const styles = useStyles();
  return (
    <div className={styles.root} data-testid="save-result">
      <div className={styles.header}>
        <CheckmarkCircle24Filled className={styles.icon} data-testid="save-result-icon" />
        <Body1 data-testid="save-result-message">Saved to {locationLabel}</Body1>
      </div>
      <SavedFile file={result.document} testId="save-result-document" />
      {result.pdf && <SavedFile file={result.pdf} testId="save-result-pdf" />}
      <Button onClick={onSaveAnother} data-testid="save-result-save-another">
        Save again
      </Button>
    </div>
  );
};
