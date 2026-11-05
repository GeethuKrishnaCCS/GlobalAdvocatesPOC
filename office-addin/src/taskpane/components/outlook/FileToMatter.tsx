import * as React from "react";
import { Body1, Link, MessageBar, MessageBarBody, Spinner, makeStyles, tokens } from "@fluentui/react-components";
import { useFileToMatter } from "../../../shared/hooks/useFileToMatter";
import { testIdPart } from "../../../shared/testId";
import { MatterStep } from "../saveToMatter/MatterStep";
import { FolderStep } from "../saveToMatter/FolderStep";
import { StepHeader, type IStep } from "../saveToMatter/StepHeader";
import { ItemsStep } from "./ItemsStep";
import { PaneFooter } from "./PaneFooter";

const STEPS: IStep[] = [
  { step: "matter", label: "1. Matter" },
  { step: "folder", label: "2. Folder" },
  { step: "items", label: "3. Items" },
];

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  list: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
});

/** File to Matter (Outlook read): files the open email (.eml) and its attachments into a matter folder. */
export const FileToMatter: React.FC = () => {
  const styles = useStyles();
  const { state, actions, computed } = useFileToMatter();
  const { browser } = state;
  const { location } = browser.state;
  const loading = browser.state.isLoading || state.isPrefilling;
  const busy = loading || state.isFiling;
  const error = state.error ?? browser.state.error;
  const onFolder = state.step === "folder";

  return (
    <div className={styles.root} data-testid="file-to-matter">
      <StepHeader step={state.step} steps={STEPS} />
      {error && (
        <MessageBar intent="error" layout="multiline" data-testid="file-to-matter-error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      {loading && <Spinner size="tiny" label="Loading…" data-testid="file-to-matter-loading" />}

      {state.step === "matter" && !state.isPrefilling && (
        <MatterStep
          matters={browser.state.matters}
          failed={!!browser.state.error}
          disabled={busy}
          onSelect={(m) => void actions.selectMatter(m)}
          onRetry={() => void browser.actions.loadMatters()}
        />
      )}
      {onFolder && location && (
        <FolderStep
          matter={location.matter}
          folderPath={location.folderPath}
          folders={browser.state.folders}
          disabled={busy}
          onOpen={(f) => void browser.actions.openFolder(f)}
          onOpenLevel={(depth) => void browser.actions.openPathLevel(depth)}
        />
      )}
      {state.step === "items" && (
        <ItemsStep rows={state.rows} ticked={state.ticked} disabled={state.isFiling} onToggle={actions.toggle} />
      )}
      {state.step === "done" && (
        <div className={styles.list} data-testid="file-to-matter-result">
          <MessageBar intent="success" data-testid="file-to-matter-success">
            <MessageBarBody>
              Filed {state.filed.length} item(s) to {browser.computed.locationLabel}. The email stays in your mailbox.
            </MessageBarBody>
          </MessageBar>
          {state.filed.map((f) => (
            <Body1 key={f.itemId} data-testid={`file-to-matter-filed-${testIdPart(f.name)}`}>
              <Link href={f.webUrl} target="_blank" data-testid={`file-to-matter-link-${testIdPart(f.name)}`}>
                {f.name}
              </Link>
            </Body1>
          ))}
        </div>
      )}

      {(onFolder || state.step === "items") && (
        <PaneFooter
          testId="file-to-matter"
          summary={onFolder ? undefined : `${computed.tickedCount} item(s) → ${browser.computed.locationLabel}`}
          primaryLabel={onFolder ? "File here" : "Submit"}
          busyLabel={onFolder ? "Loading…" : "Filing…"}
          busy={busy}
          canSubmit={onFolder ? !!location : computed.canSubmit}
          onBack={actions.back}
          onPrimary={onFolder ? actions.confirmFolder : () => void actions.submit()}
        />
      )}
    </div>
  );
};
