import * as React from "react";
import { MessageBar, MessageBarBody, Spinner, makeStyles, tokens } from "@fluentui/react-components";
import { useAttachFromMatter } from "../../../shared/hooks/useAttachFromMatter";
import { MatterStep } from "../saveToMatter/MatterStep";
import { FolderStep } from "../saveToMatter/FolderStep";
import { StepHeader, type IStep } from "../saveToMatter/StepHeader";
import { DocumentsStep } from "./DocumentsStep";
import { PaneFooter } from "./PaneFooter";

const STEPS: IStep[] = [
  { step: "matter", label: "1. Matter" },
  { step: "documents", label: "2. Documents" },
];

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
});

/** Attach from Matter (Outlook compose): attaches matter documents to the email being written. */
export const AttachFromMatter: React.FC = () => {
  const styles = useStyles();
  const { state, actions, computed } = useAttachFromMatter();
  const { browser } = state;
  const { location } = browser.state;
  const loading = browser.state.isLoading || state.isSearching;
  const busy = loading || state.isAttaching;
  const error = state.error ?? browser.state.error;

  return (
    <div className={styles.root} data-testid="attach-from-matter">
      <StepHeader step={state.step} steps={STEPS} />
      {error && (
        <MessageBar intent="error" layout="multiline" data-testid="attach-from-matter-error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      {state.attachedNow.length > 0 && (
        <MessageBar intent="success" data-testid="attach-from-matter-success">
          <MessageBarBody>Attached: {state.attachedNow.join(", ")}</MessageBarBody>
        </MessageBar>
      )}
      {loading && <Spinner size="tiny" label="Loading…" data-testid="attach-from-matter-loading" />}

      {state.step === "matter" && (
        <MatterStep
          matters={browser.state.matters}
          failed={!!browser.state.error}
          disabled={busy}
          onSelect={(m) => void actions.selectMatter(m)}
          onRetry={() => void browser.actions.loadMatters()}
        />
      )}
      {state.step === "documents" && location && (
        <>
          {!state.results && (
            <FolderStep
              matter={location.matter}
              folderPath={location.folderPath}
              folders={browser.state.folders}
              disabled={busy}
              onOpen={(f) => void browser.actions.openFolder(f)}
              onOpenLevel={(depth) => void browser.actions.openPathLevel(depth)}
            />
          )}
          <DocumentsStep
            query={state.query}
            searching={!!state.results}
            files={computed.shownFiles}
            ticked={state.ticked}
            disabled={busy}
            isAttached={computed.isAttached}
            onQuery={actions.setQuery}
            onSearch={() => void actions.search()}
            onToggle={actions.toggle}
          />
          <PaneFooter
            testId="attach-from-matter"
            summary={`${state.ticked.size} selected`}
            primaryLabel="Attach to email"
            busyLabel="Attaching…"
            busy={busy}
            canSubmit={computed.canAttach}
            onBack={actions.back}
            onPrimary={() => void actions.attach()}
          />
        </>
      )}
    </div>
  );
};
