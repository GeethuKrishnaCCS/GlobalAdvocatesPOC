import * as React from "react";
import { Body1, Caption1, MessageBar, MessageBarBody, Spinner, makeStyles, tokens } from "@fluentui/react-components";
import { Document20Regular } from "@fluentui/react-icons";
import { useSaveToMatter } from "../../../shared/hooks/useSaveToMatter";
import { StepHeader } from "./StepHeader";
import { MatterStep } from "./MatterStep";
import { FolderStep } from "./FolderStep";
import { DetailsStep } from "./DetailsStep";
import { SaveResult } from "./SaveResult";
import { SaveFooter } from "./SaveFooter";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  file: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS },
  fileText: { display: "flex", flexDirection: "column", minWidth: 0, overflowWrap: "anywhere" },
  body: { outlineStyle: "none" },
});

/** Save to Matter: saves the open document (and optionally a PDF copy) into a matter folder. */
export const SaveToMatter: React.FC<{ host: Office.HostType }> = ({ host }) => {
  const styles = useStyles();
  const { state, actions, computed } = useSaveToMatter(host);
  const busy = state.isLoading || state.isSaving;
  const { document } = state;

  // The button that moved the wizard on disappears with its step; keep keyboard focus in the new step.
  const bodyRef = React.useRef<HTMLDivElement>(null);
  const firstStep = React.useRef(true);
  React.useEffect(() => {
    if (firstStep.current) {
      firstStep.current = false;
      return;
    }
    bodyRef.current?.focus();
  }, [state.step]);

  return (
    <div className={styles.root} data-testid="save-to-matter">
      <div className={styles.file} data-testid="save-to-matter-file">
        <Document20Regular data-testid="save-to-matter-file-icon" />
        <span className={styles.fileText}>
          <Body1 data-testid="save-to-matter-file-name">{document.baseName + document.extension}</Body1>
          <Caption1 data-testid="save-to-matter-file-location">
            {document.url ? ["Open document", document.folderName].filter(Boolean).join(" · ") : "Not saved yet"}
          </Caption1>
        </span>
      </div>
      <StepHeader step={state.step} />
      {state.error && (
        <MessageBar intent="error" layout="multiline" data-testid="save-to-matter-error">
          <MessageBarBody>{state.error}</MessageBarBody>
        </MessageBar>
      )}
      {state.isLoading && <Spinner size="tiny" label="Loading…" data-testid="save-to-matter-loading" />}

      <div ref={bodyRef} tabIndex={-1} className={styles.body} aria-live="polite" data-testid="save-to-matter-body">
        {state.step === "matter" && (
          <MatterStep
            matters={state.matters}
            failed={!!state.error}
            disabled={busy}
            onSelect={(m) => void actions.selectMatter(m)}
            onRetry={() => void actions.loadMatters()}
          />
        )}
        {state.step === "folder" && state.matter && (
          <FolderStep
            matter={state.matter}
            folderPath={state.folderPath}
            folders={state.folders}
            disabled={busy}
            onOpen={(f) => void actions.openFolder(f)}
            onOpenLevel={(depth) => void actions.openPathLevel(depth)}
          />
        )}
        {state.step === "details" && (
          <DetailsStep
            locationLabel={computed.locationLabel}
            matterTitle={state.matter?.title ?? ""}
            extension={document.extension}
            details={state.details}
            fileNameError={computed.fileNameError}
            disabled={state.isSaving}
            onChange={actions.updateDetails}
            onChangeLocation={actions.back}
          />
        )}
        {state.step === "done" && state.result && (
          <SaveResult result={state.result} locationLabel={computed.locationLabel} onSaveAnother={actions.saveAnother} />
        )}
      </div>

      {(state.step === "folder" || state.step === "details") && (
        <SaveFooter
          mode={state.step}
          summary={`${computed.fileName} → ${computed.locationLabel}`}
          busy={busy}
          isSaving={state.isSaving}
          canSubmit={computed.canSubmit}
          onBack={actions.back}
          onConfirmFolder={actions.confirmFolder}
          onSubmit={() => void actions.submit()}
        />
      )}
    </div>
  );
};
