import * as React from "react";
import { Body1, Link, MessageBar, MessageBarBody, Spinner, Tab, TabList, makeStyles, tokens } from "@fluentui/react-components";
import { useFileToMatter } from "../../../shared/hooks/useFileToMatter";
import { useMailItem } from "../../../shared/hooks/useMailItem";
import { useMatterPins, type IMatterPins } from "../../../shared/hooks/useMatterPins";
import { testIdPart } from "../../../shared/testId";
import { MatterStep } from "../saveToMatter/MatterStep";
import { FolderStep } from "../saveToMatter/FolderStep";
import { StepHeader, type IStep } from "../saveToMatter/StepHeader";
import { ItemsStep } from "./ItemsStep";
import { MatterBrowser } from "./MatterBrowser";
import { PaneFooter } from "./PaneFooter";

const STEPS: IStep[] = [
  { step: "matter", label: "1. Matter" },
  { step: "folder", label: "2. Folder" },
  { step: "items", label: "3. Items" },
];

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  list: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
  hidden: { display: "none" },
});

type FileTab = "email" | "matters";

/**
 * File to Matter (Outlook read), two tabs: "This email" (wizard for the selected email) and "My matters" (drag emails
 * onto matters and folders, see filed emails). The pane can be pinned: selecting another email starts the wizard over
 * for it (the key remounts the body, so no state of the previous email survives); "My matters" keeps its place.
 * Both tabs stay mounted once shown, so switching tabs loses neither the wizard's progress nor a running drop.
 */
export const FileToMatter: React.FC = () => {
  const styles = useStyles();
  const { itemKey, hasItem } = useMailItem();
  // Opened with no single email (no item context): start on My matters; later selection changes never move the tab.
  const [tab, setTab] = React.useState<FileTab>(() => (hasItem ? "email" : "matters"));
  const [mattersShown, setMattersShown] = React.useState(() => !hasItem);
  // One instance for both tabs: a matter pinned in one is pinned in the other at once.
  const pins = useMatterPins();

  return (
    <div className={styles.root} data-testid="file-to-matter-tabs">
      <TabList
        size="small"
        selectedValue={tab}
        onTabSelect={(_, data) => {
          setTab(data.value as FileTab);
          if (data.value === "matters") setMattersShown(true);
        }}
        data-testid="file-to-matter-tablist"
      >
        <Tab value="email" data-testid="file-to-matter-tab-email">
          This email
        </Tab>
        <Tab value="matters" data-testid="file-to-matter-tab-matters">
          My matters
        </Tab>
      </TabList>
      {pins.state.error && (
        <MessageBar intent="error" layout="multiline" data-testid="file-to-matter-pin-error">
          <MessageBarBody>{pins.state.error}</MessageBarBody>
        </MessageBar>
      )}
      <div className={tab === "email" ? undefined : styles.hidden} data-testid="file-to-matter-panel-email">
        {hasItem ? (
          <FileToMatterBody key={itemKey} pins={pins} />
        ) : (
          <Body1 data-testid="file-to-matter-no-item">Select one email to file it here, or drag emails onto a matter in My matters.</Body1>
        )}
      </div>
      {mattersShown && (
        <div className={tab === "matters" ? undefined : styles.hidden} data-testid="file-to-matter-panel-matters">
          <MatterBrowser pins={pins} />
        </div>
      )}
    </div>
  );
};

/** The wizard for one email. */
const FileToMatterBody: React.FC<{ pins: IMatterPins }> = ({ pins }) => {
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
          pinned={pins.state.pinned}
          onTogglePin={(m) => pins.actions.toggle(m.ref)}
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
