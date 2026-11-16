import * as React from "react";
import { Button, Caption1, MessageBar, MessageBarBody, Spinner, makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import { ArrowLeft16Regular } from "@fluentui/react-icons";
import { useDropToFile, type DropTarget, type IDropDestination } from "../../../shared/hooks/useDropToFile";
import { useMatterFolders } from "../../../shared/hooks/useMatterFolders";
import { useMatterFolderStatus } from "../../../shared/hooks/useMatterFolderStatus";
import { folderStatusLine } from "../../../shared/services/matterFolderStatus";
import type { IMatterPins } from "../../../shared/hooks/useMatterPins";
import { DROP_KEY_ATTRIBUTE, folderDropKey, matterDropKey } from "../../../shared/services/mailDrop";
import { MatterStep } from "../saveToMatter/MatterStep";
import { FolderStep } from "../saveToMatter/FolderStep";
import { DropQueue } from "./DropQueue";
import { FiledEmailList } from "./FiledEmailList";
import { dropHoverStyle } from "./dropHoverStyle";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  hint: { color: tokens.colorNeutralForeground3 },
  folderError: { color: tokens.colorPaletteRedForeground1 },
  back: { alignSelf: "flex-start" },
  // The open folder is itself a drop target: padded so its dashed outline does not touch the rows.
  folder: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalXS,
  },
  dropHover: dropHoverStyle,
});

/**
 * "My matters" (Outlook read): browse matters and their folders, drag emails from the message list onto a matter or
 * folder to file them, and see the emails already filed in the open folder. Independent of the selected email, so a
 * pinned pane keeps its place while the user works through the inbox.
 */
export const MatterBrowser: React.FC<{ pins: IMatterPins }> = ({ pins }) => {
  const styles = useStyles();
  const browser = useMatterFolders();
  const { matters, location, folders, files, isLoading, error } = browser.state;
  const { loadMatters, selectMatter, openFolder, openPathLevel, clearMatter } = browser.actions;
  const { folderId, locationLabel } = browser.computed;

  React.useEffect(() => {
    void loadMatters();
  }, [loadMatters]);

  const { syncPinned } = pins.actions;
  React.useEffect(() => {
    if (matters) syncPinned(matters);
  }, [matters, syncPinned]);

  const folderStatus = useMatterFolderStatus(pins.state.pinned.size > 0, pins.state.withMailFolder).state;
  const folderLine = folderStatusLine(
    folderStatus.status,
    pins.state.pinned,
    folderStatus.checkedAt,
    pins.state.folderError ?? folderStatus.error
  );

  const resolveTarget = React.useCallback(
    (key: string): DropTarget | undefined => {
      if (!location) {
        const matter = matters?.find((m) => matterDropKey(m.siteId) === key);
        return matter && { kind: "matter", matter };
      }
      const { driveId } = location;
      if (key === folderDropKey(folderId)) return { kind: "folder", driveId, folderId, label: locationLabel };
      const folder = folders?.find((f) => folderDropKey(f.id) === key);
      return folder && { kind: "folder", driveId, folderId: folder.id, label: `${locationLabel} › ${folder.name}` };
    },
    [folderId, folders, location, locationLabel, matters]
  );

  /** Filed into the folder on screen: reload it so the new emails show in "Filed emails". */
  const onFiled = React.useCallback(
    (destination: IDropDestination) => {
      if (location && destination.driveId === location.driveId && destination.folderId === folderId)
        void openPathLevel(location.folderPath.length);
    },
    [folderId, location, openPathLevel]
  );

  const drop = useDropToFile(resolveTarget, onFiled);
  const { hoverKey, notice, items } = drop.state;

  return (
    <div className={styles.root} data-testid="matter-browser">
      <Caption1 className={styles.hint} data-testid="matter-browser-hint">
        Drag emails from your message list onto a matter or folder to file them. Pin this pane to keep it open while you
        work through your inbox.
      </Caption1>
      {folderLine && (
        <Caption1
          className={mergeClasses(styles.hint, folderLine.isError && styles.folderError)}
          data-testid="matter-folders-status"
        >
          {folderLine.text}
          {folderLine.isError && (
            <Button
              appearance="transparent"
              size="small"
              onClick={pins.actions.retry}
              data-testid="matter-folders-retry"
            >
              Retry
            </Button>
          )}
        </Caption1>
      )}
      {error && (
        <MessageBar intent="error" layout="multiline" data-testid="matter-browser-error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      {notice && (
        <MessageBar intent="warning" layout="multiline" data-testid="matter-browser-notice">
          <MessageBarBody>{notice}</MessageBarBody>
        </MessageBar>
      )}
      <DropQueue
        items={items}
        filedCount={drop.computed.filedCount}
        failedCount={drop.computed.failedCount}
        isFiling={drop.computed.isFiling}
        onRetry={(id) => void drop.actions.retry(id)}
        onClearFinished={drop.actions.clearFinished}
      />
      {isLoading && <Spinner size="tiny" label="Loading…" data-testid="matter-browser-loading" />}

      {!location && (
        <MatterStep
          matters={matters}
          failed={!!error}
          disabled={isLoading}
          onSelect={(m) => void selectMatter(m)}
          onRetry={() => void loadMatters()}
          droppable
          hoverKey={hoverKey}
          pinned={pins.state.pinned}
          onTogglePin={(m) => pins.actions.toggle(m)}
          withMailFolder={pins.state.withMailFolder}
        />
      )}
      {location && (
        <>
          <Button
            appearance="subtle"
            size="small"
            icon={<ArrowLeft16Regular />}
            className={styles.back}
            onClick={clearMatter}
            data-testid="matter-browser-back"
          >
            All matters
          </Button>
          <div
            className={mergeClasses(styles.folder, hoverKey === folderDropKey(folderId) && styles.dropHover)}
            {...{ [DROP_KEY_ATTRIBUTE]: folderDropKey(folderId) }}
            data-testid="matter-browser-folder"
          >
            <FolderStep
              matter={location.matter}
              folderPath={location.folderPath}
              folders={folders}
              disabled={isLoading}
              onOpen={(f) => void openFolder(f)}
              onOpenLevel={(depth) => void openPathLevel(depth)}
              emptyText="No sub-folders."
              droppable
              hoverKey={hoverKey}
            />
            {files && <FiledEmailList files={files} />}
          </div>
        </>
      )}
    </div>
  );
};
