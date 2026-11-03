import * as React from "react";
import {
  Body1,
  Breadcrumb,
  BreadcrumbButton,
  BreadcrumbDivider,
  BreadcrumbItem,
  Button,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { Folder20Regular } from "@fluentui/react-icons";
import type { IMatter, IMatterFolder } from "../../../shared/services/matterService";
import { testIdPart } from "../../../shared/testId";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  list: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
  item: { justifyContent: "flex-start", width: "100%" },
});

export interface IFolderStepProps {
  matter: IMatter;
  folderPath: IMatterFolder[];
  folders: IMatterFolder[] | undefined;
  disabled: boolean;
  onOpen: (folder: IMatterFolder) => void;
  onOpenLevel: (depth: number) => void;
}

/** Step 2: browse the matter library; the open folder is the save target ("Save here" confirms it). */
export const FolderStep: React.FC<IFolderStepProps> = ({ matter, folderPath, folders, disabled, onOpen, onOpenLevel }) => {
  const styles = useStyles();
  const levels = [matter.ref, ...folderPath.map((f) => f.name)];
  const currentDepth = levels.length - 1;

  return (
    <div className={styles.root} data-testid="folder-step">
      <Breadcrumb aria-label="Folder path" size="small" data-testid="folder-step-path">
        {levels.map((name, depth) => (
          <React.Fragment key={`${depth}-${name}`}>
            {depth > 0 && <BreadcrumbDivider data-testid={`folder-step-path-divider-${depth}`} />}
            <BreadcrumbItem data-testid={`folder-step-path-item-${depth}`}>
              <BreadcrumbButton
                current={depth === currentDepth}
                disabled={disabled}
                // The current level is already open; only ancestors navigate.
                onClick={depth === currentDepth ? undefined : () => onOpenLevel(depth)}
                data-testid={`folder-step-path-${depth}`}
              >
                {name}
              </BreadcrumbButton>
            </BreadcrumbItem>
          </React.Fragment>
        ))}
      </Breadcrumb>
      {folders?.length === 0 && (
        <Body1 data-testid="folder-step-empty">No sub-folders. Save here, or go back up.</Body1>
      )}
      <div className={styles.list} role="list" aria-label="Folders" data-testid="folder-step-list">
        {folders?.map((f) => (
          <div key={f.id} role="listitem" data-testid={`folder-step-row-${testIdPart(f.name)}`}>
            <Button
              appearance="subtle"
              icon={<Folder20Regular />}
              className={styles.item}
              disabled={disabled}
              onClick={() => onOpen(f)}
              data-testid={`folder-step-item-${testIdPart(f.name)}`}
            >
              {f.name}
            </Button>
          </div>
        ))}
      </div>
    </div>
  );
};
