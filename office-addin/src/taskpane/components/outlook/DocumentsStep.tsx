import * as React from "react";
import { Body1, Caption1, Checkbox, SearchBox, makeStyles, tokens } from "@fluentui/react-components";
import type { IMatterFile } from "../../../shared/services/matterService";
import { testIdPart } from "../../../shared/testId";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
  row: { display: "flex", flexDirection: "column" },
  note: { paddingLeft: tokens.spacingHorizontalXXL, color: tokens.colorNeutralForeground3 },
});

export interface IDocumentsStepProps {
  query: string;
  searching: boolean;
  files: IMatterFile[];
  ticked: ReadonlySet<string>;
  disabled: boolean;
  isAttached: (file: IMatterFile) => boolean;
  onQuery: (query: string) => void;
  onSearch: () => void;
  onToggle: (file: IMatterFile, on: boolean) => void;
}

const KB = 1024;

function note(file: IMatterFile, attached: boolean): string {
  return [attached ? "Attached" : undefined, file.docId ?? "No Doc ID", `${Math.max(1, Math.round(file.size / KB))} KB`]
    .filter(Boolean)
    .join(" · ");
}

/** Documents of the open folder (or search hits) with their Doc ID; already attached ones are marked and locked. */
export const DocumentsStep: React.FC<IDocumentsStepProps> = (props) => {
  const { query, searching, files, ticked, disabled, isAttached, onQuery, onSearch, onToggle } = props;
  const styles = useStyles();
  return (
    <div className={styles.root} data-testid="documents-step">
      <SearchBox
        placeholder="Search documents in this matter"
        aria-label="Search documents in this matter"
        value={query}
        onChange={(_, data) => onQuery(data.value)}
        onKeyDown={(e) => e.key === "Enter" && onSearch()}
        data-testid="documents-step-search"
      />
      {!files.length && (
        <Body1 data-testid="documents-step-empty">
          {searching ? "No document matches your search." : "No documents in this folder."}
        </Body1>
      )}
      <div role="list" aria-label="Documents" className={styles.root} data-testid="documents-step-list">
        {files.map((f) => {
          const attached = isAttached(f);
          const id = testIdPart(f.name);
          return (
            <div key={f.id} role="listitem" className={styles.row} data-testid={`documents-step-row-${id}`}>
              <Checkbox
                label={f.name}
                checked={attached || ticked.has(f.id)}
                disabled={disabled || attached}
                onChange={(_, data) => onToggle(f, !!data.checked)}
                data-testid={`documents-step-check-${id}`}
              />
              <Caption1 className={styles.note} data-testid={`documents-step-note-${id}`}>
                {note(f, attached)}
              </Caption1>
            </div>
          );
        })}
      </div>
    </div>
  );
};
