import * as React from "react";
import { Body1, Button, Caption1, SearchBox, makeStyles, tokens } from "@fluentui/react-components";
import { Briefcase20Regular } from "@fluentui/react-icons";
import type { IMatter } from "../../../shared/services/matterService";
import { testIdPart } from "../../../shared/testId";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  list: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
  item: { justifyContent: "flex-start", textAlign: "start", width: "100%" },
  text: { display: "flex", flexDirection: "column" },
});

export interface IMatterStepProps {
  matters: IMatter[] | undefined;
  /** Loading the list failed: offer Retry instead of an empty list. */
  failed: boolean;
  disabled: boolean;
  onSelect: (matter: IMatter) => void;
  onRetry: () => void;
}

/** Step 1: pick the matter (searchable by reference or title). */
export const MatterStep: React.FC<IMatterStepProps> = ({ matters, failed, disabled, onSelect, onRetry }) => {
  const styles = useStyles();
  const [query, setQuery] = React.useState("");
  const shown = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return (matters ?? []).filter((m) => !q || `${m.ref} ${m.title}`.toLowerCase().includes(q));
  }, [matters, query]);

  return (
    <div className={styles.root} data-testid="matter-step">
      <SearchBox
        placeholder="Search matters"
        value={query}
        onChange={(_, data) => setQuery(data.value)}
        aria-label="Search matters"
        data-testid="matter-step-search"
      />
      {failed && !matters && (
        <Button disabled={disabled} onClick={onRetry} data-testid="matter-step-retry">
          Retry
        </Button>
      )}
      {matters && shown.length === 0 && (
        <Body1 data-testid="matter-step-empty">
          {matters.length ? "No matter matches your search." : "You have no matters you can save to."}
        </Body1>
      )}
      <div className={styles.list} role="list" aria-label="Matters" data-testid="matter-step-list">
        {shown.map((m) => (
          <div key={m.siteId} role="listitem" data-testid={`matter-step-row-${testIdPart(m.ref)}`}>
            <Button
              appearance="subtle"
              icon={<Briefcase20Regular />}
              className={styles.item}
              disabled={disabled}
              onClick={() => onSelect(m)}
              data-testid={`matter-step-item-${testIdPart(m.ref)}`}
            >
              <span className={styles.text}>
                <Body1 data-testid={`matter-step-ref-${testIdPart(m.ref)}`}>{m.ref}</Body1>
                <Caption1 data-testid={`matter-step-title-${testIdPart(m.ref)}`}>{m.title}</Caption1>
              </span>
            </Button>
          </div>
        ))}
      </div>
    </div>
  );
};
