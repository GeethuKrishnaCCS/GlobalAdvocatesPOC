import * as React from "react";
import { Body1, Button, Caption1, SearchBox, makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import { Briefcase20Regular, Pin20Filled, Pin20Regular } from "@fluentui/react-icons";
import type { IMatter } from "../../../shared/services/matterService";
import { DROP_KEY_ATTRIBUTE, matterDropKey } from "../../../shared/services/mailDrop";
import { pinnedFirst } from "../../../shared/services/matterPins";
import { testIdPart } from "../../../shared/testId";
import { dropHoverStyle } from "../outlook/dropHoverStyle";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  list: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
  row: { display: "flex", alignItems: "center" },
  item: { justifyContent: "flex-start", textAlign: "start", flexGrow: 1, minWidth: 0 },
  pin: { flexShrink: 0 },
  pinned: { color: tokens.colorBrandForeground1 },
  text: { display: "flex", flexDirection: "column" },
  dropHover: dropHoverStyle,
});

export interface IMatterStepProps {
  matters: IMatter[] | undefined;
  /** Loading the list failed: offer Retry instead of an empty list. */
  failed: boolean;
  disabled: boolean;
  onSelect: (matter: IMatter) => void;
  onRetry: () => void;
  /** Outlook "My matters": rows accept dropped emails; `hoverKey` is the drop target under the pointer. */
  droppable?: boolean;
  hoverKey?: string;
  /** Outlook panes: references of the pinned matters (listed first) and the pin / unpin button of each row. */
  pinned?: ReadonlySet<string>;
  onTogglePin?: (matter: IMatter) => void;
}

const NO_PINS: ReadonlySet<string> = new Set();

/** Step 1: pick the matter (searchable by reference or title). */
export const MatterStep: React.FC<IMatterStepProps> = (props) => {
  const { matters, failed, disabled, onSelect, onRetry, droppable, hoverKey, pinned = NO_PINS, onTogglePin } = props;
  const styles = useStyles();
  const [query, setQuery] = React.useState("");
  const shown = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    const found = (matters ?? []).filter((m) => !q || `${m.ref} ${m.title}`.toLowerCase().includes(q));
    return pinnedFirst(found, pinned);
  }, [matters, pinned, query]);

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
          <div
            key={m.siteId}
            role="listitem"
            className={mergeClasses(styles.row, droppable && hoverKey === matterDropKey(m.siteId) && styles.dropHover)}
            {...(droppable ? { [DROP_KEY_ATTRIBUTE]: matterDropKey(m.siteId) } : {})}
            data-testid={`matter-step-row-${testIdPart(m.ref)}`}
          >
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
            {onTogglePin && (
              <Button
                appearance="subtle"
                className={styles.pin}
                icon={pinned.has(m.ref) ? <Pin20Filled className={styles.pinned} /> : <Pin20Regular />}
                aria-pressed={pinned.has(m.ref)}
                aria-label={`${pinned.has(m.ref) ? "Unpin" : "Pin"} ${m.ref}`}
                title={pinned.has(m.ref) ? "Unpin: stop listing this matter first" : "Pin: list this matter first"}
                onClick={() => onTogglePin(m)}
                data-testid={`matter-step-pin-${testIdPart(m.ref)}`}
              />
            )}
          </div>
        ))}
      </div>
    </div>
  );
};
