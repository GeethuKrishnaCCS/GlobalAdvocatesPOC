import * as React from "react";
import { Body1, Button, Caption1, MessageBar, MessageBarBody, Spinner, makeStyles, tokens } from "@fluentui/react-components";
import { PaneView } from "../../shared/config";
import { isNaaSupported } from "../../shared/auth/naa";
import { useSignIn } from "../../shared/hooks/useSignIn";

export interface IAppProps {
  host: Office.HostType;
  view: PaneView | undefined;
}

const VIEW_TITLES: Record<PaneView, string> = {
  [PaneView.SaveToMatter]: "Save to Matter",
  [PaneView.SendAsAttachment]: "Send as Attachment",
};

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, padding: tokens.spacingHorizontalL },
});

/** Step 0 spike: proves view routing + NAA sign-in + Graph /me in Word. */
const App: React.FC<IAppProps> = ({ host, view }) => {
  const styles = useStyles();
  const { state, actions, computed } = useSignIn();

  return (
    <div className={styles.root} data-testid="app-root">
      <Body1 data-testid="app-title">{view ? VIEW_TITLES[view] : "Global Advocate"}</Body1>
      <Caption1 data-testid="app-host">
        Host: {host} · NAA: {isNaaSupported() ? "yes" : "no"}
      </Caption1>
      {state.isLoading && <Spinner size="tiny" label="Signing in…" data-testid="app-signin-spinner" />}
      {state.user && (
        <Body1 data-testid="app-user">
          Signed in as {state.user.displayName} ({state.user.userPrincipalName})
        </Body1>
      )}
      {state.error && (
        <MessageBar intent="error" data-testid="app-error">
          <MessageBarBody>{state.error}</MessageBarBody>
        </MessageBar>
      )}
      {computed.needsSignIn && (
        <Button appearance="primary" onClick={() => void actions.signIn()} data-testid="app-signin-button">
          Sign in
        </Button>
      )}
    </div>
  );
};

export default App;
