import * as React from "react";
import { getGraphToken } from "../auth/naa";
import { friendlyAuthError } from "../auth/authErrors";
import { graphClient } from "../services/graphClient";

/* global setTimeout, atob, TextDecoder */

/** A silent sign-in that has not settled by then is treated as failed, so the Sign in button shows. */
const SIGN_IN_TIMEOUT_MS = 20000;

function withTimeout<T>(promise: Promise<T>, ms: number, message: string): Promise<T> {
  return Promise.race([
    promise,
    new Promise<T>((_, reject) => setTimeout(() => reject(new Error(message)), ms)),
  ]);
}

interface IGraphUser {
  displayName: string;
  userPrincipalName: string;
}

/** Signed-in user from the Graph access token's claims (no extra Graph permission needed). */
function userFromToken(accessToken: string): IGraphUser {
  const payload = accessToken.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
  // atob yields Latin-1 bytes; decode them as UTF-8 so accented names survive.
  const json = new TextDecoder().decode(Uint8Array.from(atob(payload), (c) => c.charCodeAt(0)));
  const claims = JSON.parse(json) as {
    name?: string;
    upn?: string;
    preferred_username?: string;
  };
  return {
    displayName: claims.name ?? "",
    userPrincipalName: claims.upn ?? claims.preferred_username ?? "",
  };
}

export interface ISignIn {
  state: { user: IGraphUser | undefined; error: string | undefined; isLoading: boolean };
  actions: { signIn: () => Promise<void> };
  computed: { needsSignIn: boolean };
}

/** Silent NAA sign-in on load; `signIn` retries with the popup fallback (call from a click). */
export function useSignIn(): ISignIn {
  const [user, setUser] = React.useState<IGraphUser | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);

  const loadUser = React.useCallback(async (allowPopup: boolean): Promise<void> => {
    setIsLoading(true);
    setError(undefined);
    try {
      // Popup waits for the user, so only the silent path is time-boxed.
      const token = getGraphToken(allowPopup);
      const accessToken = await (allowPopup
        ? token
        : withTimeout(token, SIGN_IN_TIMEOUT_MS, "Silent sign-in timed out."));
      // Proves Graph with a scope the add-in holds; /me would need User.Read, which is not consented.
      await graphClient.api("/sites/root").select("id").get();
      setUser(userFromToken(accessToken));
    } catch (e) {
      // Silent failure is normal (e.g. Office on the web dev sideload): show just the Sign in button.
      if (allowPopup) setError(friendlyAuthError(e));
    } finally {
      setIsLoading(false);
    }
  }, []);

  React.useEffect(() => {
    void loadUser(false);
  }, [loadUser]);

  const signIn = React.useCallback(() => loadUser(true), [loadUser]);
  const needsSignIn = React.useMemo(() => !isLoading && !user, [isLoading, user]);

  return { state: { user, error, isLoading }, actions: { signIn }, computed: { needsSignIn } };
}
