import * as React from "react";
import { getGraphToken } from "../auth/naa";
import { graphClient } from "../services/graphClient";

interface IGraphUser {
  displayName: string;
  userPrincipalName: string;
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
      await getGraphToken(allowPopup);
      setUser(await graphClient.api("/me").select("displayName,userPrincipalName").get());
    } catch (e) {
      setError((e as Error).message);
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
