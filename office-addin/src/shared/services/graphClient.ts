import {
  Client,
  CustomAuthenticationProvider,
  MiddlewareFactory,
  type Context,
  type Middleware,
} from "@microsoft/microsoft-graph-client";
import { clearDialogToken, getGraphToken } from "../auth/naa";

const HTTP_UNAUTHORIZED = 401;

/** Drops a stored dialog token Graph rejects (revoked, consent changed), so the next call asks to sign in again. */
class ClearRejectedTokenHandler implements Middleware {
  private next: Middleware | undefined;

  public async execute(context: Context): Promise<void> {
    await this.next?.execute(context);
    if (context.response?.status === HTTP_UNAUTHORIZED) clearDialogToken();
  }

  public setNext(next: Middleware): void {
    this.next = next;
  }
}

const authProvider = new CustomAuthenticationProvider((done) => {
  getGraphToken().then(
    (token) => done(null, token),
    (error: Error) => done(error, null)
  );
});

/** Graph client authenticated as the signed-in Office user (delegated), default SDK middleware (retry, redirect). */
export const graphClient: Client = Client.initWithMiddleware({
  middleware: [
    new ClearRejectedTokenHandler(),
    ...MiddlewareFactory.getDefaultMiddlewareChain(authProvider),
  ],
});
