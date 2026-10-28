import { Client } from "@microsoft/microsoft-graph-client";
import { getGraphToken } from "../auth/naa";

/** Graph client authenticated through NAA (delegated, as the signed-in Office user). */
export const graphClient: Client = Client.init({
  authProvider: (done) => {
    getGraphToken().then(
      (token) => done(null, token),
      (error: Error) => done(error, null)
    );
  },
});
