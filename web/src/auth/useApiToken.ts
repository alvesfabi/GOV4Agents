import { useMsal } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { apiScope } from './msalConfig';

/**
 * Returns a function that resolves an access token for our backend API
 * (audience = api://<clientId>/access_as_user). The backend uses OBO to swap
 * this for a Graph token with the requested scopes server-side.
 */
export function useApiToken() {
  const { instance, accounts } = useMsal();
  return async (): Promise<string> => {
    const account = accounts[0];
    if (!account) throw new Error('Not signed in');
    try {
      const r = await instance.acquireTokenSilent({ account, scopes: [apiScope] });
      return r.accessToken;
    } catch (err) {
      if (err instanceof InteractionRequiredAuthError) {
        const r = await instance.acquireTokenPopup({ account, scopes: [apiScope] });
        return r.accessToken;
      }
      throw err;
    }
  };
}
