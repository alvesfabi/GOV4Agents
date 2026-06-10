import { useEffect, useState, type ReactNode } from 'react';
import { Spinner } from '@fluentui/react-components';
import { useApiToken } from '../auth/useApiToken';
import { rehydrateServer, syncContextToLocal } from '../auth/sessionContext';

/**
 * On load, restores the demo context: pushes the locally persisted snapshot
 * back into the server session (covers a restarted API server), then pulls the
 * server's current snapshot so localStorage reflects the latest truth. Blocks
 * route rendering briefly so journey pages see a populated session.
 */
export function SessionGate({ children }: { children: ReactNode }) {
  const getToken = useApiToken();
  const [ready, setReady] = useState(false);

  useEffect(() => {
    let active = true;
    (async () => {
      try {
        const token = await getToken();
        await rehydrateServer(token);
        await syncContextToLocal(token);
      } catch {
        /* proceed even if restore fails */
      }
      if (active) setReady(true);
    })();
    return () => {
      active = false;
    };
  }, [getToken]);

  if (!ready) {
    return (
      <div style={{ minHeight: '100vh', display: 'grid', placeItems: 'center' }}>
        <Spinner label="Restoring your session…" />
      </div>
    );
  }
  return <>{children}</>;
}
