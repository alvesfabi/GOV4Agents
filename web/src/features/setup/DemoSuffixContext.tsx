import { createContext, useContext, useState, useEffect, type ReactNode } from 'react';
import { useApi } from '../../api/client';

interface SuffixCtx {
  suffix: string;
  setSuffix: (s: string) => void;
}

const DemoSuffixContext = createContext<SuffixCtx>({ suffix: '', setSuffix: () => {} });

const SUFFIX_KEY = 'gov4.demoSuffix';

function readStoredSuffix(): string {
  try {
    return localStorage.getItem(SUFFIX_KEY) ?? '';
  } catch {
    return '';
  }
}

export function DemoSuffixProvider({ children }: { children: ReactNode }) {
  const apiFetch = useApi();
  // Seed from localStorage so the environment name is pre-filled after closing
  // and reopening the browser (it is not a secret). The server session is the
  // source of truth while it is alive; if it was restarted and has no suffix,
  // we keep the locally stored value so the user can re-run setup as-is.
  const [suffix, setSuffixState] = useState(readStoredSuffix);

  const setSuffix = (s: string) => {
    setSuffixState(s);
    try {
      if (s) localStorage.setItem(SUFFIX_KEY, s);
    } catch {
      /* ignore */
    }
  };

  useEffect(() => {
    apiFetch<{ suffix: string }>('/api/setup/suffix')
      .then((r) => { if (r.suffix) setSuffix(r.suffix); })
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [apiFetch]);

  return (
    <DemoSuffixContext.Provider value={{ suffix, setSuffix }}>
      {children}
    </DemoSuffixContext.Provider>
  );
}

export function useDemoSuffix() {
  return useContext(DemoSuffixContext);
}
