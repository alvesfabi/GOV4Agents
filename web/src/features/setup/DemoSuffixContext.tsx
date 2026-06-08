import { createContext, useContext, useState, useEffect, type ReactNode } from 'react';
import { useApi } from '../../api/client';

interface SuffixCtx {
  suffix: string;
  setSuffix: (s: string) => void;
}

const DemoSuffixContext = createContext<SuffixCtx>({ suffix: '', setSuffix: () => {} });

export function DemoSuffixProvider({ children }: { children: ReactNode }) {
  const apiFetch = useApi();
  const [suffix, setSuffix] = useState('');

  useEffect(() => {
    apiFetch<{ suffix: string }>('/api/setup/suffix')
      .then((r) => { if (r.suffix) setSuffix(r.suffix); })
      .catch(() => {});
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
