import { jwtDecode } from 'jwt-decode';
import {
  Body1,
  Caption1,
  makeStyles,
  tokens,
  Tab,
  TabList,
  type SelectTabEvent,
  type SelectTabData,
} from '@fluentui/react-components';
import { useState } from 'react';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    padding: '12px',
  },
  pre: {
    margin: 0,
    padding: '8px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusSmall,
    overflowX: 'auto',
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '12px',
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-all',
  },
});

interface TokenViewerProps {
  token: string;
  expiresOn?: string | null;
}

export function TokenViewer({ token, expiresOn }: TokenViewerProps) {
  const styles = useStyles();
  const [tab, setTab] = useState<'raw' | 'header' | 'payload'>('payload');
  let header: unknown = null;
  let payload: unknown = null;
  try {
    header = jwtDecode(token, { header: true });
    payload = jwtDecode(token);
  } catch {
    /* ignore */
  }

  const onTab = (_e: SelectTabEvent, data: SelectTabData) => {
    setTab(data.value as 'raw' | 'header' | 'payload');
  };

  return (
    <div className={styles.root}>
      <TabList selectedValue={tab} onTabSelect={onTab} size="small">
        <Tab value="payload">Decoded payload</Tab>
        <Tab value="header">Header</Tab>
        <Tab value="raw">Raw JWT</Tab>
      </TabList>
      {expiresOn && <Caption1>Expires: {expiresOn}</Caption1>}
      {tab === 'raw' && <pre className={styles.pre}>{token}</pre>}
      {tab === 'header' && (
        <pre className={styles.pre}>{JSON.stringify(header, null, 2)}</pre>
      )}
      {tab === 'payload' && (
        <pre className={styles.pre}>{JSON.stringify(payload, null, 2)}</pre>
      )}
      {payload != null && typeof payload === 'object' && 'roles' in payload ? (
        <Body1>
          <strong>roles claim:</strong>{' '}
          {JSON.stringify((payload as { roles: unknown }).roles)}
        </Body1>
      ) : null}
    </div>
  );
}
