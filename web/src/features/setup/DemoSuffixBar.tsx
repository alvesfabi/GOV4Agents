import { useState, useEffect } from 'react';
import {
  Field,
  Input,
  Button,
  Spinner,
  makeStyles,
  tokens,
  Body1,
  Caption1,
} from '@fluentui/react-components';
import { useApi } from '../../api/client';
import { useDemoSuffix } from './DemoSuffixContext';

const useStyles = makeStyles({
  bar: {
    display: 'flex',
    alignItems: 'flex-end',
    gap: '12px',
    padding: '12px 16px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    flexWrap: 'wrap',
  },
  saved: {
    color: tokens.colorPaletteGreenForeground1,
    fontWeight: 600,
  },
});

export function DemoSuffixBar() {
  const styles = useStyles();
  const apiFetch = useApi();
  const { suffix: ctxSuffix, setSuffix: setCtxSuffix } = useDemoSuffix();
  const [input, setInput] = useState('');
  const [saved, setSaved] = useState('');
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    if (ctxSuffix) {
      setInput(ctxSuffix);
      setSaved(ctxSuffix);
    }
  }, [ctxSuffix]);

  async function save() {
    if (!input.trim()) return;
    setBusy(true);
    try {
      await apiFetch('/api/setup/suffix', {
        method: 'POST',
        body: JSON.stringify({ suffix: input.trim() }),
      });
      setSaved(input.trim());
      setCtxSuffix(input.trim());
    } catch {
      // ignore
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className={styles.bar}>
      <Field
        label="Demo suffix"
        hint="Appended to all resource names (catalog, packages, workflow, CSA). Set before running steps."
      >
        <Input
          value={input}
          placeholder="e.g. Fabian1"
          onChange={(_e, data) => setInput(data.value)}
          onKeyDown={(e) => { if (e.key === 'Enter') save(); }}
        />
      </Field>
      <Button appearance="primary" disabled={busy || !input.trim() || input.trim() === saved} onClick={save}>
        {busy ? <Spinner size="tiny" /> : 'Set'}
      </Button>
      {saved && (
        <Body1>
          Current: <span className={styles.saved}>{saved}</span>
          <br />
          <Caption1>
            Names will be like: "GOV4Agents Blueprint {saved}", "Agent — Group.Read.All {saved}", etc.
          </Caption1>
        </Body1>
      )}
    </div>
  );
}
