import { useEffect, useState } from 'react';
import {
  Body1,
  Button,
  Caption1,
  Spinner,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { useApi } from '../../api/client';
import { JsonViewer } from '../../components/JsonViewer';
import { ErrorPanel } from '../../components/ErrorPanel';
import { PortalLink } from '../../components/Links';
import { portal } from '../../lib/portalLinks';
import { Page, PageHero, StepCard } from '../../components/PageLayout';

interface Summary {
  agent: { id: string; clientId: string; name: string } | null;
  caPolicy: { id: string; displayName: string } | null;
  csa: { setName: string; attributeName: string } | null;
}

interface GraphCallResponse {
  status: number;
  ok: boolean;
  body: unknown;
}

const useStyles = makeStyles({
  actions: { display: 'flex', gap: '8px', flexWrap: 'wrap' },
  status: {
    padding: '6px 10px',
    borderRadius: tokens.borderRadiusMedium,
    fontSize: tokens.fontSizeBase200,
    fontWeight: 600,
    alignSelf: 'flex-start',
  },
  ok: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
  },
  fail: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
  },
  note: {
    padding: '10px 14px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorPaletteYellowBackground1,
    border: `1px solid ${tokens.colorPaletteYellowBorder1}`,
    color: tokens.colorPaletteDarkOrangeForeground1,
    fontSize: tokens.fontSizeBase200,
  },
});

export function ProtectPage() {
  const styles = useStyles();
  const apiFetch = useApi();
  const [summary, setSummary] = useState<Summary | null>(null);

  const [enableResp, setEnableResp] = useState<unknown>(null);
  const [enableErr, setEnableErr] = useState<unknown>(null);
  const [enableBusy, setEnableBusy] = useState(false);

  const [groupsBeforeResp, setGroupsBeforeResp] = useState<GraphCallResponse | null>(null);
  const [groupsBeforeErr, setGroupsBeforeErr] = useState<unknown>(null);
  const [groupsBeforeBusy, setGroupsBeforeBusy] = useState(false);

  const [groupsAfterResp, setGroupsAfterResp] = useState<GraphCallResponse | null>(null);
  const [groupsAfterErr, setGroupsAfterErr] = useState<unknown>(null);
  const [groupsAfterBusy, setGroupsAfterBusy] = useState(false);

  const [csaResp, setCsaResp] = useState<unknown>(null);
  const [csaErr, setCsaErr] = useState<unknown>(null);
  const [csaBusy, setCsaBusy] = useState(false);

  useEffect(() => {
    apiFetch<Summary>('/api/setup/summary').then(setSummary).catch(() => {});
  }, [apiFetch]);

  async function enablePolicy() {
    setEnableBusy(true);
    setEnableErr(null);
    try {
      setEnableResp(await apiFetch('/api/protect/enable-ca', { method: 'POST' }));
    } catch (err) {
      setEnableErr(err);
    } finally {
      setEnableBusy(false);
    }
  }

  async function fetchGroups(): Promise<GraphCallResponse> {
    // Force a fresh agent token so the latest CA evaluation runs.
    await apiFetch('/api/agent/token', { method: 'POST' });
    return apiFetch<GraphCallResponse>('/api/agent/graph', {
      method: 'POST',
      body: JSON.stringify({ path: '/groups?$top=10' }),
    });
  }

  async function getGroupsBefore() {
    setGroupsBeforeBusy(true);
    setGroupsBeforeErr(null);
    setGroupsBeforeResp(null);
    try {
      setGroupsBeforeResp(await fetchGroups());
    } catch (err) {
      setGroupsBeforeErr(err);
    } finally {
      setGroupsBeforeBusy(false);
    }
  }

  async function getGroupsAfter() {
    setGroupsAfterBusy(true);
    setGroupsAfterErr(null);
    setGroupsAfterResp(null);
    try {
      setGroupsAfterResp(await fetchGroups());
    } catch (err) {
      setGroupsAfterErr(err);
    } finally {
      setGroupsAfterBusy(false);
    }
  }

  async function approveAgent() {
    setCsaBusy(true);
    setCsaErr(null);
    try {
      setCsaResp(
        await apiFetch('/api/protect/set-csa', {
          method: 'POST',
          body: JSON.stringify({ value: 'approved' }),
        }),
      );
    } catch (err) {
      setCsaErr(err);
    } finally {
      setCsaBusy(false);
    }
  }

  return (
    <Page>
      <PageHero
        title="Protect agent access to resources"
        subtitle={
          <>
            Conditional Access for Agents lets you apply policy controls to workload
            identities the same way you do for users. We've already created a policy{' '}
            <strong>Block not approved agents</strong> in <em>report-only</em> mode that
            targets any agent whose <code>AgentsCSA / TAG</code> attribute is not{' '}
            <code>approved</code>.
          </>
        }
        banner="/banners/protect.png"
        bannerAlt="Protect pillar diagram"
      />

      <StepCard number={1} label="Enable the policy">
        <Body1>
          Flip the policy from report-only to enabled. From here on, any agent without the{' '}
          <code>approved</code> tag will be blocked.
        </Body1>
        <div className={styles.actions}>
          <Button appearance="primary" disabled={enableBusy || !summary?.caPolicy} onClick={enablePolicy}>
            {enableBusy ? <Spinner size="tiny" /> : 'Enable policy'}
          </Button>
          {summary?.caPolicy?.id && (
            <PortalLink href={portal.conditionalAccessPolicy(summary.caPolicy.id)}>
              Open policy in Entra
            </PortalLink>
          )}
        </div>
        {enableErr ? <ErrorPanel error={enableErr} /> : null}
        {enableResp != null && <JsonViewer value={enableResp} label="Patch result" />}
      </StepCard>

      <StepCard number={2} label="Try to call Graph as the agent">
        <Body1>
          The agent currently has no <code>TAG</code> value, so the policy will block its
          authentication. <strong>Get groups</strong> should fail.
        </Body1>
        <div className={styles.note}>
          <strong>Note:</strong> Conditional Access policy changes can take up to a minute to
          propagate. If the call still succeeds right after enabling the policy, wait about a
          minute and try again.
        </div>
        <div className={styles.actions}>
          <Button disabled={groupsBeforeBusy} onClick={getGroupsBefore}>
            {groupsBeforeBusy ? <Spinner size="tiny" /> : 'Get groups'}
          </Button>
        </div>
        {groupsBeforeErr ? <ErrorPanel error={groupsBeforeErr} /> : null}
        {groupsBeforeResp && <CallResult resp={groupsBeforeResp} />}
      </StepCard>

      <StepCard number={3} label="Approve the agent (update CSA)">
        <Body1>
          To satisfy the policy, set the agent's <code>AgentsCSA / TAG</code> attribute to{' '}
          <code>approved</code>. In a real deployment this is the gate your governance team
          would control.
        </Body1>
        <div className={styles.actions}>
          <Button
            appearance="primary"
            disabled={csaBusy || !summary?.csa || !summary?.agent}
            onClick={approveAgent}
          >
            {csaBusy ? <Spinner size="tiny" /> : 'Update CSA → approved'}
          </Button>
          {summary?.agent?.id && (
            <PortalLink href={portal.agentId(summary.agent.id)}>
              View the agent in Entra
            </PortalLink>
          )}
        </div>
        {csaErr ? <ErrorPanel error={csaErr} /> : null}
        {csaResp != null && <JsonViewer value={csaResp} label="CSA update result" />}
      </StepCard>

      <StepCard number={4} label="Try again">
        <Body1>
          Now that the agent is tagged <code>approved</code>, the Conditional Access filter no
          longer matches it. Click <strong>Get groups</strong> again — it should succeed.
        </Body1>
        <div className={styles.note}>
          <strong>Note:</strong> Custom security attribute changes can also take up to a
          minute to be reflected during policy evaluation. If the call still fails right after
          updating the CSA, wait about a minute and try again.
        </div>
        <div className={styles.actions}>
          <Button disabled={groupsAfterBusy} onClick={getGroupsAfter}>
            {groupsAfterBusy ? <Spinner size="tiny" /> : 'Get groups'}
          </Button>
        </div>
        {groupsAfterErr ? <ErrorPanel error={groupsAfterErr} /> : null}
        {groupsAfterResp && <CallResult resp={groupsAfterResp} />}
      </StepCard>
    </Page>
  );
}

function CallResult({ resp }: { resp: GraphCallResponse }) {
  const styles = useStyles();
  const isList = resp.ok && Array.isArray((resp.body as { value?: unknown[] })?.value);
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
      <span className={`${styles.status} ${resp.ok ? styles.ok : styles.fail}`}>
        HTTP {resp.status} {resp.ok ? 'OK' : 'failed'}
      </span>
      {isList ? (
        <ResultTable items={(resp.body as { value: Array<Record<string, unknown>> }).value} />
      ) : (
        <JsonViewer value={resp.body} label="Graph response body" />
      )}
    </div>
  );
}

function ResultTable({ items }: { items: Array<Record<string, unknown>> }) {
  if (items.length === 0) return <Caption1>No results</Caption1>;
  const cols = ['displayName', 'userPrincipalName', 'mail', 'id'].filter((c) =>
    items.some((it) => it[c] !== undefined),
  );
  return (
    <table style={{ borderCollapse: 'collapse', fontSize: '12px' }}>
      <thead>
        <tr>
          {cols.map((c) => (
            <th
              key={c}
              style={{ textAlign: 'left', padding: '4px 8px', borderBottom: '1px solid #ccc' }}
            >
              {c}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {items.map((it, i) => (
          <tr key={i}>
            {cols.map((c) => (
              <td key={c} style={{ padding: '4px 8px', borderBottom: '1px solid #eee' }}>
                {String(it[c] ?? '')}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
