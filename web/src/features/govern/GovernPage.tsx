import { useEffect, useState } from 'react';
import {
  Body1,
  Button,
  Caption1,
  Spinner,
  Subtitle2,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { useApi, ApiError } from '../../api/client';
import { TokenViewer } from '../../components/TokenViewer';
import { JsonViewer } from '../../components/JsonViewer';
import { ErrorPanel } from '../../components/ErrorPanel';
import { PortalLink, AppLink } from '../../components/Links';
import { portal } from '../../lib/portalLinks';
import { Page, PageHero, StepCard, Card } from '../../components/PageLayout';

interface Summary {
  agent: { id: string; clientId: string; name: string } | null;
  accessPackage: { id: string; name: string; approverUpn?: string } | null;
  accessPackage2: { id: string; name: string; approverUpn?: string } | null;
  lcw: { id: string; name: string } | null;
}

interface AgentTokenResponse {
  accessToken: string;
  expiresOn: string | null;
  scopes: string[];
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
  lifecycleImage: {
    width: '100%',
    maxHeight: '320px',
    objectFit: 'contain',
    borderRadius: tokens.borderRadiusXLarge,
    boxShadow: tokens.shadow8,
  },
  anatomy: {
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusLarge,
    padding: '16px 20px',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  anatomyTitle: { color: tokens.colorBrandForeground1 },
  claimList: { margin: 0, paddingLeft: '20px', display: 'flex', flexDirection: 'column', gap: '4px' },
});

export function GovernPage() {
  const styles = useStyles();
  const apiFetch = useApi();

  const [summary, setSummary] = useState<Summary | null>(null);
  const [token, setToken] = useState<AgentTokenResponse | null>(null);
  const [tokenError, setTokenError] = useState<unknown>(null);
  const [tokenBusy, setTokenBusy] = useState(false);

  const [usersResp, setUsersResp] = useState<GraphCallResponse | null>(null);
  const [usersErr, setUsersErr] = useState<unknown>(null);
  const [usersBusy, setUsersBusy] = useState(false);

  const [groupsResp, setGroupsResp] = useState<GraphCallResponse | null>(null);
  const [groupsErr, setGroupsErr] = useState<unknown>(null);
  const [groupsBusy, setGroupsBusy] = useState(false);

  const [lcwResp, setLcwResp] = useState<unknown>(null);
  const [lcwErr, setLcwErr] = useState<unknown>(null);
  const [lcwBusy, setLcwBusy] = useState(false);

  useEffect(() => {
    apiFetch<Summary>('/api/setup/summary').then(setSummary).catch(() => {});
  }, [apiFetch]);

  async function authenticate() {
    setTokenBusy(true);
    setTokenError(null);
    try {
      const r = await apiFetch<AgentTokenResponse>('/api/agent/token', { method: 'POST' });
      setToken(r);
    } catch (err) {
      setTokenError(err);
    } finally {
      setTokenBusy(false);
    }
  }

  async function callGraph(path: string): Promise<GraphCallResponse> {
    return apiFetch<GraphCallResponse>('/api/agent/graph', {
      method: 'POST',
      body: JSON.stringify({ path }),
    });
  }

  async function getUsers() {
    setUsersBusy(true);
    setUsersErr(null);
    try {
      setUsersResp(await callGraph('/users?$top=10'));
    } catch (err) {
      setUsersErr(err);
    } finally {
      setUsersBusy(false);
    }
  }

  async function getGroups() {
    setGroupsBusy(true);
    setGroupsErr(null);
    setGroupsResp(null);
    try {
      // Fresh token first so the agent picks up newly-granted permissions or
      // CA evaluation changes. Token API will hit Entra each click.
      await apiFetch<AgentTokenResponse>('/api/agent/token', { method: 'POST' });
      setGroupsResp(await callGraph('/groups?$top=10'));
    } catch (err) {
      setGroupsErr(err);
    } finally {
      setGroupsBusy(false);
    }
  }

  async function runOffboarding() {
    setLcwBusy(true);
    setLcwErr(null);
    try {
      setLcwResp(await apiFetch('/api/govern/run-offboarding', { method: 'POST' }));
    } catch (err) {
      setLcwErr(err);
    } finally {
      setLcwBusy(false);
    }
  }

  return (
    <Page>
      <PageHero
        title="Govern agent identities and lifecycle"
        subtitle="Governance ensures every agent has the right permissions, the right sponsors, and a well-defined lifecycle. We'll walk through the canonical agent journey."
      />
      <Card>
        <img
          src="/banners/govern-lifecycle.png"
          alt="Agent lifecycle: create, sponsor requests access, access approved, review, expire, delete"
          className={styles.lifecycleImage}
        />
      </Card>

      <StepCard number={1} label="Sponsor requests access for the agent">
        <Body1>
          First, let's see what the agent can do out of the box. Click <strong>Authenticate</strong>{' '}
          to get a Microsoft Graph access token for your agent (client-credentials flow). The
          decoded token will only have the <code>User.Read.All</code> role inherited from the
          blueprint.
        </Body1>
        <div className={styles.actions}>
          <Button appearance="primary" disabled={tokenBusy} onClick={authenticate}>
            {tokenBusy ? <Spinner size="tiny" /> : 'Authenticate'}
          </Button>
        </div>
        {tokenError ? <ErrorPanel error={tokenError} /> : null}
        {token && <TokenViewer token={token.accessToken} expiresOn={token.expiresOn} />}
        {token && (
          <div className={styles.anatomy}>
            <Subtitle2 className={styles.anatomyTitle}>Anatomy of this token</Subtitle2>
            <Body1>
              This is a JSON Web Token (JWT) issued by Microsoft Entra. Switch the tabs above
              to inspect the three parts — <strong>header</strong>, <strong>payload</strong>,
              and the raw signed string. The interesting bits live in the payload:
            </Body1>
            <ul className={styles.claimList}>
              <li>
                <Body1>
                  <code>aud</code> — the audience. For Microsoft Graph it's{' '}
                  <code>https://graph.microsoft.com</code>. The resource validates this claim
                  before honoring the token.
                </Body1>
              </li>
              <li>
                <Body1>
                  <code>iss</code> + <code>tid</code> — the issuer and your tenant ID. Confirms
                  the token came from your Entra tenant.
                </Body1>
              </li>
              <li>
                <Body1>
                  <code>appid</code> / <code>azp</code> — the client that requested the token.
                  This is your <strong>Agent ID's</strong> appId, not the blueprint's.
                </Body1>
              </li>
              <li>
                <Body1>
                  <code>idtyp: "app"</code> — marks this as an <strong>app-only</strong> token
                  (client credentials). There is no signed-in user, so you'll <em>not</em> see
                  a <code>scp</code> (delegated scopes) claim — permissions live in{' '}
                  <code>roles</code> instead.
                </Body1>
              </li>
              <li>
                <Body1>
                  <code>roles</code> — the application permissions granted to this identity.
                  You should see <code>User.Read.All</code>. That single role was{' '}
                  <em>inherited</em> from the blueprint: we assigned it to the blueprint
                  principal as an app role, and the agent identity automatically picks up
                  every inheritable role its blueprint holds.
                </Body1>
              </li>
              <li>
                <Body1>
                  <code>iat</code> / <code>nbf</code> / <code>exp</code> — issued-at,
                  not-before, and expiry timestamps. App-only Graph tokens are typically valid
                  for about 60–90 minutes.
                </Body1>
              </li>
            </ul>
            <Body1>
              <strong>How we got here:</strong> the agent identity itself can't hold
              credentials — those live on the <strong>blueprint</strong>'s application object.
              The server first signed in as the blueprint (clientId + blueprint client
              secret), then used a <strong>Federated Identity Credential</strong> exchange to
              ask Entra for a token whose <code>appid</code> is the agent's clientId. That is
              the token you are looking at now.
            </Body1>
          </div>
        )}

        <Body1>
          Now have the agent query the directory. <strong>Get first 10 users</strong> should
          succeed — the agent has <code>User.Read.All</code>.
        </Body1>
        <div className={styles.actions}>
          <Button disabled={usersBusy || !token} onClick={getUsers}>
            {usersBusy ? <Spinner size="tiny" /> : 'Query first 10 users'}
          </Button>
        </div>
        {usersErr ? <ErrorPanel error={usersErr} /> : null}
        {usersResp && <CallResult resp={usersResp} />}

        <Body1>
          Next, try <strong>Get first 10 groups</strong>. This will <em>fail</em> with a 403
          because the agent has not been granted <code>Groups.Read.All</code>. That's exactly
          what we want — least privilege by default.
        </Body1>
        <div className={styles.actions}>
          <Button disabled={groupsBusy || !token} onClick={getGroups}>
            {groupsBusy ? <Spinner size="tiny" /> : 'Query first 10 groups'}
          </Button>
        </div>
        {groupsErr ? <ErrorPanel error={groupsErr} /> : null}
        {groupsResp && <CallResult resp={groupsResp} />}

        <Body1>
          Now act as the <strong>sponsor</strong>: open the access package in MyAccess and
          request access for the agent. Then act as the <strong>approver</strong>: open
          MyAccess again and approve the request.
        </Body1>
        <div className={styles.actions}>
          {summary?.accessPackage?.id && (
            <PortalLink href={portal.myAccessRequest(summary.accessPackage.id)}>
              Open MyAccess request URL
            </PortalLink>
          )}
          <PortalLink href={portal.myAccessAgents()}>Open MyAccess (Agents tab)</PortalLink>
        </div>
        <Body1>
          Once the request is approved, click <strong>Query first 10 groups</strong> again.
          The agent will refresh its token and the call will succeed because the access
          package granted it <code>Groups.Read.All</code>.
        </Body1>
      </StepCard>

      <StepCard number={2} label="Separation of Duties: try requesting an incompatible package">
        <Body1>
          Now that the agent has an active assignment (or open request) for the{' '}
          <strong>Group.Read.All</strong> access package, try to request the second access
          package <strong>{summary?.accessPackage2?.name ?? 'Directory.Read.All'}</strong> from
          MyAccess <em>as the sponsor, on behalf of the agent</em>. The request will be{' '}
          <strong>blocked by the Separation of Duties rule</strong> we configured: the second
          package was declared incompatible with the first, so a requestor that already has (or
          is requesting) the first cannot acquire the second.
        </Body1>
        <Body1>
          This is how Entitlement Management enforces toxic-combination guardrails: instead of
          relying on reviewers to catch conflicting privileges, the platform refuses the
          request at submission time.
        </Body1>
        <div className={styles.actions}>
          {summary?.accessPackage2 && (
            <PortalLink href={portal.myAccessRequest(summary.accessPackage2.id)}>
              Request {summary.accessPackage2.name} in MyAccess (expect SoD failure)
            </PortalLink>
          )}
        </div>
      </StepCard>

      <StepCard number={3} label="Review and expire access assignments">
        <Body1>
          As the sponsor, open MyAccess and select your agent. You'll see every access
          assignment, who approved it, and when it will expire. This is where ongoing review
          ensures least privilege over time.
        </Body1>
        <div className={styles.actions}>
          <PortalLink href={portal.myAccessAgents()}>Open MyAccess (Agents tab)</PortalLink>
          {summary?.accessPackage2 && (
            <PortalLink href={portal.myAccessRequest(summary.accessPackage2.id)}>
              Request {summary.accessPackage2.name} in MyAccess
            </PortalLink>
          )}
        </div>
      </StepCard>

      <StepCard number={4} label="When the sponsor leaves">
        <Body1>
          Agents must always have a valid sponsor. <strong>Lifecycle Workflows</strong> let
          you automate sponsor offboarding: when an employee leaves, sponsorships are
          transferred to their manager and notifications are sent.
        </Body1>
        <div className={styles.actions}>
          {summary?.lcw?.id && (
            <PortalLink href={portal.lcw(summary.lcw.id, summary.lcw.name)}>
              Open the offboarding workflow in Entra
            </PortalLink>
          )}
        </div>
        <Body1>
          Click <strong>Run offboarding</strong> to execute the workflow now. After it runs,
          open the agent in Entra and confirm that the sponsor has been replaced by the
          previous sponsor's manager.
        </Body1>
        <div className={styles.actions}>
          <Button appearance="primary" disabled={lcwBusy || !summary?.lcw} onClick={runOffboarding}>
            {lcwBusy ? <Spinner size="tiny" /> : 'Run offboarding'}
          </Button>
          {summary?.agent?.id && (
            <PortalLink href={portal.agentId(summary.agent.id)}>Open the agent in Entra</PortalLink>
          )}
        </div>
        {lcwErr ? <ErrorPanel error={lcwErr} /> : null}
        {lcwResp != null && <JsonViewer value={lcwResp} label="Workflow run" />}
      </StepCard>

      <StepCard number={5} label="Lifecycle policies for agents">
        <Body1>
          IT admins can configure lifecycle policies to periodically re-confirm whether agent
          IDs should remain active, with customizable frequency, scope, and notification
          timing for sponsors. Sponsors receive notifications and can review agent details in
          the Manage Agents experience to decide whether to extend their lifecycle. If no
          action is taken, the agent ID is automatically disabled, then soft-deleted, and
          eventually permanently deleted if not restored.
        </Body1>
        <div className={styles.actions}>
          <PortalLink href={portal.lifecyclePolicies()}>
            Open Lifecycle policies in Entra
          </PortalLink>
        </div>
      </StepCard>

      <Card>
        <Body1>
          That covers governance. Continue to the <AppLink to="/protect">Protect</AppLink>{' '}
          journey to see Conditional Access for agents.
        </Body1>
      </Card>
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

// Re-export so unused-warning silenced if ApiError ever becomes useful here.
export type { ApiError };
