import { useEffect, useState } from 'react';
import { Title3, Body1, Spinner, makeStyles, tokens } from '@fluentui/react-components';
import { useApi } from '../../api/client';
import { ErrorPanel } from '../../components/ErrorPanel';
import { PortalLink, AppLink } from '../../components/Links';
import { portal } from '../../lib/portalLinks';
import { Page, PageHero, Card } from '../../components/PageLayout';

interface Summary {
  blueprint: { id: string; name: string; principalId?: string } | null;
  agent: { id: string; clientId: string; name: string } | null;
}

const useStyles = makeStyles({
  cards: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))',
    gap: '16px',
  },
  pillTitle: { color: tokens.colorBrandForeground1 },
});

export function ManagePage() {
  const styles = useStyles();
  const apiFetch = useApi();
  const [summary, setSummary] = useState<Summary | null>(null);
  const [error, setError] = useState<unknown>(null);

  useEffect(() => {
    apiFetch<Summary>('/api/setup/summary').then(setSummary).catch(setError);
  }, [apiFetch]);

  return (
    <Page>
      <PageHero
        title="Manage AI agents at scale"
        subtitle="The first pillar of Microsoft's strategy is managing AI agents consistently. The two foundational artifacts are the Agent Blueprint and the Agent ID."
        banner="/banners/manage.png"
        bannerAlt="Manage pillar diagram"
      />
      <Card>
        <Title3>Foundations</Title3>
        <Body1>
          <strong>Agent Blueprint</strong> — a reusable template that captures how an agent
          should be provisioned, including ownership, sponsorship, and the inheritable
          permissions every agent created from it should automatically receive.
        </Body1>
        <Body1>
          <strong>Agent ID</strong> — the actual identity that represents an AI agent in your
          tenant. It inherits its baseline permissions from a blueprint, has owners and
          sponsors, and authenticates to Entra to access resources.
        </Body1>
      </Card>

      {error ? <ErrorPanel error={error} /> : null}
      {!summary && !error && <Spinner />}
      {summary && (
        <div className={styles.cards}>
          <Card>
            <Title3 className={styles.pillTitle}>Your Agent Blueprint</Title3>
            {summary.blueprint ? (
              <>
                <Body1>{summary.blueprint.name}</Body1>
                <PortalLink
                  href={portal.agentBlueprint(
                    summary.blueprint.principalId ?? summary.blueprint.id,
                  )}
                >
                  Open this blueprint in Entra
                </PortalLink>
                <Body1>
                  Look at the <em>Inheritable permissions</em> blade — you'll see{' '}
                  <code>User.Read.All</code> assigned and marked as inheritable.
                </Body1>
              </>
            ) : (
              <Body1>
                Not yet created. <AppLink to="/setup">Go to Setup</AppLink>
              </Body1>
            )}
          </Card>
          <Card>
            <Title3 className={styles.pillTitle}>Your Agent ID</Title3>
            {summary.agent ? (
              <>
                <Body1>{summary.agent.name}</Body1>
                <PortalLink href={portal.agentId(summary.agent.id)}>
                  Open this agent in Entra
                </PortalLink>
                <Body1>
                  Notice the agent is linked to your blueprint and uses the sponsor you
                  selected during setup.
                </Body1>
              </>
            ) : (
              <Body1>
                Not yet created. <AppLink to="/setup">Go to Setup</AppLink>
              </Body1>
            )}
          </Card>
        </div>
      )}

      <Card>
        <Body1>
          Once you've inspected the blueprint and the agent, continue to the{' '}
          <AppLink to="/govern">Govern</AppLink> journey to see governance in action.
        </Body1>
      </Card>
    </Page>
  );
}
