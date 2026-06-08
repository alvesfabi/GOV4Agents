import { useEffect, useState } from 'react';
import {
  Card,
  CardHeader,
  Subtitle2,
  Body1,
  Caption1,
  makeStyles,
  Spinner,
  tokens,
} from '@fluentui/react-components';
import { useApi } from '../../api/client';
import { ErrorPanel } from '../../components/ErrorPanel';
import { PortalLink, AppLink } from '../../components/Links';
import { portal } from '../../lib/portalLinks';

interface Summary {
  blueprint: { id: string; name: string; principalId?: string } | null;
  agent: { id: string; clientId: string; name: string; hasSecret: boolean } | null;
  catalog: { id: string; name: string } | null;
  accessPackage: { id: string; name: string; approverUpn?: string } | null;
  accessPackage2: { id: string; name: string; approverUpn?: string } | null;
  lcw: { id: string; name: string } | null;
  csa: { setName: string; attributeName: string } | null;
  caPolicy: { id: string; displayName: string } | null;
  tenantId: string;
}

const useStyles = makeStyles({
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))',
    gap: '12px',
  },
  card: { padding: '12px', display: 'flex', flexDirection: 'column', gap: '8px' },
  missing: { color: tokens.colorNeutralForeground3, fontStyle: 'italic' },
});

export function SummaryPage() {
  const styles = useStyles();
  const apiFetch = useApi();
  const [summary, setSummary] = useState<Summary | null>(null);
  const [error, setError] = useState<unknown>(null);

  useEffect(() => {
    apiFetch<Summary & { ok: boolean }>('/api/setup/summary').then(setSummary).catch(setError);
  }, [apiFetch]);

  if (error) return <ErrorPanel error={error} />;
  if (!summary) return <Spinner />;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
      <Body1>
        Tenant <code>{summary.tenantId}</code>. Each tile below links to the exact resource in
        the Entra portal.
      </Body1>
      <div className={styles.grid}>
        <ResourceCard
          title="Agent Blueprint"
          ready={Boolean(summary.blueprint)}
          name={summary.blueprint?.name}
          id={summary.blueprint?.id}
          href={
            summary.blueprint
              ? portal.agentBlueprint(summary.blueprint.principalId ?? summary.blueprint.id)
              : undefined
          }
          missingTo="/setup"
        />
        <ResourceCard
          title="Agent ID"
          ready={Boolean(summary.agent)}
          name={summary.agent?.name}
          id={summary.agent?.id}
          href={summary.agent ? portal.agentId(summary.agent.id) : undefined}
          missingTo="/setup"
        />
        <ResourceCard
          title="Catalog"
          ready={Boolean(summary.catalog)}
          name={summary.catalog?.name}
          id={summary.catalog?.id}
          href={summary.catalog ? portal.catalog(summary.catalog.id, summary.catalog.name) : undefined}
          missingTo="/setup"
        />
        <ResourceCard
          title="Access Package (Group.Read.All)"
          ready={Boolean(summary.accessPackage)}
          name={summary.accessPackage?.name}
          id={summary.accessPackage?.id}
          href={
            summary.accessPackage
              ? portal.accessPackage(
                  summary.accessPackage.id,
                  summary.catalog?.id,
                  summary.catalog?.name,
                  summary.accessPackage.name,
                )
              : undefined
          }
          missingTo="/setup"
        />
        <ResourceCard
          title="Access Package (Group.ReadWrite.All — SoD)"
          ready={Boolean(summary.accessPackage2)}
          name={summary.accessPackage2?.name}
          id={summary.accessPackage2?.id}
          href={
            summary.accessPackage2
              ? portal.accessPackage(
                  summary.accessPackage2.id,
                  summary.catalog?.id,
                  summary.catalog?.name,
                  summary.accessPackage2.name,
                )
              : undefined
          }
          missingTo="/setup"
        />
        <ResourceCard
          title="Lifecycle Workflow"
          ready={Boolean(summary.lcw)}
          name={summary.lcw?.name}
          id={summary.lcw?.id}
          href={summary.lcw ? portal.lcw(summary.lcw.id, summary.lcw.name) : undefined}
          missingTo="/setup"
        />
        <ResourceCard
          title="Custom Security Attribute"
          ready={Boolean(summary.csa)}
          name={summary.csa ? `${summary.csa.setName} / ${summary.csa.attributeName}` : undefined}
          href={summary.csa ? portal.csaSet(summary.csa.setName) : undefined}
          missingTo="/setup"
        />
        <ResourceCard
          title="Conditional Access Policy"
          ready={Boolean(summary.caPolicy)}
          name={summary.caPolicy?.displayName}
          id={summary.caPolicy?.id}
          href={summary.caPolicy ? portal.conditionalAccessPolicy(summary.caPolicy.id) : undefined}
          missingTo="/setup"
        />
      </div>
      <Body1>
        When everything is ready, continue to the <AppLink to="/manage">Manage</AppLink>{' '}
        journey.
      </Body1>
    </div>
  );
}

function ResourceCard(props: {
  title: string;
  ready: boolean;
  name?: string;
  id?: string;
  href?: string;
  missingTo: string;
}) {
  const styles = useStyles();
  return (
    <Card className={styles.card}>
      <CardHeader header={<Subtitle2>{props.title}</Subtitle2>} />
      {props.ready ? (
        <>
          <Body1>{props.name}</Body1>
          {props.id && <Caption1>id: {props.id}</Caption1>}
          {props.href && <PortalLink href={props.href}>Open in Entra</PortalLink>}
        </>
      ) : (
        <span className={styles.missing}>
          Not yet created — <AppLink to={props.missingTo}>create it</AppLink>
        </span>
      )}
    </Card>
  );
}
