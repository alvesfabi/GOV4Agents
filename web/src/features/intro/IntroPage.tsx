import { makeStyles, tokens, Title3, Body1 } from '@fluentui/react-components';
import { AppLink } from '../../components/Links';
import { Page, PageHero, Card } from '../../components/PageLayout';

const useStyles = makeStyles({
  pillarGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))',
    gap: '16px',
  },
  pillar: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    padding: '20px',
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusXLarge,
    boxShadow: tokens.shadow2,
  },
  pillarTitle: { color: tokens.colorBrandForeground1 },
});

interface Pillar {
  title: string;
  body: string;
}

const pillars: Pillar[] = [
  {
    title: 'Agent Blueprint',
    body: 'A template that standardizes how agents are provisioned, including inheritable Microsoft Graph permissions.',
  },
  {
    title: 'Agent ID',
    body: 'The actual agent identity that inherits from the blueprint and can authenticate to Entra.',
  },
  {
    title: 'Catalog + Access Package',
    body: 'Entitlement management surface that grants the agent additional Microsoft Graph permissions on demand.',
  },
  {
    title: 'Lifecycle Workflow',
    body: 'Automates sponsor offboarding so agents always have a valid sponsor.',
  },
  {
    title: 'Custom Security Attribute + Conditional Access',
    body: 'Protects resources by allowing only approved agents to access them.',
  },
];

export function IntroPage() {
  const styles = useStyles();
  return (
    <Page>
      <PageHero
        title="Manage · Govern · Protect AI agents"
        subtitle="This proof of concept walks you through Microsoft Entra's strategy for AI agents and actually creates the components in your tenant so you can interact with them."
        banner="/banners/manage-govern-protect.png"
        bannerAlt="Manage, Govern, and Protect pillars for AI agents"
      />
      <Card>
        <Title3>What we'll build</Title3>
        <div className={styles.pillarGrid}>
          {pillars.map((p) => (
            <div key={p.title} className={styles.pillar}>
              <Title3 className={styles.pillarTitle}>{p.title}</Title3>
              <Body1>{p.body}</Body1>
            </div>
          ))}
        </div>
      </Card>
      <Card>
        <Title3>What comes next</Title3>
        <Body1>
          Once everything is created we'll walk you through the <strong>Manage</strong>,{' '}
          <strong>Govern</strong>, and <strong>Protect</strong> journeys against your live
          tenant.
        </Body1>
        <Body1>
          Begin with the <AppLink to="/setup">Setup</AppLink> page to create all required
          resources.
        </Body1>
      </Card>
    </Page>
  );
}
