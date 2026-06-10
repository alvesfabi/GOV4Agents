import { AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { Body1, Title3, makeStyles, tokens } from '@fluentui/react-components';
import { AuthBar } from './auth/AuthBar';
import { AppRoutes } from './routes/AppRoutes';
import { SessionGate } from './components/SessionGate';
import { Page, PageHero, Card } from './components/PageLayout';

const useStyles = makeStyles({
  splash: {
    minHeight: '100vh',
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  splashContent: {
    padding: '32px 48px 64px',
  },
  diagram: {
    width: '100%',
    maxHeight: '480px',
    objectFit: 'contain',
    borderRadius: tokens.borderRadiusXLarge,
    boxShadow: tokens.shadow8,
  },
});

export default function App() {
  const styles = useStyles();
  return (
    <>
      <AuthenticatedTemplate>
        <SessionGate>
          <AppRoutes />
        </SessionGate>
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <div className={styles.splash}>
          <AuthBar />
          <div className={styles.splashContent}>
            <Page>
              <PageHero
                title="Welcome to GOV4Agents"
                subtitle="Sign in with an Entra account in your test tenant to begin the Manage · Govern · Protect proof of concept."
              />
              <Card>
                <Title3>What is Microsoft Entra Agent ID?</Title3>
                <Body1>
                  Microsoft Entra Agent ID is an identity and security framework that extends
                  Microsoft Entra capabilities to AI agents. As organizations deploy assistive,
                  autonomous, and user-like agents, they need purpose-built identity constructs
                  to authenticate, authorize, govern, and protect these nonhuman identities.
                  Microsoft Entra Agent ID addresses these needs by providing a unified
                  platform for managing agent identities at enterprise scale.
                </Body1>
                <img
                  src="/banners/agent-id-overview.png"
                  alt="Diagram showing agent security capabilities offered by Microsoft Entra Agent ID."
                  className={styles.diagram}
                />
              </Card>
            </Page>
          </div>
        </div>
      </UnauthenticatedTemplate>
    </>
  );
}
