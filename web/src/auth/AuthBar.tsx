import { useMsal } from '@azure/msal-react';
import {
  Button,
  Avatar,
  makeStyles,
  tokens,
  Body1Strong,
  Caption1,
} from '@fluentui/react-components';
import { Scopes } from './scopes';

const useStyles = makeStyles({
  bar: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '8px 16px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  spacer: { flex: 1 },
  user: { display: 'flex', flexDirection: 'column' },
});

export function AuthBar() {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  if (!account) {
    return (
      <div className={styles.bar}>
        <Body1Strong>GOV4Agents</Body1Strong>
        <div className={styles.spacer} />
        <Button
          appearance="primary"
          onClick={() => instance.loginPopup({ scopes: Scopes.signIn })}
        >
          Sign in
        </Button>
      </div>
    );
  }

  const tenantId = (account.idTokenClaims as { tid?: string } | undefined)?.tid;
  return (
    <div className={styles.bar}>
      <Body1Strong>GOV4Agents</Body1Strong>
      <div className={styles.spacer} />
      <div className={styles.user}>
        <Body1Strong>{account.name ?? account.username}</Body1Strong>
        <Caption1>tenant: {tenantId ?? 'unknown'}</Caption1>
      </div>
      <Avatar name={account.name ?? account.username} />
      <Button appearance="subtle" onClick={() => instance.logoutPopup({ account })}>
        Sign out
      </Button>
    </div>
  );
}
