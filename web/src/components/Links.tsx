import { Link as RouterLink } from 'react-router-dom';
import { Open16Regular } from '@fluentui/react-icons';
import { Link, makeStyles, tokens } from '@fluentui/react-components';

const useStyles = makeStyles({
  link: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    color: tokens.colorBrandForegroundLink,
  },
});

interface PortalLinkProps {
  href: string;
  children: React.ReactNode;
}

export function PortalLink({ href, children }: PortalLinkProps) {
  const styles = useStyles();
  return (
    <Link href={href} target="_blank" rel="noopener noreferrer" className={styles.link}>
      {children} <Open16Regular />
    </Link>
  );
}

interface AppLinkProps {
  to: string;
  children: React.ReactNode;
}

export function AppLink({ to, children }: AppLinkProps) {
  return (
    <RouterLink to={to} style={{ color: tokens.colorBrandForegroundLink }}>
      {children}
    </RouterLink>
  );
}
