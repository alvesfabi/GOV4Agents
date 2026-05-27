import type { ReactNode } from 'react';
import { NavLink, Outlet } from 'react-router-dom';
import { makeStyles, tokens, Caption1, Subtitle2 } from '@fluentui/react-components';
import { AuthBar } from '../auth/AuthBar';

const useStyles = makeStyles({
  root: {
    minHeight: '100vh',
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  layout: {
    display: 'grid',
    gridTemplateColumns: '260px 1fr',
    flex: 1,
    minHeight: 0,
  },
  rail: {
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    background: `linear-gradient(180deg, ${tokens.colorBrandBackground2} 0%, ${tokens.colorNeutralBackground1} 280px)`,
    padding: '20px 14px',
    display: 'flex',
    flexDirection: 'column',
    gap: '20px',
    overflowY: 'auto',
  },
  brand: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '6px 8px 14px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  brandMark: {
    width: '36px',
    height: '36px',
    borderRadius: tokens.borderRadiusMedium,
    background: `linear-gradient(135deg, ${tokens.colorBrandBackground}, ${tokens.colorBrandBackgroundHover})`,
    color: tokens.colorNeutralForegroundOnBrand,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontWeight: 800,
    fontSize: tokens.fontSizeBase400,
    boxShadow: tokens.shadow4,
  },
  brandText: {
    display: 'flex',
    flexDirection: 'column',
    lineHeight: 1.1,
  },
  brandTitle: { color: tokens.colorNeutralForeground1 },
  brandSub: { color: tokens.colorNeutralForeground3 },
  group: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  groupHeader: {
    padding: '4px 12px 6px',
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.08em',
    fontWeight: 700,
  },
  link: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    padding: '9px 12px',
    borderRadius: tokens.borderRadiusMedium,
    textDecoration: 'none',
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase300,
    transition: 'background-color 120ms ease, color 120ms ease, transform 120ms ease',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      color: tokens.colorNeutralForeground1,
    },
  },
  icon: {
    width: '22px',
    textAlign: 'center',
    fontSize: '14px',
    flexShrink: 0,
  },
  active: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    fontWeight: 600,
    boxShadow: tokens.shadow4,
    ':hover': {
      backgroundColor: tokens.colorBrandBackgroundHover,
      color: tokens.colorNeutralForegroundOnBrand,
    },
  },
  content: {
    padding: '32px 48px',
    overflowY: 'auto',
  },
});

interface NavItem {
  to: string;
  label: string;
  icon: string;
  end?: boolean;
}

interface NavGroup {
  header: string;
  items: NavItem[];
}

const groups: NavGroup[] = [
  {
    header: 'Start',
    items: [{ to: '/', label: 'Intro', icon: '🏠', end: true }],
  },
  {
    header: 'Setup',
    items: [{ to: '/setup', label: 'Setup (all steps)', icon: '🧰' }],
  },
  {
    header: 'Journey',
    items: [
      { to: '/manage', label: 'Manage', icon: '🗂️' },
      { to: '/govern', label: 'Govern', icon: '⚖️' },
      { to: '/protect', label: 'Protect', icon: '🛡️' },
    ],
  },
];

interface AppShellProps {
  children?: ReactNode;
}

export function AppShell({ children }: AppShellProps) {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <AuthBar />
      <div className={styles.layout}>
        <aside className={styles.rail}>
          <div className={styles.brand}>
            <div className={styles.brandMark}>G4</div>
            <div className={styles.brandText}>
              <Subtitle2 className={styles.brandTitle}>GOV4Agents</Subtitle2>
              <Caption1 className={styles.brandSub}>Manage · Govern · Protect</Caption1>
            </div>
          </div>
          {groups.map((g) => (
            <div className={styles.group} key={g.header}>
              <Caption1 className={styles.groupHeader}>{g.header}</Caption1>
              {g.items.map((item) => (
                <NavLink
                  key={item.to}
                  to={item.to}
                  end={item.end}
                  className={({ isActive }) =>
                    `${styles.link} ${isActive ? styles.active : ''}`
                  }
                >
                  <span className={styles.icon}>{item.icon}</span>
                  <span>{item.label}</span>
                </NavLink>
              ))}
            </div>
          ))}
        </aside>
        <main className={styles.content}>{children ?? <Outlet />}</main>
      </div>
    </div>
  );
}
