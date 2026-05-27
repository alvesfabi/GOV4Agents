import type { ReactNode } from 'react';
import { makeStyles, tokens, Title1, Title3, Body1 } from '@fluentui/react-components';

const useStyles = makeStyles({
  page: {
    display: 'flex',
    flexDirection: 'column',
    gap: '24px',
    maxWidth: '1100px',
    margin: '0 auto',
  },
  hero: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    padding: '28px 32px',
    borderRadius: tokens.borderRadiusXLarge,
    background: `linear-gradient(135deg, ${tokens.colorBrandBackground2}, ${tokens.colorNeutralBackground1})`,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  heroSub: { color: tokens.colorNeutralForeground2, maxWidth: '760px' },
  heroBanner: {
    width: '100%',
    maxHeight: '320px',
    objectFit: 'contain',
    borderRadius: tokens.borderRadiusXLarge,
    marginTop: '8px',
    boxShadow: tokens.shadow8,
  },
  step: {
    display: 'grid',
    gridTemplateColumns: '56px 1fr',
    gap: '20px',
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusXLarge,
    padding: '28px 32px',
    boxShadow: tokens.shadow4,
  },
  badge: {
    width: '44px',
    height: '44px',
    borderRadius: '50%',
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontWeight: 700,
    fontSize: tokens.fontSizeBase500,
    flexShrink: 0,
  },
  stepBody: { display: 'flex', flexDirection: 'column', gap: '12px', minWidth: 0 },
  stepLabel: {
    color: tokens.colorBrandForeground1,
    fontSize: tokens.fontSizeBase200,
    fontWeight: 600,
    letterSpacing: '0.08em',
    textTransform: 'uppercase',
  },
  card: {
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusXLarge,
    padding: '24px 28px',
    boxShadow: tokens.shadow4,
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
});

interface PageProps {
  children: ReactNode;
}

export function Page({ children }: PageProps) {
  const styles = useStyles();
  return <div className={styles.page}>{children}</div>;
}

interface PageHeroProps {
  title: string;
  subtitle?: ReactNode;
  banner?: string;
  bannerAlt?: string;
}

export function PageHero({ title, subtitle, banner, bannerAlt = '' }: PageHeroProps) {
  const styles = useStyles();
  return (
    <div className={styles.hero}>
      <Title1>{title}</Title1>
      {subtitle && <Body1 className={styles.heroSub}>{subtitle}</Body1>}
      {banner && <img src={banner} alt={bannerAlt} className={styles.heroBanner} />}
    </div>
  );
}

interface StepCardProps {
  number: number;
  label: string;
  children: ReactNode;
}

export function StepCard({ number, label, children }: StepCardProps) {
  const styles = useStyles();
  return (
    <section className={styles.step}>
      <div className={styles.badge}>{number}</div>
      <div className={styles.stepBody}>
        <span className={styles.stepLabel}>Step {number}</span>
        <Title3>{label}</Title3>
        {children}
      </div>
    </section>
  );
}

interface CardProps {
  children: ReactNode;
}

export function Card({ children }: CardProps) {
  const styles = useStyles();
  return <div className={styles.card}>{children}</div>;
}
