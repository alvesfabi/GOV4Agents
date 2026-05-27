import type { ReactNode } from 'react';
import { makeStyles, tokens, Title2, Body1, Subtitle2 } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
  },
  banner: {
    width: '100%',
    maxHeight: '220px',
    objectFit: 'cover',
    borderRadius: tokens.borderRadiusMedium,
  },
  body: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
  },
  actions: {
    display: 'flex',
    gap: '8px',
    flexWrap: 'wrap',
  },
});

interface StepShellProps {
  title: string;
  intro?: ReactNode;
  banner?: string;
  children: ReactNode;
  footer?: ReactNode;
}

export function StepShell({ title, intro, banner, children, footer }: StepShellProps) {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      {banner && <img src={banner} alt="" className={styles.banner} />}
      <Title2>{title}</Title2>
      {intro && <Body1>{intro}</Body1>}
      <Subtitle2 as="h3">Actions</Subtitle2>
      <div className={styles.body}>{children}</div>
      {footer && <div className={styles.actions}>{footer}</div>}
    </div>
  );
}
