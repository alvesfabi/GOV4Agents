import { Body1, Subtitle2, makeStyles } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: '12px' },
});

interface StubProps {
  title: string;
  description: string;
}

export function Stub({ title, description }: StubProps) {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <Subtitle2>{title}</Subtitle2>
      <Body1>{description}</Body1>
      <Body1>This page will be implemented in the next phase.</Body1>
    </div>
  );
}
