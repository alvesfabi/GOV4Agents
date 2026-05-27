import { makeStyles, tokens, Caption1 } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    padding: '8px 12px',
    overflowX: 'auto',
  },
  pre: {
    margin: 0,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '12px',
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
  },
});

interface JsonViewerProps {
  value: unknown;
  label?: string;
}

export function JsonViewer({ value, label }: JsonViewerProps) {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      {label && <Caption1>{label}</Caption1>}
      <pre className={styles.pre}>{stringify(value)}</pre>
    </div>
  );
}

function stringify(value: unknown): string {
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}
