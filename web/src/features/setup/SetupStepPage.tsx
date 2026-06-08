import { useState, useEffect, type ReactNode } from 'react';
import {
  Button,
  Field,
  Input,
  Spinner,
  makeStyles,
  Subtitle2,
  tokens,
  Body1,
} from '@fluentui/react-components';
import { useApi, ApiError } from '../../api/client';
import { JsonViewer } from '../../components/JsonViewer';
import { ErrorPanel } from '../../components/ErrorPanel';
import { useDemoSuffix } from './DemoSuffixContext';

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: '16px' },
  banner: {
    width: '100%',
    maxHeight: '180px',
    objectFit: 'cover',
    borderRadius: tokens.borderRadiusMedium,
  },
  intro: { color: tokens.colorNeutralForeground2 },
  form: {
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    maxWidth: '520px',
    padding: '16px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  submit: { alignSelf: 'flex-start', marginTop: '4px' },
  result: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    backgroundColor: tokens.colorPaletteGreenBackground1,
    border: `1px solid ${tokens.colorPaletteGreenBorder1}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: '12px 16px',
  },
  resultHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    color: tokens.colorPaletteGreenForeground1,
  },
});

export interface FieldDef {
  name: string;
  label: string;
  helper?: ReactNode;
  defaultValue?: string;
  /** If set, the default value is `${defaultPrefix}${suffix}` where suffix comes from context. */
  defaultPrefix?: string;
  required?: boolean;
}

interface SetupStepPageProps {
  title: string;
  intro: ReactNode;
  banner?: string;
  fields?: FieldDef[];
  endpoint: string;
  submitLabel?: string;
  renderResult?: (result: unknown) => ReactNode;
}

export function SetupStepPage({
  intro,
  banner,
  fields = [],
  endpoint,
  submitLabel = 'Create',
  renderResult,
}: SetupStepPageProps) {
  const styles = useStyles();
  const apiFetch = useApi();
  const { suffix } = useDemoSuffix();

  function computeDefault(f: FieldDef): string {
    if (f.defaultPrefix && suffix) return `${f.defaultPrefix}${suffix}`;
    return f.defaultValue ?? '';
  }

  const [values, setValues] = useState<Record<string, string>>(() =>
    Object.fromEntries(fields.map((f) => [f.name, computeDefault(f)])),
  );

  // Update defaults when suffix changes (only fields the user hasn't manually edited)
  useEffect(() => {
    setValues((prev) => {
      const next = { ...prev };
      for (const f of fields) {
        if (f.defaultPrefix) {
          const currentVal = prev[f.name] ?? '';
          // Update if empty or still matches a previous auto-generated value
          if (!currentVal || currentVal.startsWith(f.defaultPrefix)) {
            next[f.name] = `${f.defaultPrefix}${suffix}`;
          }
        }
      }
      return next;
    });
  }, [suffix, fields]);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<unknown>(null);
  const [error, setError] = useState<unknown>(null);

  async function submit() {
    setBusy(true);
    setResult(null);
    setError(null);
    try {
      const data = await apiFetch(endpoint, {
        method: 'POST',
        body: JSON.stringify(values),
      });
      setResult(data);
    } catch (err) {
      setError(err);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className={styles.root}>
      {banner && <img src={banner} alt="" className={styles.banner} />}
      <Body1 className={styles.intro}>{intro}</Body1>
      {fields.length > 0 ? (
        <div className={styles.form}>
          {fields.map((f) => (
            <Field
              key={f.name}
              label={f.label}
              hint={f.helper ? <>{f.helper}</> : undefined}
              required={f.required}
            >
              <Input
                value={values[f.name] ?? ''}
                onChange={(_e, data) => setValues((v) => ({ ...v, [f.name]: data.value }))}
              />
            </Field>
          ))}
          <Button
            appearance="primary"
            className={styles.submit}
            disabled={busy || fields.some((f) => f.required && !values[f.name])}
            onClick={submit}
          >
            {busy ? <Spinner size="tiny" /> : submitLabel}
          </Button>
        </div>
      ) : (
        <Button
          appearance="primary"
          className={styles.submit}
          disabled={busy}
          onClick={submit}
        >
          {busy ? <Spinner size="tiny" /> : submitLabel}
        </Button>
      )}
      {error != null && <ErrorPanel error={error} />}
      {result != null && (
        <div className={styles.result}>
          <div className={styles.resultHeader}>
            <Subtitle2>✓ Created</Subtitle2>
          </div>
          {renderResult ? renderResult(result) : null}
          <JsonViewer value={result} label="Raw response" />
        </div>
      )}
    </div>
  );
}

export { ApiError };
