import { useState, useEffect } from 'react';
import {
  Field,
  Input,
  Select,
  Button,
  Spinner,
  Subtitle2,
  Body1,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { useApi } from '../../api/client';
import { useDemoSuffix } from './DemoSuffixContext';
import { JsonViewer } from '../../components/JsonViewer';
import { ErrorPanel } from '../../components/ErrorPanel';
import { PortalLink } from '../../components/Links';
import { portal } from '../../lib/portalLinks';

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: '16px' },
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
  resultHeader: { color: tokens.colorPaletteGreenForeground1 },
});

interface AttributeSet {
  id: string;
  description?: string;
}

interface CsaResult {
  caPolicy?: { id?: string };
  csa?: { setName?: string; usingExistingSet?: boolean };
}

const POLICY_PREFIX = 'Default deny agents except approved (Agent ID only) ';

export function ConditionalAccessPage() {
  const styles = useStyles();
  const apiFetch = useApi();
  const { suffix } = useDemoSuffix();

  const [policyName, setPolicyName] = useState(`${POLICY_PREFIX}${suffix}`);
  const [policyEdited, setPolicyEdited] = useState(false);
  const [sets, setSets] = useState<AttributeSet[]>([]);
  const [setsError, setSetsError] = useState<string | null>(null);
  const [selectedSet, setSelectedSet] = useState(''); // '' = create a new set

  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<CsaResult | null>(null);
  const [error, setError] = useState<unknown>(null);

  // Keep the policy name default in sync with the suffix unless the user edited it.
  useEffect(() => {
    if (!policyEdited) setPolicyName(`${POLICY_PREFIX}${suffix}`);
  }, [suffix, policyEdited]);

  // Load existing attribute sets so the user can reuse one.
  useEffect(() => {
    apiFetch<{ sets: AttributeSet[] }>('/api/setup/attribute-sets')
      .then((r) => setSets(r.sets ?? []))
      .catch((err) => setSetsError(err instanceof Error ? err.message : String(err)));
  }, [apiFetch]);

  async function submit() {
    setBusy(true);
    setResult(null);
    setError(null);
    try {
      const data = await apiFetch<CsaResult>('/api/setup/csa-and-ca', {
        method: 'POST',
        body: JSON.stringify({
          name: policyName,
          existingSetName: selectedSet || undefined,
        }),
      });
      setResult(data);
    } catch (err) {
      setError(err);
    } finally {
      setBusy(false);
    }
  }

  const safeSuffix = suffix.replace(/[^a-zA-Z0-9]/g, '');
  const newSetLabel = `Create a new set (AgentsCSA${safeSuffix})`;

  return (
    <div className={styles.root}>
      <Body1 className={styles.intro}>
        Creates the <strong>AgentsCSA / TAG</strong> custom security attribute (with allowed
        values <code>approved</code> and <code>notApproved</code>) and the{' '}
        <strong>Block not approved agents</strong> Conditional Access policy. The policy targets
        service principals whose <code>TAG</code> attribute is not <code>approved</code>. It is
        created in <strong>report-only</strong> mode — the Protect journey will turn it on later.
        <br />
        <br />
        Choose <em>Create a new set</em> (needs the directory-scoped{' '}
        <strong>Attribute Definition Administrator</strong> role) or reuse an{' '}
        <strong>existing set</strong> (works with that role scoped just to the chosen set).
      </Body1>

      <div className={styles.form}>
        <Field
          label="Attribute set"
          hint={
            setsError
              ? `Could not list existing sets (${setsError}). A new set will be created.`
              : 'Reuse an existing set to avoid the directory-wide Attribute Definition Administrator role.'
          }
        >
          <Select value={selectedSet} onChange={(_e, data) => setSelectedSet(data.value)}>
            <option value="">{newSetLabel}</option>
            {sets.map((s) => (
              <option key={s.id} value={s.id}>
                {s.id}
                {s.description ? ` — ${s.description}` : ''}
              </option>
            ))}
          </Select>
        </Field>

        <Field label="Policy name" required>
          <Input
            value={policyName}
            onChange={(_e, data) => {
              setPolicyEdited(true);
              setPolicyName(data.value);
            }}
          />
        </Field>

        <Button
          appearance="primary"
          className={styles.submit}
          disabled={busy || !policyName.trim()}
          onClick={submit}
        >
          {busy ? <Spinner size="tiny" /> : 'Create CSA & policy'}
        </Button>
      </div>

      {error != null && <ErrorPanel error={error} />}

      {result != null && (
        <div className={styles.result}>
          <Subtitle2 className={styles.resultHeader}>✓ Created</Subtitle2>
          {result.csa?.usingExistingSet ? (
            <Body1>
              Reused existing attribute set <strong>{result.csa.setName}</strong>.
            </Body1>
          ) : (
            <Body1>
              Created new attribute set <strong>{result.csa?.setName}</strong>.
            </Body1>
          )}
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            {result.caPolicy?.id && (
              <PortalLink href={portal.conditionalAccessPolicy(result.caPolicy.id)}>
                Open Conditional Access policy
              </PortalLink>
            )}
            {result.csa?.setName && (
              <PortalLink href={portal.csaSet(result.csa.setName)}>
                Open custom attribute set
              </PortalLink>
            )}
          </div>
          <JsonViewer value={result} label="Raw response" />
        </div>
      )}
    </div>
  );
}
