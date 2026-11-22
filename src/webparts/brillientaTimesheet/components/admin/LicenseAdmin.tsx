import * as React from 'react';
import { MessageBar, MessageBarType, PrimaryButton, TextField } from '@fluentui/react';
import { ILicenseStatus, checkLicenseKey, describeStatus } from '../../services/licensing/license';
import { LICENSE_PUBLIC_KEY } from '../../services/licensing/publicKey';
import { FEATURES, PLANS, Plan } from '../../services/licensing/plans';
import { SETTING_KEYS, SettingsStore } from '../../services/SettingsStore';
import { DEFAULT_LOGO } from '../../assets/defaultLogo';

/** Enter and check a licence key (signature and tenant), then store it. */
export function KeyEntry(props: { tenantId?: string; settings: SettingsStore; onChanged: () => void }): JSX.Element {
  const [key, setKey] = React.useState('');
  const [error, setError] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const activate = async (): Promise<void> => {
    setBusy(true);
    setError('');
    try {
      const p = await checkLicenseKey(key, props.tenantId, LICENSE_PUBLIC_KEY);
      if (p.expires && Date.parse(p.expires + 'T23:59:59Z') < Date.now()) throw new Error(`This licence ended on ${p.expires}.`);
      await props.settings.set(SETTING_KEYS.licenseKey, key.replace(/\s+/g, ''));
      props.onChanged();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };
  return (
    <div className="ts-stack">
      <TextField label="Licence key" multiline rows={3} value={key} onChange={(_, v) => setKey(v || '')} placeholder="BT1.…" styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }} />
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      <div className="ts-row">
        <PrimaryButton text="Activate" onClick={() => void activate()} disabled={busy || !key.trim()} />
        <span className="ts-muted ts-small">
          Tenant id: <code>{props.tenantId || 'unknown'}</code>
        </span>
      </div>
    </div>
  );
}

export function PlanCards(props: { current?: Plan }): JSX.Element {
  return (
    <div className="ts-plans">
      {PLANS.map((p) => (
        <div key={p.plan} className={'ts-plan' + (props.current === p.plan ? ' is-current' : '')}>
          <h3>{p.name}</h3>
          <div className="ts-muted ts-small">{p.tagline}</div>
          <div className="ts-small" style={{ marginTop: 8 }}>
            <strong>{p.users} users</strong>
            {p.extraUserPrice ? ` · extra users ₹${p.extraUserPrice}/user` : ''} · {p.tickets ? `${p.tickets} support tickets a year` : 'unlimited support'}
          </div>
          <ul>
            {FEATURES.filter((f) => f.plan === p.plan).map((f) => (
              <li key={f.feature}>{f.label}</li>
            ))}
          </ul>
          {p.plan !== 'standard' && <div className="ts-muted ts-small">…and everything in {p.plan === 'pro' ? 'Standard' : 'Pro'}.</div>}
        </div>
      ))}
    </div>
  );
}

/** Shown instead of the app when the trial or the licence has ended. */
export function LicenseLocked(props: { status: ILicenseStatus; isAdmin: boolean; tenantId?: string; settings: SettingsStore; onChanged: () => void }): JSX.Element {
  return (
    <div className="ts-setup">
      <img className="ts-logo is-lg" src={DEFAULT_LOGO} alt="" />
      <h1>{describeStatus(props.status)}</h1>
      <p className="ts-muted">Your timesheets, projects and reports are safe in your SharePoint lists. {props.isAdmin ? 'Enter a licence key from Brillienta to continue.' : 'Ask your timesheet administrator to activate a licence.'}</p>
      {props.status.keyProblem && <MessageBar messageBarType={MessageBarType.warning}>The stored key was not accepted: {props.status.keyProblem}</MessageBar>}
      {props.isAdmin && <KeyEntry tenantId={props.tenantId} settings={props.settings} onChanged={props.onChanged} />}
      <div style={{ marginTop: 24 }}>
        <PlanCards />
      </div>
    </div>
  );
}
