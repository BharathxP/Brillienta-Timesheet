import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, DialogType, Icon, MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, TextField, TooltipHost } from '@fluentui/react';
import { initials } from '../../services/format';
import { useApp } from '../AppContext';
import { Feature, featurePlan, planInfo } from '../../services/licensing/plans';
import { ROLE_LABELS } from '../../services/access';
import { Role } from '../../models/types';

function colorIndex(s: string): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
  return (Math.abs(h) % 6) + 1;
}

/** Initials avatar, with the Microsoft 365 photo from SharePoint when there is one. */
export function Avatar(props: { name: string; email?: string; size?: 'sm' | 'md' | 'lg' }): JSX.Element {
  const { sp } = useApp();
  const [state, setState] = React.useState<'loading' | 'ok' | 'none'>('loading');
  const size = props.size || 'md';
  const cls = `ts-avatar ts-av-${colorIndex(props.name || '?')}` + (size === 'md' ? '' : ` is-${size}`);
  const src = props.email ? `${sp.webUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(props.email)}` : '';
  return (
    <span className={cls} title={props.name}>
      <span className="ts-avatar-inner">
        {state !== 'ok' && initials(props.name)}
        {src && state !== 'none' && (
          <img
            src={src}
            alt=""
            style={state === 'ok' ? undefined : { position: 'absolute', opacity: 0, width: 1, height: 1 }}
            onLoad={(e) => setState(e.currentTarget.naturalWidth > 0 && e.currentTarget.naturalWidth < 20 ? 'none' : 'ok')}
            onError={() => setState('none')}
          />
        )}
      </span>
    </span>
  );
}

export function Person(props: { name: string; email?: string; sub?: string }): JSX.Element {
  return (
    <span className="ts-person">
      <Avatar name={props.name} email={props.email} size="sm" />
      <span>
        {props.name}
        {props.sub && <span className="ts-muted ts-small"> · {props.sub}</span>}
      </span>
    </span>
  );
}

export function Empty(props: { icon?: string; title: string; children?: React.ReactNode }): JSX.Element {
  return (
    <div className="ts-empty">
      <Icon iconName={props.icon || 'Clock'} />
      <div style={{ fontWeight: 650, color: 'var(--ts-text)' }}>{props.title}</div>
      {props.children && <div style={{ marginTop: 6 }}>{props.children}</div>}
    </div>
  );
}

export function Loading(props: { label?: string }): JSX.Element {
  return (
    <div className="ts-center" style={{ minHeight: 200 }}>
      <Spinner size={SpinnerSize.large} label={props.label || 'Loading…'} />
    </div>
  );
}

export function ErrorBox(props: { error?: string; onRetry?: () => void }): JSX.Element | null {
  if (!props.error) return null;
  return (
    <MessageBar messageBarType={MessageBarType.error} isMultiline actions={props.onRetry ? <DefaultButton text="Retry" onClick={props.onRetry} /> : undefined}>
      {props.error}
    </MessageBar>
  );
}

export function PageHeader(props: { title: React.ReactNode; sub?: React.ReactNode; actions?: React.ReactNode; back?: () => void; eyebrow?: string }): JSX.Element {
  return (
    <div className="ts-page-header">
      {props.back && <DefaultButton iconProps={{ iconName: 'Back' }} onClick={props.back} styles={{ root: { minWidth: 32, padding: 0 } }} ariaLabel="Back" />}
      <div>
        {props.eyebrow && (
          <p className="ts-eyebrow">
            <span className="ts-star">✳</span> {props.eyebrow}
          </p>
        )}
        <h1>{props.title}</h1>
        {props.sub && <div className="ts-sub">{props.sub}</div>}
      </div>
      {props.actions && <div className="ts-page-actions">{props.actions}</div>}
    </div>
  );
}

export function Kpi(props: { label: string; value: React.ReactNode; hint?: string; icon?: string; tint?: number; tone?: 'danger' | 'warning' | 'success'; onClick?: () => void }): JSX.Element {
  return (
    <div
      className={'ts-kpi' + (props.onClick ? ' is-link' : '') + (props.tone ? ' is-' + props.tone : '')}
      onClick={props.onClick}
      onKeyDown={(e) => props.onClick && (e.key === 'Enter' || e.key === ' ') && props.onClick()}
      role={props.onClick ? 'button' : undefined}
      tabIndex={props.onClick ? 0 : undefined}
    >
      <div className="ts-kpi-icon">
        <div className="ts-kpi-label">{props.label}</div>
        {props.icon && (
          <span className={`ts-icon-tile is-sm ts-tint-${props.tint || 1}`}>
            <Icon iconName={props.icon} />
          </span>
        )}
      </div>
      <div className="ts-kpi-value">{props.value}</div>
      {props.hint && <div className="ts-kpi-hint">{props.hint}</div>}
    </div>
  );
}

export function Card(props: { title?: React.ReactNode; icon?: string; actions?: React.ReactNode; children: React.ReactNode; className?: string; style?: React.CSSProperties }): JSX.Element {
  return (
    <div className={'ts-card ' + (props.className || '')} style={props.style}>
      {props.title && (
        <h3 className="ts-card-title">
          {props.icon && <Icon iconName={props.icon} />}
          {props.title}
          {props.actions && <span className="ts-card-actions">{props.actions}</span>}
        </h3>
      )}
      {props.children}
    </div>
  );
}

export function Segmented<T extends string>(props: { value: T; options: Array<{ key: T; text: string; icon?: string }>; onChange: (v: T) => void; label?: string }): JSX.Element {
  return (
    <div className="ts-seg" role="group" aria-label={props.label}>
      {props.options.map((o) => (
        <button key={o.key} className={o.key === props.value ? 'is-active' : ''} aria-pressed={o.key === props.value} onClick={() => props.onChange(o.key)} title={o.text}>
          {o.icon && <Icon iconName={o.icon} />}
          {o.text}
        </button>
      ))}
    </div>
  );
}

/** Tabs drawn as view chips (the Brillienta design system). */
export function Tabs<T extends string>(props: { value: T; tabs: Array<{ key: T; text: string; icon?: string; count?: number; locked?: boolean }>; onChange: (v: T) => void }): JSX.Element {
  return (
    <div className="ts-views" role="tablist">
      {props.tabs.map((t) => (
        <button key={t.key} role="tab" aria-selected={t.key === props.value} className={'ts-view-chip' + (t.key === props.value ? ' is-active' : '')} onClick={() => props.onChange(t.key)}>
          {t.icon && <Icon iconName={t.icon} />}
          {t.text}
          {t.count !== undefined && <span className="ts-view-count">{t.count}</span>}
          {t.locked && <Icon iconName="Lock" />}
        </button>
      ))}
    </div>
  );
}

export function StatusBadge(props: { status: string; title?: string }): JSX.Element {
  return (
    <span className={'ts-status is-' + props.status.toLowerCase().replace(/\s+/g, '-')} title={props.title}>
      {props.status}
    </span>
  );
}

export function RoleBadge(props: { role: Role }): JSX.Element {
  return <span className="ts-chip is-soft">{ROLE_LABELS[props.role]}</span>;
}

export function Progress(props: { value: number; max: number; label?: string }): JSX.Element {
  const pct = props.max > 0 ? Math.min(100, (props.value / props.max) * 100) : 0;
  const over = props.max > 0 && props.value > props.max;
  return (
    <TooltipHost content={props.label || `${Math.round((props.value / (props.max || 1)) * 100)}%`}>
      <div className={'ts-progress' + (over ? ' is-over' : pct > 85 ? ' is-warn' : '')}>
        <span style={{ width: `${over ? 100 : pct}%` }} />
      </div>
    </TooltipHost>
  );
}

/** Shown in place of a feature the plan does not include. */
export function Locked(props: { feature: Feature; children?: React.ReactNode }): JSX.Element {
  const { navigate, isAdmin } = useApp();
  const plan = planInfo(featurePlan(props.feature));
  return (
    <div className="ts-card" style={{ display: 'flex', gap: 14, alignItems: 'center', background: 'var(--ts-surface-2)' }}>
      <span className="ts-icon-tile ts-tint-4">
        <Icon iconName="Lock" />
      </span>
      <div style={{ flex: 1 }}>
        <strong>Available on the {plan.name} plan</strong>
        <div className="ts-muted ts-small">{props.children || plan.tagline}</div>
      </div>
      {isAdmin && <DefaultButton text="See plans" onClick={() => navigate('admin/licence')} />}
    </div>
  );
}

interface IConfirmState {
  title: string;
  message: string;
  confirmText: string;
  withInput?: string;
  required?: boolean;
  resolve: (value: string | false) => void;
}

/** Promise-based confirm / prompt dialog. Resolves the typed text ('ok' without input) or false. */
export function useConfirm(): [(title: string, message: string, confirmText?: string, withInput?: string, required?: boolean) => Promise<string | false>, JSX.Element | null] {
  const [state, setState] = React.useState<IConfirmState | undefined>();
  const [text, setText] = React.useState('');
  const ask = React.useCallback(
    (title: string, message: string, confirmText: string = 'OK', withInput?: string, required?: boolean): Promise<string | false> =>
      new Promise((resolve) => {
        setText('');
        setState({ title, message, confirmText, withInput, required, resolve });
      }),
    []
  );
  const close = (v: string | false): void => {
    if (state) state.resolve(v);
    setState(undefined);
  };
  const dialog = state ? (
    <Dialog hidden={false} onDismiss={() => close(false)} dialogContentProps={{ type: DialogType.normal, title: state.title, subText: state.message }} modalProps={{ isBlocking: true }} minWidth={420}>
      {state.withInput !== undefined && <TextField label={state.withInput} multiline rows={3} value={text} onChange={(_, v) => setText(v || '')} autoFocus />}
      <DialogFooter>
        <PrimaryButton text={state.confirmText} disabled={!!state.required && !text.trim()} onClick={() => close(text || 'ok')} />
        <DefaultButton text="Cancel" onClick={() => close(false)} />
      </DialogFooter>
    </Dialog>
  ) : null;
  return [ask, dialog];
}

/** Async loader hook with reload and error state. */
export function useLoader<T>(load: () => Promise<T>, deps: React.DependencyList): { data?: T; error?: string; loading: boolean; reload: () => void; setData: (d: T) => void } {
  const [data, setData] = React.useState<T | undefined>();
  const [error, setError] = React.useState<string | undefined>();
  const [loading, setLoading] = React.useState(true);
  const [tick, setTick] = React.useState(0);
  React.useEffect(() => {
    let alive = true;
    setLoading(true);
    setError(undefined);
    load()
      .then((d) => alive && setData(d))
      .catch((e: unknown) => alive && setError(e instanceof Error ? e.message : String(e)))
      .finally(() => alive && setLoading(false));
    return () => {
      alive = false;
    };
  }, deps.concat([tick]));
  return { data, error, loading, reload: () => setTick((t) => t + 1), setData };
}

/** Run an async action with a busy flag and an error toast. */
export function useBusy(): [boolean, (fn: () => Promise<void>, success?: string) => Promise<void>] {
  const { toast } = useApp();
  const [busy, setBusy] = React.useState(false);
  const run = React.useCallback(
    async (fn: () => Promise<void>, success?: string): Promise<void> => {
      setBusy(true);
      try {
        await fn();
        if (success) toast(success, 'success');
      } catch (e) {
        toast(e instanceof Error ? e.message : String(e), 'error');
      } finally {
        setBusy(false);
      }
    },
    [toast]
  );
  return [busy, run];
}

export function Field(props: { label: string; children: React.ReactNode; hint?: string; wide?: boolean }): JSX.Element {
  return (
    <div className={props.wide ? 'is-wide' : undefined}>
      <div className="ms-Label" style={{ padding: '5px 0' }}>
        {props.label}
      </div>
      {props.children}
      {props.hint && <div className="ts-muted ts-small" style={{ marginTop: 4 }}>{props.hint}</div>}
    </div>
  );
}

export function ColorInput(props: { value: string; onChange: (v: string) => void; label?: string }): JSX.Element {
  return <input type="color" className="ts-color" value={/^#[0-9a-f]{6}$/i.test(props.value) ? props.value : '#2a78d6'} onChange={(e) => props.onChange(e.target.value)} aria-label={props.label || 'Colour'} />;
}
