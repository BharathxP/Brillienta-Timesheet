import * as React from 'react';
import { Checkbox, ChoiceGroup, DefaultButton, Dropdown, Icon, MessageBar, MessageBarType, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { CostMode, ICustomColumn, IHourRule, ITimesheetConfig, ROLES, Role } from '../../models/types';
import { Feature } from '../../services/licensing/plans';
import { ROLE_LABELS } from '../../services/access';
import { CURRENCIES } from '../../services/format';
import { DAY_NAMES, DAY_SHORT, formatDay } from '../../services/dates';
import { Card, ColorInput, DateField, Field, Locked, PageHeader, useBusy } from '../common/Ui';
import { PeopleField } from '../common/Pickers';
import { NotificationsAdmin, IntegrationsAdmin, LicenceAdmin, SecurityAdmin } from './AdminMore';

const SECTIONS: Array<{ key: string; text: string; icon: string }> = [
  { key: 'general', text: 'General & branding', icon: 'Settings' },
  { key: 'tracking', text: 'Week & tracking days', icon: 'Calendar' },
  { key: 'hours', text: 'Hour restrictions', icon: 'Clock' },
  { key: 'approvals', text: 'Approvals', icon: 'DocumentApproval' },
  { key: 'costs', text: 'Costs & rates', icon: 'Money' },
  { key: 'columns', text: 'Custom columns', icon: 'TripleColumn' },
  { key: 'notifications', text: 'Notifications & email', icon: 'Mail' },
  { key: 'integrations', text: 'Integrations', icon: 'PlugConnected' },
  { key: 'licence', text: 'Licence & plan', icon: 'Certificate' },
  { key: 'security', text: 'Schema & security', icon: 'Lock' }
];

/** Edits a draft of the saved settings; Save writes it (locked features keep their saved values). */
export function useDraft(): { d: ITimesheetConfig; set: (patch: Partial<ITimesheetConfig>) => void; dirty: boolean; save: (what: string) => Promise<void>; busy: boolean; reset: () => void } {
  const { savedConfig, saveConfig } = useApp();
  const [d, setD] = React.useState<ITimesheetConfig>(savedConfig);
  const [dirty, setDirty] = React.useState(false);
  const [busy, run] = useBusy();
  React.useEffect(() => {
    if (!dirty) setD(savedConfig);
  }, [savedConfig]);
  return {
    d,
    set: (patch) => {
      setD((cur) => ({ ...cur, ...patch }));
      setDirty(true);
    },
    dirty,
    busy,
    reset: () => {
      setD(savedConfig);
      setDirty(false);
    },
    save: (what) =>
      run(async () => {
        await saveConfig(d, what);
        setDirty(false);
      }, 'Settings saved.')
  };
}

export function SaveBar(props: { draft: ReturnType<typeof useDraft>; what: string }): JSX.Element {
  return (
    <div className="ts-row" style={{ marginTop: 16 }}>
      <PrimaryButton text="Save" onClick={() => void props.draft.save(props.what)} disabled={!props.draft.dirty || props.draft.busy} />
      <DefaultButton text="Undo changes" onClick={props.draft.reset} disabled={!props.draft.dirty || props.draft.busy} />
      {props.draft.dirty && <span className="ts-muted ts-small">Unsaved changes</span>}
    </div>
  );
}

/** A setting that needs a plan feature: shown with a lock when the plan does not include it. */
export function Gate(props: { feature: Feature; children: React.ReactNode; text?: string }): JSX.Element {
  const { has } = useApp();
  return has(props.feature) ? <>{props.children}</> : <Locked feature={props.feature}>{props.text}</Locked>;
}

export function AdminPage(props: { section?: string }): JSX.Element {
  const { navigate, isAdmin } = useApp();
  const section = SECTIONS.some((s) => s.key === props.section) ? (props.section as string) : 'general';
  if (!isAdmin) return <Locked feature="roles">Only administrators can change settings.</Locked>;
  let body: JSX.Element;
  switch (section) {
    case 'tracking':
      body = <TrackingAdmin />;
      break;
    case 'hours':
      body = <HoursAdmin />;
      break;
    case 'approvals':
      body = <ApprovalsAdmin />;
      break;
    case 'costs':
      body = <CostsAdmin />;
      break;
    case 'columns':
      body = <ColumnsAdmin />;
      break;
    case 'notifications':
      body = <NotificationsAdmin />;
      break;
    case 'integrations':
      body = <IntegrationsAdmin />;
      break;
    case 'licence':
      body = <LicenceAdmin />;
      break;
    case 'security':
      body = <SecurityAdmin />;
      break;
    default:
      body = <GeneralAdmin />;
  }
  return (
    <>
      <PageHeader title="Admin" sub="Settings for everyone in your organisation." />
      <div className="ts-admin">
        <nav className="ts-admin-nav" aria-label="Admin sections">
          {SECTIONS.map((s) => (
            <button key={s.key} className={'ts-nav-item' + (s.key === section ? ' is-active' : '')} onClick={() => navigate(`admin/${s.key}`)}>
              <Icon iconName={s.icon} />
              <span>{s.text}</span>
            </button>
          ))}
        </nav>
        <div style={{ minWidth: 0 }}>{body}</div>
      </div>
    </>
  );
}

function GeneralAdmin(): JSX.Element {
  const draft = useDraft();
  const { d, set } = draft;
  const { logo, saveLogo, toast } = useApp();
  const upload = (file: File): void => {
    if (file.size > 300000) {
      toast('Use an image under 300 KB.', 'warning');
      return;
    }
    const r = new FileReader();
    r.onload = () =>
      void saveLogo(String(r.result))
        .then(() => toast('Logo updated.', 'success'))
        .catch((e) => toast((e as Error).message, 'error'));
    r.readAsDataURL(file);
  };
  return (
    <Card title="General & branding" icon="Settings">
      <div className="ts-form-grid">
        <TextField label="Organisation name" value={d.orgName} onChange={(_, v) => set({ orgName: v || '' })} />
        <TextField label="App name" value={d.appName} onChange={(_, v) => set({ appName: v || 'Brillienta Timesheet' })} />
        <Field label="Brand colour" hint="Buttons, navigation highlights and emails.">
          <div className="ts-row">
            <ColorInput value={d.brandColor} onChange={(v) => set({ brandColor: v })} />
            <code>{d.brandColor}</code>
            <button className="ts-link ts-small" onClick={() => set({ brandColor: '#17232e' })}>
              Brillienta navy
            </button>
          </div>
        </Field>
        <Dropdown label="Currency" selectedKey={d.currency} options={CURRENCIES.map((c) => ({ key: c, text: c }))} onChange={(_, o) => o && set({ currency: String(o.key) })} />
        <Field label="Logo">
          <div className="ts-row">
            <img className="ts-logo" src={logo} alt="" />
            <input type="file" accept="image/png,image/jpeg,image/svg+xml,image/webp" onChange={(e) => e.target.files && e.target.files[0] && upload(e.target.files[0])} />
            <button className="ts-link ts-small" onClick={() => void saveLogo('')}>
              Default logo
            </button>
          </div>
        </Field>
        <div />
        <PeopleField label="Additional administrators" value={d.admins} onChange={(v) => set({ admins: v })} hint="Site owners and people with the Administrator role are administrators too." />
        <div>
          <Dropdown label="Role for new people" selectedKey={d.defaultRole} options={ROLES.filter((r) => r !== 'admin').map((r) => ({ key: r, text: ROLE_LABELS[r] }))} onChange={(_, o) => o && set({ defaultRole: o.key as Role })} />
          <Toggle label="Add people automatically on their first visit" checked={d.autoEnrol} onChange={(_, v) => set({ autoEnrol: !!v })} styles={{ root: { marginTop: 12 } }} />
        </div>
      </div>
      <SaveBar draft={draft} what="General" />
    </Card>
  );
}

function TrackingAdmin(): JSX.Element {
  const draft = useDraft();
  const { d, set } = draft;
  const { has, savedConfig } = useApp();
  const savedWeekStart = savedConfig.weekStart;
  const [hDate, setHDate] = React.useState('');
  const [hName, setHName] = React.useState('');
  return (
    <div className="ts-stack">
      <Card title="Week" icon="Calendar">
        <div className="ts-form-grid">
          <Gate feature="weekStart">
            <Dropdown label="Week starts on" selectedKey={d.weekStart} options={DAY_NAMES.map((n, i) => ({ key: i, text: n }))} onChange={(_, o) => o && set({ weekStart: Number(o.key) })} />
            {d.weekStart !== savedWeekStart && (
              <MessageBar messageBarType={MessageBarType.warning} isMultiline styles={{ root: { marginTop: 8, borderRadius: 8 } }}>
                Weeks already filled in keep their old dates: they stay in reports and in My timesheets, but My timesheet opens the new weeks. Change the start day before people begin, or at a
                quiet moment (every timesheet submitted).
              </MessageBar>
            )}
          </Gate>
          <TextField label="Standard working day (hours)" type="number" value={String(d.standardDay)} onChange={(_, v) => set({ standardDay: Number(v) || 8 })} description="Used for leave, targets and utilisation." />
          <Dropdown
            label="Hours are entered in steps of"
            selectedKey={d.increment}
            options={[
              { key: 0.01, text: 'Any (0.01 h)' },
              { key: 0.1, text: '6 minutes (0.1 h)' },
              { key: 0.25, text: '15 minutes (0.25 h)' },
              { key: 0.5, text: '30 minutes (0.5 h)' },
              { key: 1, text: '1 hour' }
            ]}
            onChange={(_, o) => o && set({ increment: Number(o.key) })}
          />
        </div>
      </Card>
      <Card title="Week days for time tracking" icon="CalendarWorkWeek">
        <Gate feature="workDays" text="Choose which days of the week time is tracked on (Pro).">
          <div className="ts-row">
            {[0, 1, 2, 3, 4, 5, 6].map((i) => (
              <Checkbox key={i} label={DAY_SHORT[i]} checked={d.workDays.indexOf(i) >= 0} onChange={(_, v) => set({ workDays: (v ? d.workDays.concat(i) : d.workDays.filter((x) => x !== i)).sort() })} />
            ))}
          </div>
          <div className="ts-muted ts-small" style={{ marginTop: 6 }}>
            Other days are shaded on the timesheet; whether time may be logged on them is set per hour rule.
          </div>
        </Gate>
      </Card>
      <Card title="Days without time tracking (holidays)" icon="Vacation">
        <Gate feature="trackingDays" text="Select the days time is tracked: public holidays and per-person working days (Enterprise).">
          <div className="ts-row" style={{ alignItems: 'flex-end' }}>
            <DateField label="Date" value={hDate} onChange={setHDate} />
            <TextField label="Name" value={hName} onChange={(_, v) => setHName(v || '')} placeholder="e.g. Diwali" />
            <DefaultButton
              text="Add"
              disabled={!hDate}
              onClick={() => {
                set({ holidays: d.holidays.filter((h) => h.date !== hDate).concat({ date: hDate, name: hName || 'Holiday' }).sort((a, b) => a.date.localeCompare(b.date)) });
                setHDate('');
                setHName('');
              }}
            />
          </div>
          <div className="ts-chips" style={{ marginTop: 10 }}>
            {d.holidays.map((h) => (
              <span key={h.date} className="ts-chip">
                {formatDay(h.date, true)} · {h.name}
                <button aria-label="Remove" onClick={() => set({ holidays: d.holidays.filter((x) => x.date !== h.date) })}>
                  <Icon iconName="Cancel" />
                </button>
              </span>
            ))}
            {!d.holidays.length && <span className="ts-muted ts-small">No holidays yet.</span>}
          </div>
          <div className="ts-muted ts-small" style={{ marginTop: 6 }}>
            Per-person working days are set in People &amp; rates.
          </div>
        </Gate>
      </Card>
      <Card title="Timesheet lines" icon="BulletedList">
        <div className="ts-stack" style={{ gap: 8 }}>
          {has('billable') && <Toggle label="People may change billable / non-billable on a line" checked={d.allowBillableOverride} onChange={(_, v) => set({ allowBillableOverride: !!v })} />}
          <Toggle label="A note is required on every day with hours" checked={d.requireNotes} onChange={(_, v) => set({ requireNotes: !!v })} />
          {has('activityTracking') && <Toggle label="An activity is required on every line" checked={d.requireActivity} onChange={(_, v) => set({ requireActivity: !!v })} />}
          {has('selfAssign') && <Toggle label="People may assign open tasks to themselves" checked={d.selfAssign} onChange={(_, v) => set({ selfAssign: !!v })} />}
          <Toggle label="Members may create tasks on projects they belong to" checked={d.membersCreateTasks} onChange={(_, v) => set({ membersCreateTasks: !!v })} />
        </div>
      </Card>
      <SaveBar draft={draft} what="Tracking" />
    </div>
  );
}

function LimitFields(props: { value: { dailyMin: number; dailyMax: number; weeklyMin: number; weeklyMax: number }; onChange: (v: Partial<IHourRule>) => void }): JSX.Element {
  const v = props.value;
  const num = (label: string, k: 'dailyMin' | 'dailyMax' | 'weeklyMin' | 'weeklyMax'): JSX.Element => (
    <TextField label={label} type="number" value={v[k] ? String(v[k]) : ''} placeholder="No limit" onChange={(_, x) => props.onChange({ [k]: Number(x) || 0 })} />
  );
  return (
    <div className="ts-grid ts-grid-4">
      {num('Daily minimum', 'dailyMin')}
      {num('Daily maximum', 'dailyMax')}
      {num('Weekly minimum', 'weeklyMin')}
      {num('Weekly maximum', 'weeklyMax')}
    </div>
  );
}

function HoursAdmin(): JSX.Element {
  const draft = useDraft();
  const { d, set } = draft;
  const { ref } = useApp();
  const departments = ref.users.map((u) => u.department).filter((x, i, a) => !!x && a.indexOf(x) === i);
  const setRule = (i: number, patch: Partial<IHourRule>): void => {
    const rules = d.rules.slice();
    rules[i] = { ...rules[i], ...patch };
    set({ rules });
  };
  return (
    <div className="ts-stack">
      <Card title="Daily and weekly hour restrictions" icon="Clock">
        <Gate feature="hourLimits" text="Restrict daily and weekly hours (Pro).">
          <LimitFields value={d.limits} onChange={(p) => set({ limits: { ...d.limits, ...p } })} />
          <ChoiceGroup
            label="When a week breaks a limit"
            selectedKey={d.limits.enforcement}
            options={[
              { key: 'warn', text: 'Warn, but allow submitting' },
              { key: 'block', text: 'Block submitting until it is fixed' }
            ]}
            onChange={(_, o) => o && set({ limits: { ...d.limits, enforcement: o.key as 'warn' | 'block' } })}
          />
          <div className="ts-muted ts-small">Minimums are checked on submit and count approved time off.</div>
        </Gate>
      </Card>
      <Card title="Hours based on rules" icon="BranchFork2">
        <Gate feature="hourRules" text="Different daily and weekly limits by role, department, person or external users (Enterprise).">
          <Toggle label="Use rules (the first matching rule replaces the limits above)" checked={d.rulesEnabled} onChange={(_, v) => set({ rulesEnabled: !!v })} />
          {d.rules.map((r, i) => (
            <div key={r.id} className="ts-card" style={{ padding: 14, marginTop: 10, background: 'var(--ts-surface-2)' }}>
              <div className="ts-row">
                <TextField value={r.name} onChange={(_, v) => setRule(i, { name: v || '' })} ariaLabel="Rule name" styles={{ root: { flex: 1 } }} />
                <DefaultButton iconProps={{ iconName: 'Up' }} ariaLabel="Earlier" disabled={!i} onClick={() => set({ rules: d.rules.map((x, j) => (j === i - 1 ? d.rules[i] : j === i ? d.rules[i - 1] : x)) })} styles={{ root: { minWidth: 32, padding: 0 } }} />
                <DefaultButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove rule" onClick={() => set({ rules: d.rules.filter((_, j) => j !== i) })} styles={{ root: { minWidth: 32, padding: 0 } }} />
              </div>
              <div className="ts-form-grid" style={{ marginTop: 8 }}>
                <Dropdown label="Roles (empty = any)" multiSelect selectedKeys={r.roles} options={ROLES.map((x) => ({ key: x, text: ROLE_LABELS[x] }))} onChange={(_, o) => o && setRule(i, { roles: o.selected ? r.roles.concat(o.key as Role) : r.roles.filter((x) => x !== o.key) })} />
                <Dropdown label="Departments (empty = any)" multiSelect selectedKeys={r.departments} options={departments.map((x) => ({ key: x, text: x }))} onChange={(_, o) => o && setRule(i, { departments: o.selected ? r.departments.concat(String(o.key)) : r.departments.filter((x) => x !== o.key) })} />
                <PeopleField label="People (empty = any)" value={r.emails} onChange={(v) => setRule(i, { emails: v })} onlyKnown />
                <div className="ts-stack" style={{ gap: 6, paddingTop: 28 }}>
                  <Checkbox label="Only external users" checked={r.externalOnly} onChange={(_, v) => setRule(i, { externalOnly: !!v })} />
                  <Checkbox label="May log time on non-working days" checked={r.allowNonWorkingDays} onChange={(_, v) => setRule(i, { allowNonWorkingDays: !!v })} />
                  <Checkbox label="Block submitting when broken" checked={r.enforcement === 'block'} onChange={(_, v) => setRule(i, { enforcement: v ? 'block' : 'warn' })} />
                </div>
              </div>
              <LimitFields value={r} onChange={(p) => setRule(i, p)} />
            </div>
          ))}
          <DefaultButton
            iconProps={{ iconName: 'Add' }}
            text="Add a rule"
            styles={{ root: { marginTop: 10 } }}
            onClick={() =>
              set({
                rules: d.rules.concat({ id: Math.random().toString(36).substring(2, 9), name: `Rule ${d.rules.length + 1}`, roles: [], departments: [], emails: [], externalOnly: false, allowNonWorkingDays: false, enforcement: 'warn', dailyMin: 0, dailyMax: 10, weeklyMin: 0, weeklyMax: 45 }),
                rulesEnabled: true
              })
            }
          />
        </Gate>
      </Card>
      <SaveBar draft={draft} what="Hour restrictions" />
    </div>
  );
}

function ApprovalsAdmin(): JSX.Element {
  const draft = useDraft();
  const { d, set } = draft;
  const a = d.autoApproval;
  return (
    <div className="ts-stack">
      <Card title="Who approves" icon="DocumentApproval">
        <ChoiceGroup
          selectedKey={d.approvalBy}
          options={[
            { key: 'manager', text: "The person's manager (People & rates › Manager)" },
            { key: 'project', text: 'The managers of the projects on the week (and programme managers)' },
            { key: 'either', text: 'Either — the manager or any project manager on the week' }
          ]}
          onChange={(_, o) => o && set({ approvalBy: o.key as ITimesheetConfig['approvalBy'] })}
        />
        <div className="ts-muted ts-small">Administrators can approve any timesheet. Nobody approves their own week. Without an approver, administrators are asked.</div>
      </Card>
      <Card title="Role-based auto approval" icon="Robot">
        <Gate feature="autoApproval" text="Approve the timesheets of chosen roles automatically when they are submitted (Pro).">
          <Toggle label="Approve automatically on submit" checked={a.enabled} onChange={(_, v) => set({ autoApproval: { ...a, enabled: !!v } })} />
          <div className="ms-Label">For these roles</div>
          <div className="ts-row">
            {ROLES.map((r) => (
              <Checkbox key={r} label={ROLE_LABELS[r]} checked={a.roles.indexOf(r) >= 0} onChange={(_, v) => set({ autoApproval: { ...a, roles: v ? a.roles.concat(r) : a.roles.filter((x) => x !== r) } })} />
            ))}
          </div>
          <Toggle label="Only when the week passes every hour rule without warnings" checked={a.requireWithinRules} onChange={(_, v) => set({ autoApproval: { ...a, requireWithinRules: !!v } })} styles={{ root: { marginTop: 12 } }} />
          <TextField label="Only up to this many hours a week (empty = any)" type="number" value={a.maxHours ? String(a.maxHours) : ''} onChange={(_, v) => set({ autoApproval: { ...a, maxHours: Number(v) || 0 } })} styles={{ root: { maxWidth: 280 } }} />
        </Gate>
      </Card>
      <SaveBar draft={draft} what="Approvals" />
    </div>
  );
}

function CostsAdmin(): JSX.Element {
  const draft = useDraft();
  const { d, set } = draft;
  const { has } = useApp();
  return (
    <Card title="Costs & rates" icon="Money">
      <Gate feature="projectCost">
        <ChoiceGroup
          label="Hourly cost of an hour of work"
          selectedKey={d.costMode}
          options={[
            { key: 'userFirst', text: "The person's hourly cost; else the project's flat hourly cost", disabled: !has('userCost') },
            { key: 'projectFirst', text: "The project's flat hourly cost; else the person's" },
            { key: 'userOnly', text: "Always the person's hourly cost", disabled: !has('userCost') },
            { key: 'projectOnly', text: "Always the project's flat hourly cost" }
          ]}
          onChange={(_, o) => o && set({ costMode: o.key as CostMode })}
        />
        <div className="ts-form-grid" style={{ marginTop: 10 }}>
          <TextField label={`Default hourly cost (${d.currency})`} type="number" value={d.defaultCostRate ? String(d.defaultCostRate) : ''} onChange={(_, v) => set({ defaultCostRate: Number(v) || 0 })} description="When neither the person nor the project has one." />
          <TextField label={`Default bill rate (${d.currency})`} type="number" value={d.defaultBillRate ? String(d.defaultBillRate) : ''} onChange={(_, v) => set({ defaultBillRate: Number(v) || 0 })} />
        </div>
        <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginTop: 12, borderRadius: 8 } }} isMultiline>
          Bill rates: a person-specific rate on the project{has('userRates') ? '' : ' (Pro)'}, else the person&apos;s bill rate{has('userRates') ? '' : ' (Pro)'}, else the project&apos;s rate, else the default. Rates are captured on each line when the week is submitted.
        </MessageBar>
      </Gate>
      <SaveBar draft={draft} what="Costs & rates" />
    </Card>
  );
}

function ColumnsAdmin(): JSX.Element {
  const draft = useDraft();
  const { d, set } = draft;
  const setCol = (i: number, patch: Partial<ICustomColumn>): void => {
    const cols = d.customColumns.slice();
    cols[i] = { ...cols[i], ...patch };
    set({ customColumns: cols });
  };
  return (
    <Card title="Custom columns" icon="TripleColumn">
      <Gate feature="customColumns">
        <p className="ts-muted" style={{ marginTop: 0 }}>
          Extra fields on each timesheet line (e.g. ticket number, location, work order) or on each project (e.g. PO number, region). They show in exports and can group reports.
        </p>
        {d.customColumns.map((c, i) => (
          <div key={c.key} className="ts-row" style={{ alignItems: 'flex-end', marginBottom: 10 }}>
            <TextField label="Label" value={c.label} onChange={(_, v) => setCol(i, { label: v || '' })} />
            <Dropdown
              label="Type"
              selectedKey={c.type}
              options={['text', 'number', 'choice', 'boolean', 'date'].map((k) => ({ key: k, text: k === 'boolean' ? 'Yes / no' : k[0].toUpperCase() + k.substring(1) }))}
              onChange={(_, o) => o && setCol(i, { type: o.key as ICustomColumn['type'] })}
              styles={{ root: { width: 120 } }}
            />
            <Dropdown
              label="On"
              selectedKey={c.level}
              options={[
                { key: 'entry', text: 'Timesheet lines' },
                { key: 'project', text: 'Projects' }
              ]}
              onChange={(_, o) => o && setCol(i, { level: o.key as ICustomColumn['level'] })}
              styles={{ root: { width: 150 } }}
            />
            {c.type === 'choice' && <TextField label="Choices (comma separated)" value={c.choices.join(', ')} onChange={(_, v) => setCol(i, { choices: (v || '').split(',').map((x) => x.trim()).filter((x) => !!x) })} />}
            <Checkbox label="Required" checked={c.required} onChange={(_, v) => setCol(i, { required: !!v })} />
            <Checkbox label="In reports" checked={c.inReports} onChange={(_, v) => setCol(i, { inReports: !!v })} />
            <DefaultButton iconProps={{ iconName: 'Delete' }} ariaLabel="Remove" onClick={() => set({ customColumns: d.customColumns.filter((_, j) => j !== i) })} styles={{ root: { minWidth: 32, padding: 0 } }} />
          </div>
        ))}
        <DefaultButton
          iconProps={{ iconName: 'Add' }}
          text="Add a column"
          onClick={() => set({ customColumns: d.customColumns.concat({ key: 'c' + Math.random().toString(36).substring(2, 8), label: 'New column', type: 'text', level: 'entry', choices: [], required: false, inReports: true }) })}
        />
      </Gate>
      <SaveBar draft={draft} what="Custom columns" />
    </Card>
  );
}
