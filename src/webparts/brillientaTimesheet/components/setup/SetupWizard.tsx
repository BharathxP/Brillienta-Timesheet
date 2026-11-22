import * as React from 'react';
import { Checkbox, DefaultButton, Dropdown, MessageBar, MessageBarType, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { ICurrentUser } from '../../models/types';
import { SPClient } from '../../services/SPClient';
import { TimesheetService } from '../../services/TimesheetService';
import { SETTING_KEYS, SettingsStore } from '../../services/SettingsStore';
import { LogLevel, Provisioner } from '../../services/Provisioner';
import { SCHEMA_VERSION } from '../../schema/schema';
import { DEFAULT_ACTIVITIES, PALETTE, defaultConfig } from '../../schema/defaults';
import { CURRENCIES } from '../../services/format';
import { DAY_NAMES, DAY_SHORT, addDays, today, weekStartOf } from '../../services/dates';
import { DEFAULT_LOGO } from '../../assets/defaultLogo';

interface IProps {
  sp: SPClient;
  svc: TimesheetService;
  settings: SettingsStore;
  prefix: string;
  me: ICurrentUser;
  appVersion: string;
  tenantId?: string;
  siteTitle: string;
  onDone: () => void;
}

const STEPS = ['Welcome', 'Your organisation', 'Time tracking', 'Create'];

/** First run: a site owner sets the basics, then the lists are created with sample activities (and an optional demo project). */
export function SetupWizard(props: IProps): JSX.Element {
  const [step, setStep] = React.useState(0);
  const [cfg, setCfg] = React.useState(() => ({ ...defaultConfig(props.siteTitle), orgName: props.siteTitle }));
  const [sample, setSample] = React.useState(true);
  const [log, setLog] = React.useState<Array<{ m: string; l: LogLevel }>>([]);
  const [running, setRunning] = React.useState(false);
  const [error, setError] = React.useState('');
  const add = (m: string, l: LogLevel): void => setLog((x) => x.concat({ m, l }));

  const install = async (): Promise<void> => {
    setRunning(true);
    setError('');
    try {
      const report = await new Provisioner(props.sp, props.prefix).ensureSchema(add);
      if (!report.ok) throw new Error('Some lists could not be created: ' + report.lists.filter((l) => l.errors.length).map((l) => `${l.title}: ${l.errors[0]}`).join('; '));
      await props.settings.load();
      const s = props.settings;
      await s.set(SETTING_KEYS.installedBy, props.me.email);
      await s.set(SETTING_KEYS.schemaVersion, String(SCHEMA_VERSION));
      await s.set(SETTING_KEYS.appVersion, props.appVersion);
      await s.setJson(SETTING_KEYS.config, cfg);
      add('Settings saved', 'success');
      const ref = await props.svc.loadReference();
      if (!ref.activities.length) {
        await props.svc.addActivities(DEFAULT_ACTIVITIES.map((a) => ({ title: a.title, billable: a.billable, color: a.color, isTimeOff: !!a.timeOff, active: true })));
        add(`${DEFAULT_ACTIVITIES.length} activities added`, 'success');
      }
      if (!ref.users.some((u) => u.email.toLowerCase() === props.me.email.toLowerCase())) {
        await props.svc.saveUser({ name: props.me.title, email: props.me.email, role: 'admin', active: true });
        add(`${props.me.title} added as administrator`, 'success');
      }
      if (sample && !ref.projects.length) {
        const ids = await props.svc.addProjects([
          { title: 'Website redesign', code: 'WEB-01', client: 'Contoso', managerEmails: [props.me.email], memberEmails: [], billable: true, status: 'Active', startDate: today(), endDate: addDays(today(), 90), plannedHours: 400, hourlyCost: 1200, billRate: 2500, color: PALETTE[0], memberRates: {}, custom: {}, endNotices: [], tmProjectId: 0, description: 'Sample project — rename or delete it.', programId: 0 },
          { title: 'Internal', code: 'INT', client: '', managerEmails: [props.me.email], memberEmails: [], billable: false, status: 'Active', startDate: '', endDate: '', plannedHours: 0, hourlyCost: 0, billRate: 0, color: PALETTE[8], memberRates: {}, custom: {}, endNotices: [], tmProjectId: 0, description: 'Internal, non-billable work.', programId: 0 }
        ]);
        if (ids[0]) {
          await props.svc.addTasks(
            ['Discovery workshops', 'Design', 'Build', 'Testing & launch'].map((t, i) => ({ title: t, projectId: ids[0], assigneeEmails: [], billable: true, plannedHours: 100, status: i ? 'Open' : 'In progress', dueDate: addDays(weekStartOf(today(), cfg.weekStart), 21 * (i + 1)), selfAssign: true, description: '', plannerTaskId: '', plannerPlanId: '', tmTaskId: 0 }))
          );
        }
        add('Sample projects and tasks added', 'success');
      }
      // Saved last: until it exists the app treats the site as not set up, so an interrupted setup resumes here.
      if (!s.get(SETTING_KEYS.installedAt)) await s.set(SETTING_KEYS.installedAt, new Date().toISOString());
      await props.svc.log('Brillienta Timesheet installed', `v${props.appVersion}`);
      add('Done. Open Admin › Schema & security to lock the lists down when you are ready.', 'success');
      props.onDone();
    } catch (e) {
      setError((e as Error).message);
      add((e as Error).message, 'error');
    } finally {
      setRunning(false);
    }
  };

  return (
    <div className="ts-setup">
      <div className="ts-row">
        <img className="ts-logo is-lg" src={DEFAULT_LOGO} alt="" />
        <div>
          <h1 style={{ margin: 0 }}>Set up Brillienta Timesheet</h1>
          <div className="ts-muted">On {props.siteTitle}. Everything stays in SharePoint lists on this site.</div>
        </div>
      </div>
      <div className="ts-steps">
        {STEPS.map((s, i) => (
          <div key={s} className={'ts-step' + (i === step ? ' is-active' : i < step ? ' is-done' : '')}>
            {i + 1}. {s}
          </div>
        ))}
      </div>
      {step === 0 && (
        <div className="ts-card ts-stack">
          <p style={{ margin: 0 }}>
            Weekly timesheets with approvals, projects, tasks and activities, billable hours, hourly cost and bill rates, hour rules, reminders, project and team reports, dashboards, and Microsoft 365
            integrations (Teams, Outlook, Planner, Time off Manager 365, Power Automate, Power BI).
          </p>
          <p className="ts-muted" style={{ margin: 0 }}>
            The setup creates {13} hidden lists prefixed <code>{props.prefix}</code>. The 14-day trial includes every Enterprise feature.
          </p>
        </div>
      )}
      {step === 1 && (
        <div className="ts-card ts-form-grid">
          <TextField label="Organisation name" value={cfg.orgName} onChange={(_, v) => setCfg({ ...cfg, orgName: v || '' })} />
          <Dropdown label="Currency" selectedKey={cfg.currency} options={CURRENCIES.map((c) => ({ key: c, text: c }))} onChange={(_, o) => o && setCfg({ ...cfg, currency: String(o.key) })} />
          <Dropdown
            label="Who approves timesheets"
            selectedKey={cfg.approvalBy}
            options={[
              { key: 'manager', text: "The person's manager" },
              { key: 'project', text: 'The project managers' },
              { key: 'either', text: 'Either of them' }
            ]}
            onChange={(_, o) => o && setCfg({ ...cfg, approvalBy: o.key as typeof cfg.approvalBy })}
          />
          <Toggle label="Add people automatically on their first visit" checked={cfg.autoEnrol} onChange={(_, v) => setCfg({ ...cfg, autoEnrol: !!v })} />
        </div>
      )}
      {step === 2 && (
        <div className="ts-card ts-stack">
          <div className="ts-form-grid">
            <Dropdown label="Week starts on" selectedKey={cfg.weekStart} options={DAY_NAMES.map((d, i) => ({ key: i, text: d }))} onChange={(_, o) => o && setCfg({ ...cfg, weekStart: Number(o.key) })} />
            <TextField label="Standard working day (hours)" type="number" value={String(cfg.standardDay)} onChange={(_, v) => setCfg({ ...cfg, standardDay: Number(v) || 8 })} />
          </div>
          <div>
            <div className="ms-Label">Working days</div>
            <div className="ts-row">
              {[0, 1, 2, 3, 4, 5, 6].map((d) => (
                <Checkbox key={d} label={DAY_SHORT[d]} checked={cfg.workDays.indexOf(d) >= 0} onChange={(_, v) => setCfg({ ...cfg, workDays: (v ? cfg.workDays.concat(d) : cfg.workDays.filter((x) => x !== d)).sort() })} />
              ))}
            </div>
          </div>
          <Toggle label="Add sample projects and tasks" checked={sample} onChange={(_, v) => setSample(!!v)} />
        </div>
      )}
      {step === 3 && (
        <div className="ts-stack">
          {error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>}
          <div className="ts-log" style={{ minHeight: 120 }}>
            {!log.length && <div>Ready to create the lists.</div>}
            {log.map((l, i) => (
              <div key={i} className={'is-' + l.l}>
                {l.m}
              </div>
            ))}
          </div>
        </div>
      )}
      <div className="ts-row" style={{ marginTop: 18 }}>
        {step > 0 && <DefaultButton text="Back" onClick={() => setStep(step - 1)} disabled={running} />}
        {step < 3 && <PrimaryButton text="Next" onClick={() => setStep(step + 1)} />}
        {step === 3 && <PrimaryButton text={running ? 'Creating…' : 'Create the timesheet'} onClick={() => void install()} disabled={running} />}
      </div>
    </div>
  );
}
