import * as React from 'react';
import { Customizer, ILayerStyles, MessageBar, MessageBarType, PrimaryButton, ThemeProvider } from '@fluentui/react';
import { MSGraphClientFactory, SPHttpClient } from '@microsoft/sp-http';
import styles from './Timesheet.module.scss';
import { AppCtx, IAppContext, IHostInfo, IRoute, SHAREPOINT_HOST, ToastKind, errorText, parseRoute } from './AppContext';
import { ICurrentUser, ITimesheetConfig, ITsUser } from '../models/types';
import { SPClient } from '../services/SPClient';
import { IReference, TimesheetService } from '../services/TimesheetService';
import { ISecurityMeta, SETTING_KEYS, SettingsStore } from '../services/SettingsStore';
import { GraphService } from '../services/GraphService';
import { Notifier } from '../services/Notifier';
import { Workflow } from '../services/workflow';
import { ReminderRunner } from '../services/ReminderRunner';
import { TimeOffService } from '../services/TimeOffService';
import { TaskManagerService } from '../services/TaskManagerService';
import { LogLevel, Provisioner } from '../services/Provisioner';
import { SCHEMA_VERSION } from '../schema/schema';
import { mergeConfig } from '../schema/defaults';
import { ILicenseStatus } from '../services/licensing/license';
import { loadLicenseStatus } from '../services/licensing/LicenseLoader';
import { Feature, applyPlan, planHas, restoreLocked, userLimit } from '../services/licensing/plans';
import { canApprove, effectiveRole, reportsOf, sameEmail } from '../services/access';
import { isShareablePage } from '../services/host';
import { DEFAULT_LOGO } from '../assets/defaultLogo';
import { Shell } from './Shell';
import { SetupWizard } from './setup/SetupWizard';
import { LicenseLocked } from './admin/LicenseAdmin';
import { Loading } from './common/Ui';
import { NAVY, brandTheme, ensureFonts } from './theme';

export interface IAppProps {
  http: SPHttpClient;
  webUrl: string;
  siteTitle: string;
  user: { displayName: string; email: string; loginName: string; isExternal?: boolean };
  permissions: { manageLists: boolean; manageWeb: boolean };
  tenantId?: string;
  prefix: string;
  appVersion: string;
  isEditMode: boolean;
  graphFactory?: MSGraphClientFactory;
  host?: IHostInfo;
}

type Phase = { kind: 'loading' } | { kind: 'error'; message: string } | { kind: 'setup' } | { kind: 'notInstalled' } | { kind: 'upgrading'; log: Array<{ m: string; l: LogLevel }> } | { kind: 'ready' };

interface IToast {
  id: number;
  message: string;
  kind: ToastKind;
}

const EMPTY_REF: IReference = { users: [], programs: [], projects: [], tasks: [], activities: [] };

const EDIT_BARS = ['#SuiteNavWrapper', '#SuiteNavPlaceHolder', '[data-automation-id="pageCommandBarRegion"]', '#spCommandBar', '.commandBarWrapper', '[class*="CanvasToolbar"]', '[class*="workbenchCommandBar"]'];

/** While the page is edited the app starts below SharePoint's edit bars, so Save / Publish stay reachable. */
function useEditBarOffset(editing: boolean): number {
  const [offset, setOffset] = React.useState(0);
  React.useEffect(() => {
    if (!editing) {
      setOffset(0);
      return;
    }
    const measure = (): void => {
      let bottom = 0;
      EDIT_BARS.forEach((sel) =>
        document.querySelectorAll(sel).forEach((el) => {
          if ((el as HTMLElement).closest('[data-ts-top]')) return;
          const r = el.getBoundingClientRect();
          if (r.height > 0 && r.height < 140 && r.top < 160 && r.bottom > bottom) bottom = r.bottom;
        })
      );
      setOffset(Math.round(Math.min(200, bottom || 96)));
    };
    measure();
    const t = window.setInterval(measure, 1500);
    return () => window.clearInterval(t);
  }, [editing]);
  return offset;
}

export default function App(props: IAppProps): JSX.Element {
  const host = props.host || SHAREPOINT_HOST;
  const sp = React.useMemo(() => new SPClient(props.http, props.webUrl), [props.http, props.webUrl]);
  const svc = React.useMemo(() => new TimesheetService(sp, props.prefix), [sp, props.prefix]);
  const settings = React.useMemo(() => new SettingsStore(sp, props.prefix), [sp, props.prefix]);
  const graph = React.useMemo(() => (props.graphFactory ? new GraphService(props.graphFactory) : undefined), [props.graphFactory]);
  const timeOff = React.useMemo(() => new TimeOffService(sp), [sp]);
  const configRef = React.useRef<ITimesheetConfig>(mergeConfig(undefined));
  const taskManager = React.useMemo(() => new TaskManagerService(sp, () => configRef.current.taskManager), [sp]);
  const refRef = React.useRef<IReference>(EMPTY_REF);
  const adminsRef = React.useRef<string[]>([]);
  const licenseRef = React.useRef<ILicenseStatus | undefined>();
  const linkBase = React.useMemo(() => {
    const here = window.location.href.split('#')[0];
    const stored = settings.get(SETTING_KEYS.pageUrl) || '';
    return isShareablePage(here) ? here : stored || here;
  }, [settings, settings.get(SETTING_KEYS.pageUrl)]);
  const notifier = React.useMemo(() => new Notifier(svc, () => configRef.current, window.location.href.split('#')[0], graph), [svc, graph]);
  notifier.linkBase = linkBase;
  const hasFeature = React.useCallback((f: Feature) => planHas(licenseRef.current && licenseRef.current.plan, f, licenseRef.current && licenseRef.current.addons), []);
  const workflow = React.useMemo(
    () =>
      new Workflow({
        svc,
        notifier,
        config: () => configRef.current,
        dir: () => refRef.current,
        admins: () => adminsRef.current,
        has: hasFeature,
        taskManager
      }),
    [svc, notifier, hasFeature, taskManager]
  );
  const reminders = React.useMemo(() => new ReminderRunner(svc, settings, notifier), [svc, settings, notifier]);

  const [phase, setPhase] = React.useState<Phase>({ kind: 'loading' });
  const [me, setMe] = React.useState<ICurrentUser | undefined>();
  const [saved, setSaved] = React.useState<ITimesheetConfig>(() => mergeConfig(undefined));
  const [license, setLicense] = React.useState<ILicenseStatus | undefined>();
  const [ref, setRef] = React.useState<IReference>(EMPTY_REF);
  const [security, setSecurity] = React.useState<ISecurityMeta | undefined>();
  const [logoStored, setLogoStored] = React.useState('');
  const [route, setRoute] = React.useState<IRoute>(() => parseRoute(window.location.hash));
  const [toasts, setToasts] = React.useState<IToast[]>([]);
  const [boot, setBoot] = React.useState(0);
  const [approvalsCount, setApprovalsCount] = React.useState(0);
  const [approvalsTick, setApprovalsTick] = React.useState(0);
  const [notificationsVersion, setNotificationsVersion] = React.useState(0);

  licenseRef.current = license;
  const plan = license ? license.plan : undefined;
  const config = React.useMemo(() => applyPlan(saved, plan, license && license.addons), [saved, plan, license]);
  configRef.current = config;
  refRef.current = ref;
  notifier.teamsOutlook = planHas(plan, 'teamsOutlook', license && license.addons);

  const toast = React.useCallback((message: string, kind: ToastKind = 'info') => {
    const id = Date.now() + Math.random();
    setToasts((t) => t.concat({ id, message, kind }).slice(-4));
    window.setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), kind === 'error' ? 8000 : 4500);
  }, []);
  notifier.onProblem = (m: string) => toast(m, 'warning');

  React.useEffect(() => {
    const onHash = (): void => setRoute(parseRoute(window.location.hash));
    window.addEventListener('hashchange', onHash);
    return () => window.removeEventListener('hashchange', onHash);
  }, []);

  const navigate = React.useCallback((path: string) => {
    window.location.hash = '#/' + path.replace(/^#?\/?/, '');
    const main = document.querySelector('[data-ts-main]');
    if (main) main.scrollTop = 0;
  }, []);

  const reloadRef = React.useCallback(async () => {
    setRef(await svc.loadReference());
  }, [svc]);

  const refreshLicense = React.useCallback(async () => {
    await settings.load();
    setLicense(await loadLicenseStatus(sp, settings, props.prefix, props.tenantId));
  }, [sp, settings, props.prefix, props.tenantId]);

  // ------------------------------------------------------------------ boot
  React.useEffect(() => {
    let alive = true;
    (async () => {
      setPhase({ kind: 'loading' });
      const u = await sp.get<{ Id: number; Title: string; Email: string; LoginName: string; IsSiteAdmin: boolean }>('web/currentuser?$select=Id,Title,Email,LoginName,IsSiteAdmin');
      const current: ICurrentUser = {
        id: u.Id,
        title: u.Title || props.user.displayName,
        email: u.Email || props.user.email,
        loginName: u.LoginName,
        isSiteAdmin: u.IsSiteAdmin,
        canManageLists: props.permissions.manageLists || u.IsSiteAdmin,
        canManageWeb: props.permissions.manageWeb || u.IsSiteAdmin,
        isExternal: !!props.user.isExternal || /#ext#/i.test(u.LoginName || '')
      };
      svc.actor = current.title;
      if (!alive) return;
      setMe(current);
      // Installed = the setup wizard finished (it saves the install date last). Lists without it come from an
      // interrupted first setup: the wizard runs again and completes it (it skips what already exists).
      const installed = (await settings.load()) && !!settings.get(SETTING_KEYS.installedAt);
      if (!installed) {
        setPhase({ kind: current.canManageLists ? 'setup' : 'notInstalled' });
        return;
      }
      // A new version: site owners bring the lists up to date (adds columns; never deletes data).
      const schemaOld = Number(settings.get(SETTING_KEYS.schemaVersion) || 0) < SCHEMA_VERSION;
      const appOld = (settings.get(SETTING_KEYS.appVersion) || '') !== props.appVersion;
      if ((schemaOld || appOld) && current.canManageLists) {
        const log: Array<{ m: string; l: LogLevel }> = [];
        setPhase({ kind: 'upgrading', log });
        await new Provisioner(sp, props.prefix).ensureSchema((m, l) => {
          log.push({ m, l });
          if (alive) setPhase({ kind: 'upgrading', log: log.slice() });
        });
        await settings.set(SETTING_KEYS.schemaVersion, String(SCHEMA_VERSION));
        await settings.set(SETTING_KEYS.appVersion, props.appVersion);
      }
      const cfg = mergeConfig(settings.getJson(SETTING_KEYS.config));
      const [lic, reference] = await Promise.all([loadLicenseStatus(sp, settings, props.prefix, props.tenantId), svc.loadReference()]);
      if (!alive) return;
      setSaved(cfg);
      setLicense(lic);
      setRef(reference);
      setSecurity(settings.getJson<ISecurityMeta>(SETTING_KEYS.security));
      setLogoStored(settings.get(SETTING_KEYS.brandLogo) || '');
      // Remember the SharePoint page the app lives on, so links sent from Teams open it.
      const here = window.location.href.split('#')[0];
      if (current.canManageWeb && isShareablePage(here) && settings.get(SETTING_KEYS.pageUrl) !== here) settings.set(SETTING_KEYS.pageUrl, here).catch(() => undefined);
      setPhase({ kind: 'ready' });
    })().catch((e) => alive && setPhase({ kind: 'error', message: errorText(e) }));
    return () => {
      alive = false;
    };
  }, [sp, svc, settings, boot, props.appVersion, props.prefix, props.tenantId]);

  // ------------------------------------------------------------------ who is the signed-in person
  const myUser: ITsUser | undefined = me ? ref.users.filter((u) => sameEmail(u.email, me.email))[0] : undefined;
  const configAdmin = !!me && config.admins.some((a) => sameEmail(a, me.email));
  const role = me ? effectiveRole(myUser, { siteAdmin: me.canManageWeb, configAdmin, programManagerAllowed: planHas(plan, 'programManager') }) : 'member';
  const roleReason = !me ? '' : me.canManageWeb ? 'Site owner' : configAdmin ? 'Listed as an administrator' : myUser ? 'Role set in People' : 'Not added yet (member)';
  const admins = React.useMemo(() => {
    const list = ref.users.filter((u) => u.active && u.role === 'admin').map((u) => u.email).concat(config.admins);
    return list.filter((e, i) => list.map((x) => x.toLowerCase()).indexOf(e.toLowerCase()) === i);
  }, [ref.users, config.admins]);
  adminsRef.current = admins;
  const isAdmin = role === 'admin';
  const isApprover = isAdmin || role !== 'member' || (!!me && reportsOf(ref.users, me.email, true).length > 0);

  // Users within the plan's limit: the earliest added active people.
  const limit = userLimit(plan, license && license.users);
  const activeUsers = ref.users.filter((u) => u.active).sort((a, b) => (a.created || '').localeCompare(b.created || '') || a.id - b.id);
  const usersOver = limit > 0 ? Math.max(0, activeUsers.length - limit) : 0;
  const myIndex = me ? activeUsers.map((u) => u.email.toLowerCase()).indexOf(me.email.toLowerCase()) : -1;
  const overLimit = limit > 0 && !isAdmin && (myIndex >= limit || (myIndex < 0 && activeUsers.length >= limit));

  // New people are added to People on their first visit (when allowed), so managers and reports see them.
  React.useEffect(() => {
    if (phase.kind !== 'ready' || !me || myUser || !config.autoEnrol) return;
    svc
      .saveUser({ name: me.title, email: me.email, role: config.defaultRole === 'admin' ? 'member' : config.defaultRole, external: me.isExternal, active: true })
      .then(() => reloadRef())
      .catch(() => undefined); // people without edit rights on People are added by an admin instead
  }, [phase.kind, me, myUser, config.autoEnrol]);

  // Timesheets waiting for this person's approval (navigation badge).
  React.useEffect(() => {
    if (phase.kind !== 'ready' || !me || !isApprover) return;
    let alive = true;
    (async () => {
      const sheets = (await svc.sheetsByStatus('Submitted')).filter((s) => !sameEmail(s.userEmail, me.email));
      if (!sheets.length) return 0;
      const entries = await svc.entriesOfSheets(sheets.map((s) => s.id));
      return sheets.filter((s) =>
        canApprove(
          role,
          me.email,
          s,
          entries.filter((e) => e.sheetId === s.id),
          ref,
          config
        )
      ).length;
    })()
      .then((n) => alive && setApprovalsCount(n))
      .catch(() => undefined);
    return () => {
      alive = false;
    };
  }, [phase.kind, me, isApprover, role, ref, approvalsTick]);

  // Reminders and project end-date notices run from an administrator's browser.
  React.useEffect(() => {
    if (phase.kind !== 'ready' || !me || !isAdmin || !license || !license.plan) return;
    reminders
      .run({ cfg: config, users: ref.users, projects: ref.projects, actor: me.title, remindersOn: planHas(plan, 'reminders'), endDatesOn: planHas(plan, 'projectEndDates') })
      .then((r) => {
        if (r && (r.reminders || r.notices)) toast(`Sent ${r.reminders} timesheet reminder${r.reminders === 1 ? '' : 's'} and ${r.notices} project notice${r.notices === 1 ? '' : 's'}.`, 'info');
      })
      .catch(() => undefined);
  }, [phase.kind, isAdmin, !!license]);

  // Brillienta Task Manager: projects and tasks come in automatically from an administrator's browser, at most hourly.
  React.useEffect(() => {
    if (phase.kind !== 'ready' || !isAdmin || !planHas(plan, 'taskManager') || !config.taskManager.enabled || !config.taskManager.autoSync) return;
    const last = Date.parse(config.taskManager.lastSyncAt || '');
    if (!isNaN(last) && Date.now() - last < 3600000) return;
    (async () => {
      const r = await taskManager.sync(svc, refRef.current);
      const next = { ...saved, taskManager: { ...saved.taskManager, lastSyncAt: new Date().toISOString() } };
      await settings.setJson(SETTING_KEYS.config, next);
      setSaved(next);
      if (r.projectsCreated + r.projectsUpdated + r.tasksCreated + r.tasksUpdated > 0) {
        await reloadRef();
        toast(`Task Manager: ${r.projectsCreated} projects and ${r.tasksCreated} tasks added, ${r.projectsUpdated + r.tasksUpdated} updated.`, 'info');
      }
    })().catch((e) => toast('Task Manager sync failed: ' + errorText(e), 'warning'));
  }, [phase.kind, isAdmin, plan, config.taskManager.enabled]);

  // ------------------------------------------------------------------ frame
  const brand = config.brandColor || NAVY;
  const fluentTheme = React.useMemo(() => brandTheme(brand), [brand]);
  const themeVars = React.useMemo(() => ({ '--ts-brand': brand, '--ts-navy': brand }) as unknown as React.CSSProperties, [brand]);
  const layerSettings = React.useMemo(() => ({ Layer: { styles: { root: [styles.root, themeVars] } as unknown as ILayerStyles } }), [themeVars]);
  const topOffset = useEditBarOffset(props.isEditMode);
  React.useEffect(() => {
    const html = document.documentElement;
    const previous = html.style.overflow;
    html.style.overflow = 'hidden';
    return () => {
      html.style.overflow = previous;
    };
  }, []);
  React.useEffect(() => ensureFonts(), []);
  const frameClass = 'ts-app is-fullscreen';
  const frameStyle: React.CSSProperties | undefined = topOffset ? { top: topOffset, height: `calc(100vh - ${topOffset}px)` } : undefined;
  const logo = /^data:image\/(png|jpeg|gif|webp|svg\+xml)/.test(logoStored) ? logoStored : DEFAULT_LOGO;

  const toastArea = (
    <div className="ts-toasts" role="status" aria-live="polite">
      {toasts.map((t) => (
        <MessageBar
          key={t.id}
          messageBarType={t.kind === 'success' ? MessageBarType.success : t.kind === 'error' ? MessageBarType.error : t.kind === 'warning' ? MessageBarType.warning : MessageBarType.info}
          onDismiss={() => setToasts((x) => x.filter((y) => y.id !== t.id))}
          isMultiline
          styles={{ root: { boxShadow: '0 4px 16px rgba(0,0,0,0.16)', borderRadius: 8 } }}
        >
          {t.message}
        </MessageBar>
      ))}
    </div>
  );

  let body: JSX.Element;
  if (phase.kind === 'loading' || (phase.kind === 'ready' && !me)) {
    body = <Loading label="Loading Brillienta Timesheet…" />;
  } else if (phase.kind === 'error') {
    body = (
      <div className="ts-center">
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          Brillienta Timesheet could not load: {phase.message}
        </MessageBar>
        <PrimaryButton text="Try again" onClick={() => setBoot((b) => b + 1)} />
      </div>
    );
  } else if (phase.kind === 'notInstalled') {
    body = (
      <div className="ts-center">
        <img className="ts-logo is-lg" src={DEFAULT_LOGO} alt="" />
        <h2 style={{ margin: 0 }}>Brillienta Timesheet is not set up yet</h2>
        <div className="ts-muted">{host.siteError || 'A site owner needs to open this page once to create the timesheet lists.'}</div>
      </div>
    );
  } else if (phase.kind === 'upgrading') {
    body = (
      <div className="ts-setup">
        <h1>Updating Brillienta Timesheet to {props.appVersion}</h1>
        <p className="ts-muted">Adding what the new version needs to the timesheet lists. Your data is not changed.</p>
        <div className="ts-log">
          {phase.log.map((l, i) => (
            <div key={i} className={'is-' + l.l}>
              {l.m}
            </div>
          ))}
        </div>
      </div>
    );
  } else if (phase.kind === 'setup' && me) {
    body = (
      <AppCtx.Provider value={{ sp, svc, settings, graph, toast } as unknown as IAppContext}>
        <SetupWizard
          sp={sp}
          svc={svc}
          settings={settings}
          prefix={props.prefix}
          me={me}
          appVersion={props.appVersion}
          tenantId={props.tenantId}
          siteTitle={props.siteTitle}
          onDone={() => {
            setBoot((b) => b + 1);
            navigate('home');
          }}
        />
      </AppCtx.Provider>
    );
  } else if (license && !license.plan) {
    body = <LicenseLocked status={license} isAdmin={isAdmin} tenantId={props.tenantId} settings={settings} onChanged={() => void refreshLicense()} />;
  } else {
    const ctx: IAppContext = {
      sp,
      svc,
      settings,
      graph,
      timeOff,
      taskManager,
      notifier,
      workflow,
      reminders,
      config,
      savedConfig: saved,
      saveConfig: async (c: ITimesheetConfig, what?: string) => {
        const toSave = restoreLocked(c, saved, plan, license && license.addons);
        await settings.setJson(SETTING_KEYS.config, toSave);
        setSaved(toSave);
        await svc.log('Settings saved', what || '');
      },
      me: me as ICurrentUser,
      myUser,
      role,
      roleReason,
      isAdmin,
      isApprover,
      ref,
      dir: ref,
      reloadRef,
      license: license as ILicenseStatus,
      has: (f: Feature) => planHas(plan, f, license && license.addons),
      refreshLicense,
      usersOver,
      overLimit,
      admins,
      logo,
      saveLogo: async (dataUrl: string) => {
        await settings.set(SETTING_KEYS.brandLogo, dataUrl);
        setLogoStored(dataUrl);
      },
      security,
      setSecurity,
      approvalsCount,
      refreshApprovals: () => setApprovalsTick((t) => t + 1),
      notificationsVersion,
      bumpNotifications: () => setNotificationsVersion((v) => v + 1),
      prefix: props.prefix,
      appVersion: props.appVersion,
      tenantId: props.tenantId,
      host,
      linkBase,
      route,
      navigate,
      toast
    };
    body = (
      <AppCtx.Provider value={ctx}>
        <Shell frameClass={frameClass} frameStyle={frameStyle} siteTitle={props.siteTitle} />
      </AppCtx.Provider>
    );
  }

  const ready = phase.kind === 'ready' && !!me && !!license && !!license.plan;
  return (
    <ThemeProvider theme={fluentTheme} applyTo="none">
      <Customizer scopedSettings={layerSettings}>
        <div className={styles.root} style={themeVars} data-ts-top="1">
          {ready ? (
            body
          ) : (
            <div className={frameClass} style={{ ...(frameStyle || {}), display: 'block', overflowY: 'auto' }}>
              {body}
            </div>
          )}
          {toastArea}
        </div>
      </Customizer>
    </ThemeProvider>
  );
}
