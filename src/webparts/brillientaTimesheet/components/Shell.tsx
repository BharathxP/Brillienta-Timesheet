import * as React from 'react';
import { Callout, DefaultButton, DirectionalHint, Icon, IconButton, MessageBar, MessageBarType, TooltipHost } from '@fluentui/react';
import { useApp } from './AppContext';
import { Avatar } from './common/Ui';
import { ROLE_LABELS } from '../services/access';
import { describeStatus } from '../services/licensing/license';
import { INotification } from '../models/types';
import { relativeTime } from '../services/format';
import { Home } from './home/Home';
import { WeekSheet } from './timesheet/WeekSheet';
import { History } from './timesheet/History';
import { MyTasks } from './timesheet/MyTasks';
import { Approvals } from './approvals/Approvals';
import { Projects } from './projects/Projects';
import { ProjectDetail } from './projects/ProjectDetail';
import { Programs } from './projects/Programs';
import { Activities } from './projects/Activities';
import { People } from './people/People';
import { Reports } from './reports/Reports';
import { Dashboards } from './reports/Dashboards';
import { AdminPage } from './admin/AdminPage';
import { Support } from './support/Support';

interface INavItem {
  key: string;
  label: string;
  icon: string;
  section?: string;
  count?: number;
  locked?: boolean;
}

const LABELS: Record<string, string> = {
  home: 'Overview',
  timesheet: 'My timesheet',
  history: 'My timesheets',
  tasks: 'My tasks',
  approvals: 'Approvals',
  projects: 'Projects',
  programs: 'Programmes',
  activities: 'Activities',
  people: 'People & rates',
  reports: 'Reports',
  dashboards: 'Dashboards',
  admin: 'Admin',
  support: 'Support'
};

export function Shell(props: { frameClass: string; frameStyle?: React.CSSProperties; siteTitle: string }): JSX.Element {
  const app = useApp();
  const { role, isAdmin, isApprover, route, navigate, me, config, has, license } = app;
  const [navOpen, setNavOpen] = React.useState(false);
  const [collapsed, setCollapsed] = React.useState(false);
  const managesProjects = isAdmin || role === 'projectManager' || role === 'programManager';

  const items: INavItem[] = [
    { key: 'home', label: 'Overview', icon: 'ViewDashboard', section: 'Your work' },
    { key: 'timesheet', label: 'My timesheet', icon: 'Clock' },
    { key: 'history', label: 'My timesheets', icon: 'History' },
    { key: 'tasks', label: 'My tasks', icon: 'TaskManager' }
  ];
  if (isApprover) items.push({ key: 'approvals', label: 'Approvals', icon: 'DocumentApproval', section: 'Manage', count: app.approvalsCount });
  if (managesProjects || isApprover) items.push({ key: 'projects', label: 'Projects', icon: 'ProjectCollection', section: isApprover ? undefined : 'Manage' });
  if (isAdmin || role === 'programManager') items.push({ key: 'programs', label: 'Programmes', icon: 'Org', locked: !has('programManager') });
  if (isAdmin) {
    items.push({ key: 'activities', label: 'Activities', icon: 'Tag' });
    items.push({ key: 'people', label: 'People & rates', icon: 'People' });
  }
  if (isApprover) {
    items.push({ key: 'reports', label: 'Reports', icon: 'ReportDocument', section: 'Insights' });
    items.push({ key: 'dashboards', label: 'Dashboards', icon: 'BarChartVertical', locked: !has('customReports') });
  }
  if (isAdmin) {
    items.push({ key: 'admin', label: 'Admin', icon: 'Settings', section: 'Administration' });
    items.push({ key: 'support', label: 'Support', icon: 'Headset' });
  }

  const allowed = items.map((i) => i.key);
  const path = route.parts[0] || 'home';
  const effective = allowed.indexOf(path) >= 0 ? path : 'home';
  const p = route.parts;
  let page: JSX.Element;
  switch (effective) {
    case 'timesheet':
      page = <WeekSheet />;
      break;
    case 'history':
      page = <History />;
      break;
    case 'tasks':
      page = <MyTasks />;
      break;
    case 'approvals':
      page = <Approvals />;
      break;
    case 'projects':
      page = p[1] ? <ProjectDetail id={Number(p[1])} /> : <Projects />;
      break;
    case 'programs':
      page = <Programs />;
      break;
    case 'activities':
      page = <Activities />;
      break;
    case 'people':
      page = <People />;
      break;
    case 'reports':
      page = <Reports />;
      break;
    case 'dashboards':
      page = <Dashboards />;
      break;
    case 'admin':
      page = <AdminPage section={p[1]} />;
      break;
    case 'support':
      page = <Support />;
      break;
    default:
      page = <Home />;
  }
  let crumb = LABELS[effective] || 'Overview';
  if (effective === 'projects' && p[1]) crumb = (app.ref.projects.filter((x) => x.id === Number(p[1]))[0] || { title: 'Project' }).title;
  let lastSection: string | undefined;
  let pendingSection: string | undefined;

  const banner =
    license.state === 'trial' && isAdmin ? (
      <MessageBar messageBarType={MessageBarType.info} isMultiline={false} actions={<DefaultButton text="Activate a licence" onClick={() => navigate('admin/licence')} />}>
        {describeStatus(license)}.
      </MessageBar>
    ) : license.state === 'grace' ? (
      <MessageBar messageBarType={MessageBarType.severeWarning}>{describeStatus(license)}. Ask Brillienta to renew it.</MessageBar>
    ) : app.overLimit ? (
      <MessageBar messageBarType={MessageBarType.warning}>
        Your organisation has more active people than its plan includes ({license.users}). Ask an administrator to add users to the licence; until then you can fill in your timesheet but not submit it.
      </MessageBar>
    ) : isAdmin && app.usersOver > 0 ? (
      <MessageBar messageBarType={MessageBarType.warning} isMultiline={false} actions={<DefaultButton text="People" onClick={() => navigate('people')} />}>
        {app.usersOver} active {app.usersOver === 1 ? 'person is' : 'people are'} beyond the {license.users} users of your plan and cannot submit timesheets. Deactivate people or add users to the licence.
      </MessageBar>
    ) : null;

  return (
    <div className={props.frameClass} style={props.frameStyle}>
      <nav className={'ts-nav' + (navOpen ? ' is-open' : '') + (collapsed ? ' is-collapsed' : '')} aria-label="Brillienta Timesheet">
        <div className="ts-nav-brand">
          <img className="ts-logo" src={app.logo} alt="" />
          <span className="ts-nav-label">{config.appName === 'Brillienta Timesheet' ? 'timesheet' : config.appName}</span>
          <button type="button" className="ts-nav-collapse" onClick={() => setCollapsed(!collapsed)} aria-label={collapsed ? 'Expand navigation' : 'Collapse navigation'} aria-expanded={!collapsed}>
            <span aria-hidden="true">{collapsed ? '»' : '«'}</span>
          </button>
        </div>
        <div className="ts-nav-workspace" title={`${props.siteTitle} — ${app.roleReason}`}>
          <strong>{config.orgName || props.siteTitle}</strong>
          <span>{ROLE_LABELS[role]}</span>
        </div>
        {items.map((it) => {
          if (it.section) pendingSection = it.section;
          const header = pendingSection && pendingSection !== lastSection ? pendingSection : undefined;
          if (header) lastSection = header;
          pendingSection = undefined;
          return (
            <React.Fragment key={it.key}>
              {header && (
                <div className="ts-nav-section">
                  <span className="ts-nav-label">{header}</span>
                </div>
              )}
              <button
                className={'ts-nav-item' + (effective === it.key ? ' is-active' : '')}
                aria-current={effective === it.key ? 'page' : undefined}
                title={collapsed ? it.label : undefined}
                onClick={() => {
                  setNavOpen(false);
                  navigate(it.key);
                }}
              >
                <Icon iconName={it.icon} />
                <span className="ts-nav-label">{it.label}</span>
                {!!it.count && <span className="ts-nav-count">{it.count}</span>}
                {it.locked && <Icon iconName="Lock" className="ts-nav-lock" aria-label="Higher plan" />}
              </button>
            </React.Fragment>
          );
        })}
        <div className="ts-nav-footer" title={`Version ${app.appVersion}`}>
          <span className="ts-msicon">
            <i />
            <i />
            <i />
            <i />
          </span>
          <div className="ts-nav-label">
            Brillienta Timesheet
            <span className="ts-sub">
              {app.ref.users.filter((u) => u.active).length} people · v{app.appVersion}
            </span>
          </div>
        </div>
      </nav>
      <div className="ts-main" data-ts-main="1">
        <div className="ts-topbar">
          <IconButton className="ts-menu-btn" iconProps={{ iconName: 'GlobalNavButton' }} onClick={() => setNavOpen(!navOpen)} ariaLabel="Menu" />
          <span className="ts-crumb">
            Timesheet<span className="ts-crumb-sep">/</span>
            <strong>{crumb}</strong>
          </span>
          <div className="ts-topbar-actions">
            <DefaultButton iconProps={{ iconName: 'Add' }} text="Log time" onClick={() => navigate('timesheet')} className="ts-hide-sm" />
            {(isAdmin || me.canManageLists) && !app.host.inTeams && (
              <TooltipHost content="Site contents">
                <a className="ts-iconbtn" style={{ width: 36, height: 36 }} href={`${app.sp.webUrl}/_layouts/15/viewlsts.aspx`} target="_top" rel="noopener noreferrer" data-interception="off" aria-label="Site contents">
                  <Icon iconName="AllApps" />
                </a>
              </TooltipHost>
            )}
            <NotificationBell />
            <span title={`${me.title} — ${ROLE_LABELS[role]}`}>
              <Avatar name={me.title} email={me.email} />
            </span>
          </div>
        </div>
        {banner}
        <div className="ts-page" key={route.path} onClick={() => navOpen && setNavOpen(false)}>
          {page}
        </div>
      </div>
    </div>
  );
}

/** In-app notifications (TS_Notifications), refreshed every two minutes. */
function NotificationBell(): JSX.Element {
  const { svc, me, navigate, notificationsVersion } = useApp();
  const [items, setItems] = React.useState<INotification[]>([]);
  const [open, setOpen] = React.useState(false);
  const btn = React.useRef<HTMLButtonElement>(null);
  const load = React.useCallback(() => {
    svc
      .notifications(me.email)
      .then(setItems)
      .catch(() => undefined);
  }, [svc, me.email]);
  React.useEffect(() => {
    load();
    const t = window.setInterval(() => !document.hidden && load(), 120000);
    return () => window.clearInterval(t);
  }, [load, notificationsVersion]);
  const unread = items.filter((i) => !i.read);
  const markAll = (): void => {
    if (!unread.length) return;
    svc
      .markRead(unread.map((i) => i.id))
      .then(load)
      .catch(() => undefined);
  };
  return (
    <span className="ts-bell">
      <button ref={btn} className="ts-iconbtn" style={{ width: 36, height: 36 }} onClick={() => setOpen(!open)} aria-label={`Notifications, ${unread.length} unread`}>
        <Icon iconName="Ringer" />
      </button>
      {unread.length > 0 && <span className="ts-bell-count">{unread.length > 9 ? '9+' : unread.length}</span>}
      {open && (
        <Callout target={btn.current} onDismiss={() => setOpen(false)} directionalHint={DirectionalHint.bottomRightEdge} isBeakVisible={false} gapSpace={6}>
          <div className="ts-bell-menu">
            <div className="ts-row" style={{ padding: '10px 14px', borderBottom: '1px solid var(--ts-border)' }}>
              <strong>Notifications</strong>
              <span className="ts-spacer" />
              {unread.length > 0 && (
                <button className="ts-link ts-small" onClick={markAll}>
                  Mark all read
                </button>
              )}
            </div>
            {!items.length && <div className="ts-muted ts-small" style={{ padding: 16 }}>Nothing yet.</div>}
            {items.map((n) => (
              <button
                key={n.id}
                className={'ts-bell-item' + (n.read ? '' : ' is-unread')}
                onClick={() => {
                  setOpen(false);
                  if (!n.read)
                    svc
                      .markRead([n.id])
                      .then(load)
                      .catch(() => undefined);
                  if (n.link) navigate(n.link);
                }}
              >
                <div className="ts-bell-title">{n.title}</div>
                <div className="ts-bell-body">{n.body}</div>
                <div className="ts-small ts-muted">{relativeTime(n.created)}</div>
              </button>
            ))}
          </div>
        </Callout>
      )}
    </span>
  );
}
