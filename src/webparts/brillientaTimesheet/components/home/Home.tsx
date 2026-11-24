import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react';
import { useApp, buildQuery } from '../AppContext';
import { addDays, daysBetween, formatDay, formatWeek, periodRange, today, weekStartOf } from '../../services/dates';
import { isTrackingDay, limitsFor, weeklyTarget } from '../../services/rules';
import { managedProjects, sameEmail } from '../../services/access';
import { hours } from '../../services/format';
import { Card, Empty, Kpi, Progress, StatusBadge, useLoader } from '../common/Ui';
import { BarList } from '../common/Charts';

export function Home(): JSX.Element {
  const { svc, me, myUser, config, navigate, ref, role, isApprover, approvalsCount, isAdmin } = useApp();
  const thisWeek = weekStartOf(today(), config.weekStart);
  const month = periodRange('thisMonth', today(), config.weekStart);
  const { data } = useLoader(async () => {
    const [sheets, mine] = await Promise.all([svc.sheetsOf(me.email, 12), svc.weeks(weekStartOf(month.from, config.weekStart), thisWeek, [me.email])]);
    return { sheets, entries: mine.entries };
  }, [me.email]);
  const used = useLoader(() => (isApprover ? svc.hoursByProject() : Promise.resolve({} as Record<number, number>)), [isApprover]);
  const current = data ? data.sheets.filter((s) => s.weekStart === thisWeek)[0] : undefined;
  const weekHours = data ? data.entries.filter((e) => e.weekStart === thisWeek).reduce((s, e) => s + e.total, 0) : 0;
  const days = [0, 1, 2, 3, 4, 5, 6].map((i) => addDays(thisWeek, i));
  const limits = limitsFor(config, myUser);
  const target = weeklyTarget(config, days, myUser, limits.weeklyMin);
  const monthFacts = data
    ? data.entries.reduce(
        (acc, e) => {
          Object.keys(e.hours).forEach((d) => {
            if (d < month.from || d > month.to) return;
            acc.total += e.hours[d];
            if (e.billable) acc.billable += e.hours[d];
            acc.byProject[e.projectId] = (acc.byProject[e.projectId] || 0) + e.hours[d];
          });
          return acc;
        },
        { total: 0, billable: 0, byProject: {} as Record<number, number> }
      )
    : { total: 0, billable: 0, byProject: {} as Record<number, number> };
  const pending = data ? data.sheets.filter((s) => s.status === 'Rejected' || (s.status === 'Draft' && s.weekStart < thisWeek)) : [];
  const missing = data ? [1, 2, 3, 4].map((i) => addDays(thisWeek, -7 * i)).filter((w) => !data.sheets.some((s) => s.weekStart === w)) : [];
  const myTasks = ref.tasks.filter((t) => t.status !== 'Done' && t.assigneeEmails.some((e) => sameEmail(e, me.email))).sort((a, b) => (a.dueDate || '9').localeCompare(b.dueDate || '9'));
  const projects = managedProjects(role, me.email, ref).filter((p) => p.status === 'Active');
  const ending = projects.filter((p) => p.endDate && daysBetween(today(), p.endDate) <= 21).sort((a, b) => a.endDate.localeCompare(b.endDate));
  const pct = target ? Math.min(1, weekHours / target) : 0;
  const R = 54;
  const C = 2 * Math.PI * R;
  const first = (me.title || '').split(' ')[0];
  const hour = new Date().getHours();
  return (
    <>
      <div className="ts-hero-week">
        <div className="ts-ring" aria-label={`${hours(weekHours)} of ${hours(target)} hours this week`}>
          <svg viewBox="0 0 128 128" width="128" height="128">
            <circle cx="64" cy="64" r={R} fill="none" stroke="#e7efde" strokeWidth="12" />
            <circle cx="64" cy="64" r={R} fill="none" stroke="#9dc56a" strokeWidth="12" strokeLinecap="round" strokeDasharray={`${pct * C} ${C}`} transform="rotate(-90 64 64)" />
          </svg>
          <div className="ts-ring-text">
            <div>
              {hours(weekHours)}
              <span>of {hours(target)} h</span>
            </div>
          </div>
        </div>
        <div style={{ flex: 1, minWidth: 260 }}>
          <p className="ts-eyebrow">
            <span className="ts-star">✳</span> {formatWeek(thisWeek)}
          </p>
          <h1>
            Good {hour < 12 ? 'morning' : hour < 18 ? 'afternoon' : 'evening'}, <span className="ts-underline">{first}</span>
          </h1>
          <div className="ts-row" style={{ marginTop: 10 }}>
            <StatusBadge status={current ? current.status : 'Draft'} />
            <span className="ts-muted">{current && current.status === 'Approved' ? 'This week is approved.' : weekHours >= target ? 'Your week is complete — submit it when you are ready.' : `${hours(Math.max(0, target - weekHours))} h to go this week.`}</span>
          </div>
          <div className="ts-row" style={{ marginTop: 14 }}>
            <PrimaryButton iconProps={{ iconName: 'Clock' }} text="Fill in this week" onClick={() => navigate('timesheet')} />
            {isApprover && approvalsCount > 0 && <DefaultButton iconProps={{ iconName: 'DocumentApproval' }} text={`${approvalsCount} to approve`} onClick={() => navigate('approvals')} />}
          </div>
        </div>
      </div>
      <div className="ts-kpis">
        <Kpi label="This week" value={`${hours(weekHours)} h`} icon="Clock" tint={1} onClick={() => navigate('timesheet')} />
        <Kpi label="This month" value={`${hours(monthFacts.total)} h`} icon="Calendar" tint={2} />
        <Kpi label="Billable this month" value={`${hours(monthFacts.billable)} h`} icon="Money" tint={5} hint={monthFacts.total ? `${Math.round((monthFacts.billable / monthFacts.total) * 100)}% billable` : undefined} />
        <Kpi label="Weeks to fix" value={pending.length + missing.length} icon="Warning" tint={3} tone={pending.length + missing.length ? 'warning' : undefined} onClick={() => navigate('history')} />
        {isApprover && <Kpi label="Waiting for my approval" value={approvalsCount} icon="DocumentApproval" tint={4} onClick={() => navigate('approvals')} />}
      </div>
      <div className="ts-grid ts-grid-2">
        <Card title="Weeks that need you" icon="History">
          {!pending.length && !missing.length ? (
            <Empty icon="CheckMark" title="All caught up" />
          ) : (
            <div className="ts-stack" style={{ gap: 8 }}>
              {pending.map((s) => (
                <div key={s.id} className="ts-row" style={{ justifyContent: 'space-between' }}>
                  <span>{formatWeek(s.weekStart)}</span>
                  <span className="ts-row">
                    <StatusBadge status={s.status} />
                    <DefaultButton text="Open" onClick={() => navigate('timesheet' + buildQuery({ week: s.weekStart }))} />
                  </span>
                </div>
              ))}
              {missing.map((w) => (
                <div key={w} className="ts-row" style={{ justifyContent: 'space-between' }}>
                  <span>{formatWeek(w)}</span>
                  <span className="ts-row">
                    <span className="ts-status is-rejected">Not started</span>
                    <DefaultButton text="Fill in" onClick={() => navigate('timesheet' + buildQuery({ week: w }))} />
                  </span>
                </div>
              ))}
            </div>
          )}
        </Card>
        <Card title="Where my time went this month" icon="BarChartHorizontal">
          <BarList
            items={Object.keys(monthFacts.byProject)
              .map((k) => {
                const p = ref.projects.filter((x) => x.id === Number(k))[0];
                return { label: p ? p.title : 'Time off', value: Math.round(monthFacts.byProject[Number(k)] * 100) / 100, color: (p && p.color) || undefined };
              })
              .sort((a, b) => b.value - a.value)
              .slice(0, 8)}
            format={(v) => `${hours(v)} h`}
            empty="No time logged this month yet."
          />
        </Card>
        <Card title="My tasks" icon="TaskManager" actions={<button className="ts-link" onClick={() => navigate('tasks')}>All tasks</button>}>
          {!myTasks.length ? (
            <div className="ts-muted ts-small">No open tasks assigned to you.</div>
          ) : (
            myTasks.slice(0, 6).map((t) => (
              <div key={t.id} className="ts-row" style={{ justifyContent: 'space-between', padding: '4px 0' }}>
                <span>
                  <strong>{t.title}</strong> <span className="ts-muted ts-small">· {(ref.projects.filter((p) => p.id === t.projectId)[0] || { title: '' }).title}</span>
                </span>
                <span className={t.dueDate && t.dueDate < today() ? 'ts-status-bad ts-small' : 'ts-muted ts-small'}>{t.dueDate ? formatDay(t.dueDate) : ''}</span>
              </div>
            ))
          )}
        </Card>
        {(isApprover || isAdmin) && (
          <Card title="My projects" icon="ProjectCollection" actions={<button className="ts-link" onClick={() => navigate('projects')}>All projects</button>}>
            {!projects.length ? (
              <div className="ts-muted ts-small">You do not manage any project.</div>
            ) : (
              <div className="ts-stack" style={{ gap: 10 }}>
                {(ending.length ? ending : projects).slice(0, 6).map((p) => (
                  <div key={p.id} style={{ cursor: 'pointer' }} onClick={() => navigate(`projects/${p.id}`)}>
                    <div className="ts-row" style={{ justifyContent: 'space-between' }}>
                      <span>
                        <span className="ts-swatch" style={{ background: p.color || '#c8d0d2' }} />
                        <strong>{p.title}</strong>
                      </span>
                      <span className="ts-small ts-muted">{p.endDate ? `ends ${formatDay(p.endDate)}` : ''}</span>
                    </div>
                    {p.plannedHours > 0 && <Progress value={(used.data || {})[p.id] || 0} max={p.plannedHours} label={`${hours((used.data || {})[p.id] || 0)} of ${hours(p.plannedHours)} planned hours`} />}
                  </div>
                ))}
              </div>
            )}
          </Card>
        )}
      </div>
    </>
  );
}
