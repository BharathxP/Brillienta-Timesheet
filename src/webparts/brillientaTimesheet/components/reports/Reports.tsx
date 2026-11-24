import * as React from 'react';
import { DefaultButton, Dropdown, IDropdownOption, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { Dimension, IReportDefinition, IReportFilter, ISavedReport, Measure, SheetStatus } from '../../models/types';
import { DIMENSIONS, IFact, MEASURES, totals, utilisation } from '../../services/reports';
import { PERIODS, formatDay, periodRange, today, weekDays, weeksIn } from '../../services/dates';
import { managedProjects, reportsOf, sameEmail } from '../../services/access';
import { isTrackingDay } from '../../services/rules';
import { hours, money, percent } from '../../services/format';
import { Card, DateField, ErrorBox, Kpi, Loading, Locked, PageHeader, Tabs, useBusy, useConfirm, useLoader } from '../common/Ui';
import { useFacts } from './useFacts';
import { ReportView } from './ReportView';
import { BarList } from '../common/Charts';

type Tab = 'project' | 'team' | 'billable' | 'cost' | 'utilisation' | 'custom';

export function defaultFilter(period: string, weekStart: number): IReportFilter {
  const r = periodRange(period, today(), weekStart);
  return { from: r.from, to: r.to, projectIds: [], userEmails: [], billable: 'all', statuses: [] };
}

/** What a viewer may report on: admins everything; others the projects they manage and the people who report to them. */
export function useScope(): { all: boolean; projectIds: number[]; emails: string[]; filter: (f: IFact) => boolean } {
  const { isAdmin, role, me, ref } = useApp();
  return React.useMemo(() => {
    if (isAdmin) return { all: true, projectIds: [], emails: [], filter: () => true };
    const projectIds = managedProjects(role, me.email, ref).map((p) => p.id);
    const emails = reportsOf(ref.users, me.email).map((u) => u.email.toLowerCase());
    return { all: false, projectIds, emails, filter: (f: IFact) => projectIds.indexOf(f.projectId) >= 0 || emails.indexOf(f.userEmail.toLowerCase()) >= 0 };
  }, [isAdmin, role, me.email, ref]);
}

export function FilterBar(props: { filter: IReportFilter; period: string; onChange: (f: IReportFilter, period: string) => void; people?: boolean; projects?: boolean }): JSX.Element {
  const { ref, config } = useApp();
  const scope = useScope();
  const f = props.filter;
  const projects = scope.all ? ref.projects : ref.projects.filter((p) => scope.projectIds.indexOf(p.id) >= 0);
  const users = scope.all ? ref.users : ref.users.filter((u) => scope.emails.indexOf(u.email.toLowerCase()) >= 0);
  const multi = (keys: Array<string | number>, k: string | number, on?: boolean): Array<string | number> => (on ? keys.concat(k) : keys.filter((x) => x !== k));
  return (
    <div className="ts-filters">
      <Dropdown
        label="Period"
        selectedKey={props.period}
        options={(PERIODS as IDropdownOption[]).concat([{ key: '', text: 'Custom dates' }])}
        onChange={(_, o) => {
          if (!o) return;
          const k = String(o.key);
          props.onChange(k ? { ...f, ...periodRange(k, today(), config.weekStart) } : f, k);
        }}
      />
      <DateField label="From" value={f.from} minWidth={150} onChange={(v) => props.onChange({ ...f, from: v }, '')} />
      <DateField label="To" value={f.to} minWidth={150} onChange={(v) => props.onChange({ ...f, to: v }, '')} />
      {props.projects !== false && (
        <Dropdown
          label="Projects"
          multiSelect
          placeholder="All projects"
          selectedKeys={f.projectIds}
          options={projects.map((p) => ({ key: p.id, text: p.title }))}
          onChange={(_, o) => o && props.onChange({ ...f, projectIds: multi(f.projectIds, Number(o.key), o.selected) as number[] }, props.period)}
          styles={{ root: { minWidth: 200 } }}
        />
      )}
      {props.people !== false && (
        <Dropdown
          label="People"
          multiSelect
          placeholder="Everyone"
          selectedKeys={f.userEmails}
          options={users.map((u) => ({ key: u.email.toLowerCase(), text: u.name }))}
          onChange={(_, o) => o && props.onChange({ ...f, userEmails: multi(f.userEmails, String(o.key), o.selected) as string[] }, props.period)}
          styles={{ root: { minWidth: 200 } }}
        />
      )}
      <Dropdown
        label="Billable"
        selectedKey={f.billable}
        options={[
          { key: 'all', text: 'All hours' },
          { key: 'billable', text: 'Billable only' },
          { key: 'nonBillable', text: 'Non-billable only' }
        ]}
        onChange={(_, o) => o && props.onChange({ ...f, billable: o.key as IReportFilter['billable'] }, props.period)}
      />
      <Dropdown
        label="Timesheets"
        multiSelect
        placeholder="Any status"
        selectedKeys={f.statuses}
        options={['Draft', 'Submitted', 'Approved', 'Rejected'].map((s) => ({ key: s, text: s }))}
        onChange={(_, o) => o && props.onChange({ ...f, statuses: multi(f.statuses, String(o.key), o.selected) as SheetStatus[] }, props.period)}
      />
    </div>
  );
}

export function Reports(): JSX.Element {
  const { config, has, route } = useApp();
  const [tab, setTab] = React.useState<Tab>((route.query.tab as Tab) || 'project');
  const [period, setPeriod] = React.useState('thisMonth');
  const [filter, setFilter] = React.useState<IReportFilter>(() => defaultFilter('thisMonth', config.weekStart));
  const scope = useScope();
  const data = useFacts(filter);
  const facts = data.facts.filter(scope.filter);
  const money0 = has('projectCost') || has('userCost');
  return (
    <>
      <PageHeader title="Reports" sub={scope.all ? 'Every project and person.' : 'Projects you manage and the people who report to you.'} />
      <Tabs
        value={tab}
        onChange={setTab}
        tabs={[
          { key: 'project', text: 'Project report', icon: 'ProjectCollection', locked: !has('projectReport') },
          { key: 'team', text: 'My team', icon: 'People', locked: !has('teamReport') },
          { key: 'billable', text: 'Billable / non-billable', icon: 'Money', locked: !has('billable') },
          { key: 'cost', text: 'Cost & margin', icon: 'Calculator', locked: !money0 },
          { key: 'utilisation', text: 'Utilisation', icon: 'Chart' },
          { key: 'custom', text: 'Custom report', icon: 'BuildDefinition', locked: !has('customReports') }
        ]}
      />
      <FilterBar
        filter={filter}
        period={period}
        onChange={(f, p) => {
          setFilter(f);
          setPeriod(p);
        }}
      />
      {data.error && <ErrorBox error={data.error} onRetry={data.reload} />}
      {data.loading ? <Loading /> : <ReportTab tab={tab} facts={facts} lookups={data.lookups} filter={filter} />}
    </>
  );
}

function ReportTab(props: { tab: Tab; facts: IFact[]; lookups: ReturnType<typeof useFacts>['lookups']; filter: IReportFilter }): JSX.Element {
  const { has, config, ref, me, isAdmin } = useApp();
  const { facts, lookups } = props;
  const t = totals(facts);
  const money0 = has('projectCost') || has('userCost');
  const fmtMoney = (v: number): string => money(v, config.currency);
  const kpis = (
    <div className="ts-kpis">
      <Kpi label="Hours" value={hours(t.hours)} icon="Clock" tint={1} />
      <Kpi label="Billable" value={hours(t.billableHours)} icon="Money" tint={5} hint={t.hours ? percent((t.billableHours / t.hours) * 100) : undefined} />
      <Kpi label="Non-billable" value={hours(t.nonBillableHours)} icon="Blocked2" tint={2} />
      {money0 && <Kpi label="Cost" value={fmtMoney(t.cost)} icon="Calculator" tint={3} />}
      {money0 && <Kpi label="Billable amount" value={fmtMoney(t.revenue)} icon="Financial" tint={4} />}
    </div>
  );
  switch (props.tab) {
    case 'project':
      if (!has('projectReport')) return <Locked feature="projectReport" />;
      return (
        <>
          {kpis}
          <Card title="Hours by project, task and person" icon="ProjectCollection">
            <ReportView facts={facts} lookups={lookups} title="Project report" def={{ groupBy: ['project', 'task', 'user'], measures: money0 ? ['hours', 'billableHours', 'nonBillableHours', 'cost', 'revenue'] : ['hours', 'billableHours', 'nonBillableHours'], chart: 'bar' }} />
          </Card>
        </>
      );
    case 'team': {
      if (!has('teamReport')) return <Locked feature="teamReport" />;
      const team = isAdmin && !reportsOf(ref.users, me.email).length ? ref.users.map((u) => u.email.toLowerCase()) : reportsOf(ref.users, me.email).map((u) => u.email.toLowerCase());
      const tf = facts.filter((f) => team.indexOf(f.userEmail.toLowerCase()) >= 0);
      return (
        <>
          <div className="ts-kpis">
            <Kpi label="People in my team" value={team.length} icon="People" tint={1} />
            <Kpi label="Team hours" value={hours(totals(tf).hours)} icon="Clock" tint={2} />
            <Kpi label="Billable" value={hours(totals(tf).billableHours)} icon="Money" tint={5} />
          </div>
          <Card title="My team report" icon="People">
            {!team.length ? (
              <div className="ts-muted">Nobody reports to you. Managers are set in People &amp; rates.</div>
            ) : (
              <ReportView facts={tf} lookups={lookups} title="My team report" def={{ groupBy: ['user', 'project'], measures: ['hours', 'billableHours', 'nonBillableHours'], chart: 'bar' }} />
            )}
          </Card>
          <TeamStatus emails={team} filter={props.filter} />
        </>
      );
    }
    case 'billable':
      return (
        <>
          {kpis}
          <div className="ts-grid ts-grid-2">
            <Card title="Billable vs non-billable" icon="Money">
              <ReportView facts={facts} lookups={lookups} title="Billable hours" def={{ groupBy: ['billable'], measures: ['hours'], chart: 'donut' }} noExport />
            </Card>
            <Card title="By week" icon="Calendar">
              <ReportView facts={facts} lookups={lookups} title="Billable by week" def={{ groupBy: ['week'], measures: ['billableHours', 'nonBillableHours'], chart: 'column' }} noExport />
            </Card>
          </div>
          <Card title="By project and person" icon="ReportDocument" style={{ marginTop: 16 }}>
            <ReportView facts={facts} lookups={lookups} title="Billable report" def={{ groupBy: ['project', 'user'], measures: ['hours', 'billableHours', 'nonBillableHours', ...(money0 ? (['revenue'] as Measure[]) : [])], chart: 'table' }} />
          </Card>
        </>
      );
    case 'cost':
      if (!money0) return <Locked feature="projectCost" />;
      return (
        <>
          <div className="ts-kpis">
            <Kpi label="Cost" value={fmtMoney(t.cost)} icon="Calculator" tint={3} />
            <Kpi label="Billable amount" value={fmtMoney(t.revenue)} icon="Financial" tint={4} />
            <Kpi label="Margin" value={fmtMoney(t.margin)} icon="Chart" tint={1} tone={t.margin < 0 ? 'danger' : 'success'} hint={t.revenue ? percent((t.margin / t.revenue) * 100) : undefined} />
          </div>
          <Card title="Cost, billable amount and margin" icon="Calculator">
            <ReportView facts={facts} lookups={lookups} title="Cost report" def={{ groupBy: ['client', 'project'], measures: ['hours', 'cost', 'revenue', 'margin'], chart: 'bar' }} />
          </Card>
          <div className="ts-muted ts-small" style={{ marginTop: 8 }}>
            Submitted weeks use the rates captured when they were submitted; drafts use today&apos;s rates ({config.costMode === 'userFirst' ? 'person’s hourly cost, else the project’s flat cost' : config.costMode === 'projectFirst' ? 'project’s flat cost, else the person’s' : config.costMode === 'userOnly' ? 'person’s hourly cost' : 'project’s flat hourly cost'}).
          </div>
        </>
      );
    case 'utilisation': {
      const users = ref.users.filter((u) => u.active && (!props.filter.userEmails.length || props.filter.userEmails.indexOf(u.email.toLowerCase()) >= 0));
      const days: string[] = [];
      weeksIn(props.filter.from, props.filter.to, config.weekStart).forEach((w) => weekDays(w).forEach((d) => d >= props.filter.from && d <= props.filter.to && days.push(d)));
      const rows = users
        .map((u) => {
          const mine = facts.filter((f) => sameEmail(f.userEmail, u.email));
          const available = days.filter((d) => isTrackingDay(config, d, u)).length * config.standardDay;
          const b = mine.filter((f) => f.billable).reduce((s, f) => s + f.hours, 0);
          return { label: u.name, value: utilisation(b, available), all: mine.reduce((s, f) => s + f.hours, 0), available };
        })
        .filter((r) => r.all > 0 || r.available > 0)
        .sort((a, b) => b.value - a.value);
      return (
        <Card title="Billable utilisation" icon="Chart">
          <p className="ts-muted ts-small" style={{ marginTop: 0 }}>
            Billable hours ÷ available hours (tracking days × {config.standardDay} h) between {formatDay(props.filter.from, true)} and {formatDay(props.filter.to, true)}.
          </p>
          <BarList items={rows.slice(0, 40)} format={(v) => percent(v)} />
        </Card>
      );
    }
    default:
      return has('customReports') ? <CustomReport facts={facts} lookups={lookups} filter={props.filter} /> : <Locked feature="customReports" />;
  }
}

/** Who in the team has not submitted the weeks in the period. */
function TeamStatus(props: { emails: string[]; filter: IReportFilter }): JSX.Element | null {
  const { svc, config, ref } = useApp();
  const weeks = weeksIn(props.filter.from, props.filter.to, config.weekStart).filter((w) => w <= today());
  const { data } = useLoader(() => (props.emails.length && weeks.length ? svc.weeks(weeks[0], weeks[weeks.length - 1], props.emails).then((r) => r.sheets) : Promise.resolve([])), [props.emails.join(), weeks.join()]);
  if (!props.emails.length || !weeks.length || !data) return null;
  return (
    <Card title="Timesheet status" icon="DocumentApproval" style={{ marginTop: 16 }}>
      <div className="ts-table-wrap">
        <table className="ts-table">
          <thead>
            <tr>
              <th>Person</th>
              {weeks.map((w) => (
                <th key={w}>{formatDay(w)}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {props.emails.map((e) => (
              <tr key={e} style={{ cursor: 'default' }}>
                <td>{(ref.users.filter((u) => sameEmail(u.email, e))[0] || { name: e }).name}</td>
                {weeks.map((w) => {
                  const s = data.filter((x) => sameEmail(x.userEmail, e) && x.weekStart === w)[0];
                  return (
                    <td key={w}>
                      <span className={'ts-status is-' + (s ? s.status.toLowerCase() : 'rejected')}>{s ? `${s.status} · ${hours(s.totalHours)}` : 'Missing'}</span>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Card>
  );
}

/** Custom reports: choose groupings, measures and chart; save them (privately or shared) — Enterprise. */
function CustomReport(props: { facts: IFact[]; lookups: ReturnType<typeof useFacts>['lookups']; filter: IReportFilter }): JSX.Element {
  const { svc, me, config, toast } = useApp();
  const [def, setDef] = React.useState<IReportDefinition>({ groupBy: ['project', 'user'], measures: ['hours', 'billableHours'], chart: 'bar', filter: props.filter, period: 'thisMonth' });
  const [title, setTitle] = React.useState('');
  const [shared, setShared] = React.useState(false);
  const [busy, run] = useBusy();
  const [ask, dialog] = useConfirm();
  const saved = useLoader(() => svc.reports(me.email), []);
  const dims: IDropdownOption[] = (DIMENSIONS as IDropdownOption[]).concat(config.customColumns.filter((c) => c.inReports && c.level === 'entry').map((c) => ({ key: 'custom:' + c.key, text: c.label })));
  const reports = (saved.data || []).filter((r) => r.kind === 'report');
  const save = (): Promise<void> =>
    run(async () => {
      if (!title.trim()) throw new Error('Give the report a name.');
      await svc.saveReport({ title: title.trim(), owner: me.email, shared, kind: 'report', definition: { ...def, filter: props.filter } });
      saved.reload();
    }, 'Report saved.');
  const open = (r: ISavedReport): void => {
    const d = r.definition as IReportDefinition;
    setDef(d);
    setTitle(r.title);
    toast(`Opened ${r.title}. Its period applies when it is shown on a dashboard; here it uses the filters above.`, 'info');
  };
  return (
    <div className="ts-split">
      <Card title={title || 'Custom report'} icon="BuildDefinition">
        <div className="ts-filters">
          <Dropdown label="Group by" multiSelect selectedKeys={def.groupBy as string[]} options={dims} onChange={(_, o) => o && setDef({ ...def, groupBy: o.selected ? def.groupBy.concat(String(o.key) as Dimension) : def.groupBy.filter((x) => x !== o.key) })} styles={{ root: { minWidth: 220 } }} />
          <Dropdown label="Measures" multiSelect selectedKeys={def.measures} options={MEASURES.map((m) => ({ key: m.key, text: m.text }))} onChange={(_, o) => o && setDef({ ...def, measures: o.selected ? def.measures.concat(o.key as Measure) : def.measures.filter((x) => x !== o.key) })} styles={{ root: { minWidth: 220 } }} />
          <Dropdown
            label="Chart"
            selectedKey={def.chart}
            options={[
              { key: 'table', text: 'Table only' },
              { key: 'bar', text: 'Bars' },
              { key: 'column', text: 'Columns' },
              { key: 'line', text: 'Line' },
              { key: 'donut', text: 'Donut' }
            ]}
            onChange={(_, o) => o && setDef({ ...def, chart: o.key as IReportDefinition['chart'] })}
          />
        </div>
        <ReportView facts={props.facts} lookups={props.lookups} title={title || 'Custom report'} def={def} />
      </Card>
      <div className="ts-stack">
        <Card title="Save" icon="Save">
          <TextField label="Name" value={title} onChange={(_, v) => setTitle(v || '')} />
          <Toggle label="Share with everyone who sees reports" checked={shared} onChange={(_, v) => setShared(!!v)} />
          <PrimaryButton text="Save report" onClick={() => void save()} disabled={busy} />
        </Card>
        <Card title="Saved reports" icon="FabricFolder">
          {!reports.length && <div className="ts-muted ts-small">None yet.</div>}
          {reports.map((r) => (
            <div key={r.id} className="ts-row" style={{ justifyContent: 'space-between' }}>
              <button className="ts-link" onClick={() => open(r)}>
                {r.title}
              </button>
              <span className="ts-row">
                {r.shared && <span className="ts-chip is-soft">Shared</span>}
                {sameEmail(r.owner, me.email) && (
                  <DefaultButton
                    iconProps={{ iconName: 'Delete' }}
                    ariaLabel="Delete"
                    styles={{ root: { minWidth: 28, padding: 0 } }}
                    onClick={async () => {
                      if (await ask(`Delete ${r.title}?`, '', 'Delete'))
                        await run(async () => {
                          await svc.deleteReport(r.id);
                          saved.reload();
                        });
                    }}
                  />
                )}
              </span>
            </div>
          ))}
        </Card>
      </div>
      {dialog}
    </div>
  );
}
