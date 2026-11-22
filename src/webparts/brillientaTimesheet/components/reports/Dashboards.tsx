import * as React from 'react';
import { DefaultButton, Dropdown, Panel, PanelType, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { Dimension, IDashboardDefinition, IDashboardWidget, IReportDefinition, ISavedReport, Measure } from '../../models/types';
import { DIMENSIONS, MEASURES } from '../../services/reports';
import { PERIODS, formatDay, periodRange, today } from '../../services/dates';
import { sameEmail } from '../../services/access';
import { Card, Empty, Loading, Locked, PageHeader, useBusy, useConfirm, useLoader } from '../common/Ui';
import { useFacts } from './useFacts';
import { ReportView } from './ReportView';
import { useScope } from './Reports';

const id = (): string => Math.random().toString(36).substring(2, 9);

function widget(title: string, groupBy: Dimension[], measures: Measure[], chart: IReportDefinition['chart'], period: IReportDefinition['period'], width: 1 | 2 | 3): IDashboardWidget {
  return { id: id(), title, width, report: { groupBy, measures, chart, period, filter: { from: '', to: '', projectIds: [], userEmails: [], billable: 'all', statuses: [] } } };
}

/** A starting dashboard for new ones. */
export function starterDashboard(): IDashboardDefinition {
  return {
    widgets: [
      widget('Hours by project this month', ['project'], ['hours'], 'bar', 'thisMonth', 1),
      widget('Billable vs non-billable', ['billable'], ['hours'], 'donut', 'thisMonth', 1),
      widget('Hours by activity', ['activity'], ['hours'], 'bar', 'thisMonth', 1),
      widget('Weekly trend (90 days)', ['week'], ['billableHours', 'nonBillableHours'], 'column', 'last90', 2),
      widget('Top people this month', ['user'], ['hours', 'billableHours'], 'table', 'thisMonth', 1),
      widget('Cost and margin by client', ['client'], ['cost', 'revenue', 'margin'], 'table', 'thisQuarter', 3)
    ]
  };
}

/** Custom dashboards (Enterprise): a grid of report widgets, each with its own period, saved privately or shared. */
export function Dashboards(): JSX.Element {
  const { has, svc, me } = useApp();
  const list = useLoader(() => (has('customReports') ? svc.reports(me.email) : Promise.resolve([] as ISavedReport[])), []);
  const [current, setCurrent] = React.useState<number | 'new' | undefined>();
  const [editing, setEditing] = React.useState<ISavedReport | undefined>();
  const [busy, run] = useBusy();
  const [ask, dialog] = useConfirm();
  if (!has('customReports'))
    return (
      <>
        <PageHeader title="Dashboards" />
        <Locked feature="customReports">Build your own dashboards from report widgets, and share them.</Locked>
      </>
    );
  if (list.loading || !list.data) return <Loading />;
  const dashboards = list.data.filter((r) => r.kind === 'dashboard');
  const shown: ISavedReport =
    dashboards.filter((d) => d.id === current)[0] || dashboards[0] || { id: 0, title: 'Overview dashboard', owner: me.email, shared: true, kind: 'dashboard', definition: starterDashboard() };
  const mine = sameEmail(shown.owner, me.email);
  const save = (d: ISavedReport): Promise<void> =>
    run(async () => {
      const newId = await svc.saveReport(d);
      setEditing(undefined);
      setCurrent(newId);
      list.reload();
    }, 'Dashboard saved.');
  return (
    <>
      <PageHeader
        title="Dashboards"
        sub={shown.id ? `${shown.title}${shown.shared ? ' · shared' : ''}` : 'A starter dashboard. Save it to keep and change it.'}
        actions={
          <>
            {dashboards.length > 0 && <Dropdown selectedKey={shown.id} options={dashboards.map((d) => ({ key: d.id, text: d.title }))} onChange={(_, o) => o && setCurrent(Number(o.key))} styles={{ root: { width: 220 } }} />}
            <DefaultButton iconProps={{ iconName: 'Add' }} text="New dashboard" onClick={() => setEditing({ id: 0, title: 'New dashboard', owner: me.email, shared: false, kind: 'dashboard', definition: starterDashboard() })} />
            {(mine || !shown.id) && <PrimaryButton iconProps={{ iconName: 'Edit' }} text={shown.id ? 'Edit' : 'Save as my dashboard'} onClick={() => setEditing({ ...shown })} />}
            {mine && shown.id > 0 && (
              <DefaultButton
                iconProps={{ iconName: 'Delete' }}
                ariaLabel="Delete dashboard"
                disabled={busy}
                onClick={async () => {
                  if (await ask(`Delete ${shown.title}?`, '', 'Delete'))
                    await run(async () => {
                      await svc.deleteReport(shown.id);
                      setCurrent(undefined);
                      list.reload();
                    });
                }}
              />
            )}
          </>
        }
      />
      <div className="ts-widgets">
        {(shown.definition as IDashboardDefinition).widgets.map((w) => (
          <Widget key={w.id} w={w} />
        ))}
      </div>
      {editing && <DashboardEditor dashboard={editing} onDismiss={() => setEditing(undefined)} onSave={(d) => void save(d)} busy={busy} />}
      {dialog}
    </>
  );
}

function Widget(props: { w: IDashboardWidget }): JSX.Element {
  const { config } = useApp();
  const scope = useScope();
  const r = periodRange(props.w.report.period || 'thisMonth', today(), config.weekStart);
  const data = useFacts({ ...props.w.report.filter, from: r.from, to: r.to });
  return (
    <Card title={props.w.title} className={props.w.width === 3 ? 'is-w3' : props.w.width === 2 ? 'is-w2' : ''}>
      <div className="ts-muted ts-small" style={{ marginTop: -8, marginBottom: 10 }}>
        {formatDay(r.from, true)} – {formatDay(r.to, true)}
      </div>
      {data.loading ? <Loading /> : <ReportView facts={data.facts.filter(scope.filter)} lookups={data.lookups} def={props.w.report} title={props.w.title} compact noExport={props.w.width === 1} />}
    </Card>
  );
}

function DashboardEditor(props: { dashboard: ISavedReport; onDismiss: () => void; onSave: (d: ISavedReport) => void; busy: boolean }): JSX.Element {
  const [d, setD] = React.useState<ISavedReport>(props.dashboard);
  const def = d.definition as IDashboardDefinition;
  const setW = (i: number, patch: Partial<IDashboardWidget>, report?: Partial<IReportDefinition>): void => {
    const widgets = def.widgets.slice();
    widgets[i] = { ...widgets[i], ...patch, report: { ...widgets[i].report, ...(report || {}) } };
    setD({ ...d, definition: { widgets } });
  };
  const move = (i: number, dir: number): void => {
    const widgets = def.widgets.slice();
    const j = i + dir;
    if (j < 0 || j >= widgets.length) return;
    [widgets[i], widgets[j]] = [widgets[j], widgets[i]];
    setD({ ...d, definition: { widgets } });
  };
  return (
    <Panel
      isOpen
      type={PanelType.large}
      headerText="Edit dashboard"
      onDismiss={props.onDismiss}
      isFooterAtBottom
      onRenderFooterContent={() => (
        <div className="ts-row">
          <PrimaryButton text="Save dashboard" onClick={() => props.onSave({ ...d, id: d.owner && d.id ? d.id : 0 })} disabled={props.busy || !d.title.trim()} />
          <DefaultButton text="Cancel" onClick={props.onDismiss} />
        </div>
      )}
    >
      <div className="ts-stack" style={{ paddingTop: 8 }}>
        <div className="ts-form-grid">
          <TextField label="Name" value={d.title} onChange={(_, v) => setD({ ...d, title: v || '' })} />
          <Toggle label="Share with everyone who sees reports" checked={d.shared} onChange={(_, v) => setD({ ...d, shared: !!v })} />
        </div>
        {def.widgets.map((w, i) => (
          <div key={w.id} className="ts-card" style={{ padding: 14 }}>
            <div className="ts-form-grid">
              <TextField label="Title" value={w.title} onChange={(_, v) => setW(i, { title: v || '' })} />
              <Dropdown label="Period" selectedKey={w.report.period || 'thisMonth'} options={PERIODS} onChange={(_, o) => o && setW(i, {}, { period: o.key as IReportDefinition['period'] })} />
              <Dropdown label="Group by" selectedKey={w.report.groupBy[0]} options={DIMENSIONS.map((x) => ({ key: x.key, text: x.text }))} onChange={(_, o) => o && setW(i, {}, { groupBy: [String(o.key)] })} />
              <Dropdown
                label="Measures"
                multiSelect
                selectedKeys={w.report.measures}
                options={MEASURES.map((m) => ({ key: m.key, text: m.text }))}
                onChange={(_, o) => o && setW(i, {}, { measures: o.selected ? w.report.measures.concat(o.key as Measure) : w.report.measures.filter((x) => x !== o.key) })}
              />
              <Dropdown
                label="Chart"
                selectedKey={w.report.chart}
                options={['table', 'bar', 'column', 'line', 'donut'].map((k) => ({ key: k, text: k[0].toUpperCase() + k.substring(1) }))}
                onChange={(_, o) => o && setW(i, {}, { chart: o.key as IReportDefinition['chart'] })}
              />
              <Dropdown
                label="Width"
                selectedKey={w.width}
                options={[
                  { key: 1, text: 'One column' },
                  { key: 2, text: 'Two columns' },
                  { key: 3, text: 'Full width' }
                ]}
                onChange={(_, o) => o && setW(i, { width: Number(o.key) as 1 | 2 | 3 })}
              />
            </div>
            <div className="ts-row" style={{ marginTop: 8 }}>
              <DefaultButton iconProps={{ iconName: 'Up' }} ariaLabel="Move up" onClick={() => move(i, -1)} styles={{ root: { minWidth: 32, padding: 0 } }} />
              <DefaultButton iconProps={{ iconName: 'Down' }} ariaLabel="Move down" onClick={() => move(i, 1)} styles={{ root: { minWidth: 32, padding: 0 } }} />
              <DefaultButton text="Remove" iconProps={{ iconName: 'Delete' }} onClick={() => setD({ ...d, definition: { widgets: def.widgets.filter((_, j) => j !== i) } })} />
            </div>
          </div>
        ))}
        {!def.widgets.length && <Empty title="No widgets" />}
        <DefaultButton iconProps={{ iconName: 'Add' }} text="Add a widget" onClick={() => setD({ ...d, definition: { widgets: def.widgets.concat(widget('New widget', ['project'], ['hours'], 'bar', 'thisMonth', 1)) } })} />
      </div>
    </Panel>
  );
}
