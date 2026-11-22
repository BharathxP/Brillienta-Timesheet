import * as React from 'react';
import { Dimension, IReportDefinition, Measure } from '../../models/types';
import { DIMENSIONS, IFact, ILookups, MEASURES, groupFacts, totals } from '../../services/reports';
import { hours, money } from '../../services/format';
import { useApp } from '../AppContext';
import { BarList, ColumnChart, Donut, LineChart, PALETTE } from '../common/Charts';
import { Empty } from '../common/Ui';
import { ExportButton } from '../common/Pickers';

export function dimLabel(d: Dimension, customLabels: Record<string, string>): string {
  if (String(d).indexOf('custom:') === 0) return customLabels[String(d).substring(7)] || String(d).substring(7);
  return (DIMENSIONS.filter((x) => x.key === d)[0] || { text: d }).text;
}

export function useMeasureFormat(): (m: Measure, v: number) => string {
  const { config } = useApp();
  return (m, v) => (MEASURES.filter((x) => x.key === m)[0] || { money: false }).money ? money(v, config.currency) : m === 'entries' ? String(v) : hours(v);
}

/** A report: its chart (first measure by the first grouping) and its table (every grouping and measure, with totals). */
export function ReportView(props: { facts: IFact[]; def: Pick<IReportDefinition, 'groupBy' | 'measures' | 'chart'>; lookups: ILookups; title: string; compact?: boolean; noExport?: boolean }): JSX.Element {
  const { config, has } = useApp();
  const fmt = useMeasureFormat();
  const customLabels: Record<string, string> = {};
  config.customColumns.forEach((c) => (customLabels[c.key] = c.label));
  const money = has('projectCost') || has('userCost');
  const measures = props.def.measures.filter((m) => money || ['cost', 'revenue', 'margin'].indexOf(m) < 0);
  const groupBy = props.def.groupBy.length ? props.def.groupBy : ['project'];
  const rows = groupFacts(props.facts, groupBy, props.lookups);
  const total = totals(props.facts);
  if (!props.facts.length) return <Empty icon="BarChartVertical" title="No time logged for this selection" />;
  const m0 = measures[0] || 'hours';
  const byFirst = groupFacts(props.facts, [groupBy[0]], props.lookups);
  let chart: JSX.Element | null = null;
  const items = byFirst.slice(0, props.compact ? 8 : 14).map((r, i) => ({ label: r.labels[0], value: r.values[m0], color: PALETTE[i % PALETTE.length] }));
  switch (props.def.chart) {
    case 'bar':
      chart = <BarList items={items.map((x) => ({ ...x, color: undefined }))} format={(v) => fmt(m0, v)} />;
      break;
    case 'donut':
      chart = (
        <div className="ts-donut">
          <Donut items={items.slice(0, 8)} center={fmt(m0, total[m0])} centerSub={(MEASURES.filter((x) => x.key === m0)[0] || { text: '' }).text} />
          <div className="ts-legend" style={{ flexDirection: 'column', gap: 4 }}>
            {items.slice(0, 8).map((x) => (
              <span key={x.label}>
                <i className="ts-dot" style={{ background: x.color }} /> {x.label} · {fmt(m0, x.value)}
              </span>
            ))}
          </div>
        </div>
      );
      break;
    case 'column':
    case 'line': {
      // Over time: group by the time dimension when there is one, series = billable / non-billable or the measures.
      const series = measures.slice(0, 3).map((m, i) => ({ label: (MEASURES.filter((x) => x.key === m)[0] || { text: m }).text, color: PALETTE[i], values: byFirst.map((r) => r.values[m]) }));
      chart = props.def.chart === 'line' ? <LineChart labels={byFirst.map((r) => r.labels[0])} series={series} /> : <ColumnChart labels={byFirst.map((r) => r.labels[0])} series={series} />;
      if (series.length > 1)
        chart = (
          <>
            {chart}
            <div className="ts-legend">
              {series.map((s) => (
                <span key={s.label}>
                  <i className="ts-dot" style={{ background: s.color }} /> {s.label}
                </span>
              ))}
            </div>
          </>
        );
      break;
    }
    default:
      chart = null;
  }
  const columns = groupBy.map((d) => dimLabel(d, customLabels)).concat(measures.map((m) => (MEASURES.filter((x) => x.key === m)[0] || { text: m }).text));
  return (
    <div className="ts-stack">
      {chart && <div className="ts-chart">{chart}</div>}
      <div className="ts-table-wrap" style={props.compact ? { maxHeight: 320, overflowY: 'auto' } : undefined}>
        <table className="ts-table">
          <thead>
            <tr>
              {columns.map((c, i) => (
                <th key={c} style={i >= groupBy.length ? { textAlign: 'right' } : undefined}>
                  {c}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={i} style={{ cursor: 'default' }}>
                {r.labels.map((l, j) => (
                  <td key={j} className={j === 0 ? 'ts-subject' : undefined}>
                    {j > 0 && i > 0 && rows[i - 1].keys.slice(0, j + 1).join() === r.keys.slice(0, j + 1).join() ? '' : l}
                  </td>
                ))}
                {measures.map((m) => (
                  <td key={m} className="ts-number" style={{ textAlign: 'right' }}>
                    {fmt(m, r.values[m])}
                  </td>
                ))}
              </tr>
            ))}
            <tr style={{ cursor: 'default', background: 'var(--ts-surface-2)' }}>
              <td colSpan={groupBy.length}>
                <strong>Total</strong>
              </td>
              {measures.map((m) => (
                <td key={m} className="ts-number" style={{ textAlign: 'right', fontWeight: 700, color: 'var(--ts-text)' }}>
                  {fmt(m, total[m])}
                </td>
              ))}
            </tr>
          </tbody>
        </table>
      </div>
      {!props.noExport && (
        <div className="ts-row">
          <ExportButton
            name={props.title}
            sheets={() => [
              { name: 'Summary', columns, rows: rows.map((r) => (r.labels as Array<string | number>).concat(measures.map((m) => r.values[m]))) },
              {
                name: 'Details',
                columns: ['Date', 'Person', 'Project', 'Task', 'Activity', 'Billable', 'Hours', 'Cost', 'Billable amount', 'Status', 'Note'],
                rows: props.facts.map((f) => {
                  const p = props.lookups.projects.filter((x) => x.id === f.projectId)[0];
                  const t = props.lookups.tasks.filter((x) => x.id === f.taskId)[0];
                  const a = props.lookups.activities.filter((x) => x.id === f.activityId)[0];
                  const u = props.lookups.users.filter((x) => x.email.toLowerCase() === f.userEmail.toLowerCase())[0];
                  return [f.day, u ? u.name : f.userEmail, p ? p.title : '', t ? t.title : '', a ? a.title : '', f.billable, f.hours, Math.round(f.cost * 100) / 100, Math.round(f.revenue * 100) / 100, f.status, f.note];
                })
              }
            ]}
          />
        </div>
      )}
    </div>
  );
}
