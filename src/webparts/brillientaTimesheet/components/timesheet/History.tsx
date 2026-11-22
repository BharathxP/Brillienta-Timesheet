import * as React from 'react';
import { DefaultButton, Dropdown } from '@fluentui/react';
import { useApp, buildQuery } from '../AppContext';
import { ITimesheet } from '../../models/types';
import { addDays, formatWeek, isoWeek, today, weekStartOf } from '../../services/dates';
import { hours, formatDateTime } from '../../services/format';
import { Empty, ErrorBox, Kpi, Loading, PageHeader, StatusBadge, useLoader } from '../common/Ui';
import { ExportButton } from '../common/Pickers';

/** My timesheets: every week with its status, plus the weeks not filled in yet. */
export function History(): JSX.Element {
  const { svc, me, navigate, config } = useApp();
  const [filter, setFilter] = React.useState('all');
  const { data, error, loading, reload } = useLoader(() => svc.sheetsOf(me.email, 200), [me.email]);
  if (error) return <ErrorBox error={error} onRetry={reload} />;
  if (loading || !data) return <Loading />;
  // Weeks of the last eight with no timesheet at all show as "Not started".
  const thisWeek = weekStartOf(today(), config.weekStart);
  const rows: Array<ITimesheet | { weekStart: string; status: 'Not started' }> = [];
  for (let i = 0; i < 8; i++) {
    const w = addDays(thisWeek, -7 * i);
    if (!data.some((s) => s.weekStart === w)) rows.push({ weekStart: w, status: 'Not started' });
  }
  const all = (rows as Array<ITimesheet | { weekStart: string; status: 'Not started' }>).concat(data).sort((a, b) => b.weekStart.localeCompare(a.weekStart));
  const shown = all.filter((s) => filter === 'all' || s.status === filter);
  const count = (st: string): number => all.filter((s) => s.status === st).length;
  const sheets = data;
  return (
    <>
      <PageHeader
        title="My timesheets"
        sub="Every week you have filled in, with its approval status."
        actions={
          <ExportButton
            name="my-timesheets"
            sheets={() => [{ name: 'Timesheets', columns: ['Week starting', 'Status', 'Total hours', 'Billable hours', 'Submitted', 'Decided by'], rows: sheets.map((s) => [s.weekStart, s.status, s.totalHours, s.billableHours, s.submittedAt, s.decidedBy]) }]}
          />
        }
      />
      <div className="ts-kpis">
        <Kpi label="Drafts" value={count('Draft')} icon="Edit" tint={5} onClick={() => setFilter('Draft')} />
        <Kpi label="Waiting for approval" value={count('Submitted')} icon="Send" tint={2} onClick={() => setFilter('Submitted')} />
        <Kpi label="Sent back" value={count('Rejected')} icon="Undo" tint={3} tone={count('Rejected') ? 'danger' : undefined} onClick={() => setFilter('Rejected')} />
        <Kpi label="Approved" value={count('Approved')} icon="CheckMark" tint={1} onClick={() => setFilter('Approved')} />
        <Kpi label="Not started (8 weeks)" value={count('Not started')} icon="Clock" tint={4} tone={count('Not started') ? 'warning' : undefined} onClick={() => setFilter('Not started')} />
      </div>
      <div className="ts-filters">
        <Dropdown
          label="Status"
          selectedKey={filter}
          onChange={(_, o) => o && setFilter(String(o.key))}
          options={['all', 'Not started', 'Draft', 'Submitted', 'Rejected', 'Approved'].map((k) => ({ key: k, text: k === 'all' ? 'All weeks' : k }))}
        />
      </div>
      {!shown.length ? (
        <Empty title="No weeks here" />
      ) : (
        <div className="ts-table-wrap">
          <table className="ts-table">
            <thead>
              <tr>
                <th>Week</th>
                <th>Status</th>
                <th style={{ textAlign: 'right' }}>Hours</th>
                <th style={{ textAlign: 'right' }}>Billable</th>
                <th>Submitted</th>
                <th>Approved / rejected by</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {shown.map((s) => {
                const sheet = 'id' in s ? s : undefined;
                return (
                  <tr key={s.weekStart} onClick={() => navigate('timesheet' + buildQuery({ week: s.weekStart }))}>
                    <td>
                      <span className="ts-subject">W{isoWeek(s.weekStart)}</span> <span className="ts-muted">{formatWeek(s.weekStart)}</span>
                    </td>
                    <td>
                      <StatusBadge status={s.status} />
                    </td>
                    <td className="ts-number" style={{ textAlign: 'right' }}>{sheet ? hours(sheet.totalHours) : '–'}</td>
                    <td className="ts-number" style={{ textAlign: 'right' }}>{sheet ? hours(sheet.billableHours) : '–'}</td>
                    <td className="ts-muted ts-small">{sheet ? formatDateTime(sheet.submittedAt) : ''}</td>
                    <td className="ts-small">{sheet ? sheet.decidedBy : ''}</td>
                    <td style={{ textAlign: 'right' }}>
                      <DefaultButton text="Open" />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </>
  );
}
