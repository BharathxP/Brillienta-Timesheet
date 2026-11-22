import * as React from 'react';
import { Checkbox, DefaultButton, Dropdown, Panel, PanelType, PrimaryButton, SearchBox } from '@fluentui/react';
import { useApp } from '../AppContext';
import { ITimeEntry, ITimesheet } from '../../models/types';
import { canApprove, sameEmail } from '../../services/access';
import { addDays, formatWeek, isoWeek, today, weekDays, weekStartOf } from '../../services/dates';
import { hours, relativeTime } from '../../services/format';
import { validateWeek } from '../../services/rules';
import { Card, Empty, ErrorBox, Kpi, Loading, PageHeader, Person, StatusBadge, Tabs, useConfirm, useLoader } from '../common/Ui';
import { SheetGrid } from '../timesheet/SheetGrid';
import { Comments } from '../timesheet/Comments';

interface IItem {
  sheet: ITimesheet;
  entries: ITimeEntry[];
}

/** Timesheets waiting for my approval (manager / project manager / programme manager / admin), single or in bulk. */
export function Approvals(): JSX.Element {
  const { svc, me, role, ref, config, workflow, toast, route, refreshApprovals, isAdmin } = useApp();
  const [tab, setTab] = React.useState<'waiting' | 'decided'>('waiting');
  const [q, setQ] = React.useState('');
  const [project, setProject] = React.useState(0);
  const [selected, setSelected] = React.useState<number[]>([]);
  const [open, setOpen] = React.useState<IItem | undefined>();
  const [busy, setBusy] = React.useState(false);
  const [ask, dialog] = useConfirm();
  const { data, error, loading, reload } = useLoader(async (): Promise<IItem[]> => {
    const sheets =
      tab === 'waiting'
        ? await svc.sheetsByStatus('Submitted')
        : (await svc.weeks(addDays(weekStartOf(today(), config.weekStart), -56), weekStartOf(today(), config.weekStart))).sheets.filter((s) => s.status === 'Approved' || s.status === 'Rejected');
    const mine = sheets.filter((s) => !sameEmail(s.userEmail, me.email));
    const entries = mine.length ? await svc.entriesOfSheets(mine.map((s) => s.id)) : [];
    const items = mine.map((s) => ({ sheet: s, entries: entries.filter((e) => e.sheetId === s.id) }));
    // "Decided" shows what I could approve; canApprove checks status-independent rights with a Submitted copy.
    return items
      .filter((i) => canApprove(role, me.email, { ...i.sheet, status: 'Submitted' }, i.entries, ref, config))
      .sort((a, b) => (tab === 'waiting' ? a.sheet.submittedAt.localeCompare(b.sheet.submittedAt) : b.sheet.decidedAt.localeCompare(a.sheet.decidedAt)));
  }, [tab]);

  React.useEffect(() => {
    if (data && route.query.sheet) {
      const hit = data.filter((i) => i.sheet.id === Number(route.query.sheet))[0];
      if (hit) setOpen(hit);
    }
  }, [data]);

  const items = (data || []).filter((i) => {
    if (project && !i.entries.some((e) => e.projectId === project && e.total > 0)) return false;
    const n = q.trim().toLowerCase();
    return !n || i.sheet.userName.toLowerCase().indexOf(n) >= 0 || i.sheet.userEmail.toLowerCase().indexOf(n) >= 0;
  });
  const actor = { email: me.email, name: me.title, role };

  const decide = async (list: IItem[], approve: boolean): Promise<void> => {
    let comment = '';
    if (!approve) {
      const text = await ask(`Send back ${list.length === 1 ? 'this timesheet' : list.length + ' timesheets'}?`, 'Say what needs to change. The person is told and can correct and resubmit.', 'Send back', 'What needs to change', true);
      if (!text) return;
      comment = text;
    } else if (list.length > 1) {
      const ok = await ask(`Approve ${list.length} timesheets?`, `${hours(list.reduce((s, i) => s + i.sheet.totalHours, 0))} hours in total.`, 'Approve all');
      if (!ok) return;
    }
    setBusy(true);
    try {
      const n = await workflow.decide(actor, list, approve, comment);
      toast(`${n} timesheet${n === 1 ? '' : 's'} ${approve ? 'approved' : 'sent back'}.`, 'success');
      setSelected([]);
      setOpen(undefined);
      refreshApprovals();
      reload();
    } catch (e) {
      toast((e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };

  const reopen = async (i: IItem): Promise<void> => {
    const why = await ask('Reopen this timesheet?', 'It goes back to draft so the person can change it and submit again.', 'Reopen', 'Reason (shown to the person)', true);
    if (!why) return;
    setBusy(true);
    try {
      await workflow.reopen(actor, i.sheet, why, i.entries);
      toast('Reopened.', 'success');
      setOpen(undefined);
      reload();
    } catch (e) {
      toast((e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };

  const projectsOnSheets = ref.projects.filter((p) => (data || []).some((i) => i.entries.some((e) => e.projectId === p.id)));
  const total = items.reduce((s, i) => s + i.sheet.totalHours, 0);
  const billable = items.reduce((s, i) => s + i.sheet.billableHours, 0);
  const chosen = items.filter((i) => selected.indexOf(i.sheet.id) >= 0);

  return (
    <>
      <PageHeader title="Approvals" sub={`Timesheets you approve as ${config.approvalBy === 'manager' ? 'their manager' : config.approvalBy === 'project' ? 'project manager' : 'their manager or project manager'}${isAdmin ? ' (administrators approve any)' : ''}.`} />
      <Tabs
        value={tab}
        onChange={(t) => {
          setTab(t);
          setSelected([]);
        }}
        tabs={[
          { key: 'waiting', text: 'Waiting for me', icon: 'DocumentApproval', count: tab === 'waiting' && data ? data.length : undefined },
          { key: 'decided', text: 'Decided (8 weeks)', icon: 'History' }
        ]}
      />
      {error && <ErrorBox error={error} onRetry={reload} />}
      {loading || !data ? (
        <Loading />
      ) : (
        <>
          <div className="ts-kpis">
            <Kpi label={tab === 'waiting' ? 'Waiting' : 'Decided'} value={items.length} icon="DocumentApproval" tint={2} />
            <Kpi label="Hours" value={hours(total)} icon="Clock" tint={1} />
            <Kpi label="Billable" value={hours(billable)} icon="Money" tint={5} />
            <Kpi label="People" value={items.map((i) => i.sheet.userEmail.toLowerCase()).filter((v, k, a) => a.indexOf(v) === k).length} icon="People" tint={4} />
          </div>
          <div className="ts-filters">
            <SearchBox placeholder="Search people" value={q} onChange={(_, v) => setQ(v || '')} styles={{ root: { width: 240 } }} />
            <Dropdown placeholder="All projects" selectedKey={project} options={[{ key: 0, text: 'All projects' }].concat(projectsOnSheets.map((p) => ({ key: p.id, text: p.title })))} onChange={(_, o) => o && setProject(Number(o.key))} />
          </div>
          {chosen.length > 0 && tab === 'waiting' && (
            <div className="ts-bulkbar">
              <strong>{chosen.length} selected</strong> · {hours(chosen.reduce((s, i) => s + i.sheet.totalHours, 0))} h
              <span className="ts-spacer" />
              <PrimaryButton text="Approve" iconProps={{ iconName: 'CheckMark' }} onClick={() => void decide(chosen, true)} disabled={busy} styles={{ root: { background: 'var(--ts-lime)', color: 'var(--ts-navy)', border: 0 } }} />
              <DefaultButton text="Send back" iconProps={{ iconName: 'Undo' }} onClick={() => void decide(chosen, false)} disabled={busy} />
              <DefaultButton text="Clear" onClick={() => setSelected([])} />
            </div>
          )}
          {!items.length ? (
            <Card>
              <Empty icon="CheckMark" title={tab === 'waiting' ? 'Nothing waiting for you' : 'Nothing decided in the last eight weeks'}>
                {tab === 'waiting' ? 'Submitted timesheets you approve appear here.' : ''}
              </Empty>
            </Card>
          ) : (
            <div className="ts-table-wrap">
              <table className="ts-table">
                <thead>
                  <tr>
                    {tab === 'waiting' && (
                      <th className="ts-col-check">
                        <Checkbox checked={chosen.length === items.length} onChange={(_, v) => setSelected(v ? items.map((i) => i.sheet.id) : [])} ariaLabel="Select all" />
                      </th>
                    )}
                    <th>Person</th>
                    <th>Week</th>
                    <th>Projects</th>
                    <th style={{ textAlign: 'right' }}>Hours</th>
                    <th style={{ textAlign: 'right' }}>Billable</th>
                    <th>Checks</th>
                    <th>{tab === 'waiting' ? 'Submitted' : 'Decision'}</th>
                  </tr>
                </thead>
                <tbody>
                  {items.map((i) => {
                    const owner = ref.users.filter((u) => sameEmail(u.email, i.sheet.userEmail))[0];
                    const warnings = validateWeek({ cfg: config, user: owner, days: weekDays(i.sheet.weekStart), entries: i.entries, forSubmit: true });
                    const projects = i.entries
                      .filter((e) => e.total > 0)
                      .map((e) => (ref.projects.filter((p) => p.id === e.projectId)[0] || { title: 'Time off' }).title)
                      .filter((v, k, a) => a.indexOf(v) === k);
                    return (
                      <tr key={i.sheet.id} className={selected.indexOf(i.sheet.id) >= 0 ? 'is-selected' : ''} onClick={() => setOpen(i)}>
                        {tab === 'waiting' && (
                          <td className="ts-col-check" onClick={(e) => e.stopPropagation()}>
                            <Checkbox checked={selected.indexOf(i.sheet.id) >= 0} onChange={(_, v) => setSelected(v ? selected.concat(i.sheet.id) : selected.filter((x) => x !== i.sheet.id))} ariaLabel={`Select ${i.sheet.userName}`} />
                          </td>
                        )}
                        <td>
                          <Person name={i.sheet.userName || i.sheet.userEmail} email={i.sheet.userEmail} />
                        </td>
                        <td>
                          W{isoWeek(i.sheet.weekStart)} <span className="ts-muted">{formatWeek(i.sheet.weekStart)}</span>
                        </td>
                        <td className="ts-small">{projects.slice(0, 3).join(', ') + (projects.length > 3 ? ` +${projects.length - 3}` : '')}</td>
                        <td className="ts-number" style={{ textAlign: 'right' }}>{hours(i.sheet.totalHours)}</td>
                        <td className="ts-number" style={{ textAlign: 'right' }}>{hours(i.sheet.billableHours)}</td>
                        <td>{warnings.length ? <span className="ts-status is-submitted" title={warnings.map((w) => w.message).join('\n')}>{warnings.length} to check</span> : <span className="ts-status is-approved">OK</span>}</td>
                        <td className="ts-small">{tab === 'waiting' ? relativeTime(i.sheet.submittedAt) : <StatusBadge status={i.sheet.status} />}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
      {open && (
        <Panel
          isOpen
          type={PanelType.extraLarge}
          headerText={`${open.sheet.userName} · ${formatWeek(open.sheet.weekStart)}`}
          onDismiss={() => setOpen(undefined)}
          isFooterAtBottom
          onRenderFooterContent={() => (
            <div className="ts-row">
              {open.sheet.status === 'Submitted' ? (
                <>
                  <PrimaryButton text="Approve" iconProps={{ iconName: 'CheckMark' }} onClick={() => void decide([open], true)} disabled={busy} />
                  <DefaultButton text="Send back" iconProps={{ iconName: 'Undo' }} onClick={() => void decide([open], false)} disabled={busy} />
                </>
              ) : (
                <DefaultButton text="Reopen" iconProps={{ iconName: 'Refresh' }} onClick={() => void reopen(open)} disabled={busy} />
              )}
              <DefaultButton text="Close" onClick={() => setOpen(undefined)} />
            </div>
          )}
        >
          <div className="ts-stack" style={{ paddingTop: 10 }}>
            <div className="ts-row">
              <StatusBadge status={open.sheet.status} />
              <span className="ts-muted">
                {hours(open.sheet.totalHours)} h · {hours(open.sheet.billableHours)} billable
              </span>
            </div>
            {(() => {
              const owner = ref.users.filter((u) => sameEmail(u.email, open.sheet.userEmail))[0];
              const w = validateWeek({ cfg: config, user: owner, days: weekDays(open.sheet.weekStart), entries: open.entries, forSubmit: true });
              return w.length ? (
                <ul className="ts-issues">
                  {w.map((x, k) => (
                    <li key={k}>{x.message}</li>
                  ))}
                </ul>
              ) : null;
            })()}
            <SheetGrid days={weekDays(open.sheet.weekStart)} lines={open.entries} owner={{ email: open.sheet.userEmail, user: ref.users.filter((u) => sameEmail(u.email, open.sheet.userEmail))[0] }} readOnly timeOff={[]} />
            {open.sheet.comments.length > 0 && (
              <Card title="History" icon="History">
                <Comments comments={open.sheet.comments} />
              </Card>
            )}
          </div>
        </Panel>
      )}
      {dialog}
    </>
  );
}
