import * as React from 'react';
import { DefaultButton, Dropdown, PrimaryButton, SearchBox } from '@fluentui/react';
import { useApp } from '../AppContext';
import { IProject } from '../../models/types';
import { canCreateProjects, canManageProject, managedProjects } from '../../services/access';
import { daysBetween, formatDay, today } from '../../services/dates';
import { hours } from '../../services/format';
import { PALETTE } from '../../schema/defaults';
import { Card, Empty, Kpi, PageHeader, Progress, StatusBadge, useBusy, useConfirm, useLoader } from '../common/Ui';
import { ExportButton } from '../common/Pickers';
import { ProjectEditor, newProject } from './ProjectEditor';
import { BulkImport } from './BulkImport';
import { BulkKind } from '../../services/bulk';

export function Projects(): JSX.Element {
  const { ref, role, me, navigate, svc, reloadRef, has, toast, isAdmin } = useApp();
  const [q, setQ] = React.useState('');
  const [status, setStatus] = React.useState('Active');
  const [client, setClient] = React.useState('');
  const [program, setProgram] = React.useState(0);
  const [editing, setEditing] = React.useState<IProject | undefined>();
  const [bulk, setBulk] = React.useState<BulkKind | undefined>();
  const [busy, run] = useBusy();
  const [ask, dialog] = useConfirm();
  const used = useLoader(() => svc.hoursByProject(), [ref.projects.length]);
  const visible = isAdmin ? ref.projects : managedProjects(role, me.email, ref).concat(ref.projects.filter((p) => !p.memberEmails.length || p.memberEmails.some((e) => e.toLowerCase() === me.email.toLowerCase()))).filter((p, i, a) => a.indexOf(p) === i);
  const clients = visible.map((p) => p.client).filter((c, i, a) => !!c && a.indexOf(c) === i).sort();
  const n = q.trim().toLowerCase();
  const list = visible.filter(
    (p) =>
      (status === 'all' || p.status === status) &&
      (!client || p.client === client) &&
      (!program || p.programId === program) &&
      (!n || p.title.toLowerCase().indexOf(n) >= 0 || p.code.toLowerCase().indexOf(n) >= 0 || p.client.toLowerCase().indexOf(n) >= 0)
  );
  const t = today();
  const u = used.data || {};
  const ending = visible.filter((p) => p.status === 'Active' && p.endDate && daysBetween(t, p.endDate) <= 14 && daysBetween(t, p.endDate) >= 0).length;
  const overPlan = visible.filter((p) => p.plannedHours > 0 && (u[p.id] || 0) > p.plannedHours).length;

  const setProjectStatus = (p: IProject, s: IProject['status']): Promise<void> =>
    run(async () => {
      await svc.saveProject({ id: p.id, status: s });
      await svc.log(`Project ${s.toLowerCase()}`, p.title);
      await reloadRef();
    }, `${p.title}: ${s}.`);
  const remove = async (p: IProject): Promise<void> => {
    const tasks = ref.tasks.filter((x) => x.projectId === p.id);
    const ok = await ask(`Delete ${p.title}?`, `The project and its ${tasks.length} task${tasks.length === 1 ? '' : 's'} are deleted. Time already logged stays in timesheets and reports (as a deleted project). Disable it instead to keep it.`, 'Delete');
    if (!ok) return;
    await run(async () => {
      await svc.deleteProject(
        p.id,
        tasks.map((x) => x.id)
      );
      await svc.log('Project deleted', p.title);
      await reloadRef();
    }, 'Project deleted.');
  };

  return (
    <>
      <PageHeader
        title="Projects"
        sub="Everything time is logged on: clients, managers, members, dates, planned hours and rates."
        actions={
          <>
            <ExportButton
              name="projects"
              sheets={() => [
                {
                  name: 'Projects',
                  columns: ['Code', 'Project', 'Client', 'Status', 'Managers', 'Start', 'End', 'Planned hours', 'Logged hours', 'Hourly cost', 'Bill rate', 'Billable'],
                  rows: list.map((p) => [p.code, p.title, p.client, p.status, p.managerEmails.join('; '), p.startDate, p.endDate, p.plannedHours, u[p.id] || 0, p.hourlyCost, p.billRate, p.billable])
                }
              ]}
            />
            {canCreateProjects(role) && has('bulkCreate') && (
              <DefaultButton
                text="Bulk create"
                iconProps={{ iconName: 'BulkUpload' }}
                menuProps={{
                  items: [
                    { key: 'projects', text: 'Projects', onClick: () => setBulk('projects') },
                    { key: 'tasks', text: 'Tasks', onClick: () => setBulk('tasks') }
                  ]
                }}
              />
            )}
            {canCreateProjects(role) && <PrimaryButton iconProps={{ iconName: 'Add' }} text="New project" onClick={() => setEditing(newProject(me.email, PALETTE[ref.projects.length % PALETTE.length]))} />}
          </>
        }
      />
      <div className="ts-kpis">
        <Kpi label="Active projects" value={visible.filter((p) => p.status === 'Active').length} icon="ProjectCollection" tint={1} onClick={() => setStatus('Active')} />
        <Kpi label="Ending in 14 days" value={ending} icon="Calendar" tint={3} tone={ending ? 'warning' : undefined} />
        <Kpi label="Over planned hours" value={overPlan} icon="Warning" tint={4} tone={overPlan ? 'danger' : undefined} />
        <Kpi label="Clients" value={clients.length} icon="Contact" tint={2} />
      </div>
      <div className="ts-filters">
        <SearchBox placeholder="Search projects, codes, clients" value={q} onChange={(_, v) => setQ(v || '')} styles={{ root: { width: 260 } }} />
        <Dropdown selectedKey={status} onChange={(_, o) => o && setStatus(String(o.key))} options={['Active', 'Completed', 'Disabled', 'all'].map((k) => ({ key: k, text: k === 'all' ? 'All statuses' : k }))} />
        <Dropdown selectedKey={client} onChange={(_, o) => o && setClient(String(o.key))} options={[{ key: '', text: 'All clients' }].concat(clients.map((c) => ({ key: c, text: c })))} />
        {has('programManager') && ref.programs.length > 0 && <Dropdown selectedKey={program} onChange={(_, o) => o && setProgram(Number(o.key))} options={[{ key: 0, text: 'All programmes' }].concat(ref.programs.map((g) => ({ key: g.id, text: g.title })))} />}
      </div>
      {!list.length ? (
        <Card>
          <Empty icon="ProjectCollection" title="No projects here">
            {canCreateProjects(role) ? 'Create one, or paste many at once with Bulk create.' : 'Projects you manage or belong to appear here.'}
          </Empty>
        </Card>
      ) : (
        <div className="ts-table-wrap">
          <table className="ts-table">
            <thead>
              <tr>
                <th>Project</th>
                <th>Client</th>
                <th>Status</th>
                <th>End date</th>
                <th style={{ width: 190 }}>Logged / planned</th>
                <th>Managers</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {list.map((p) => {
                const manage = canManageProject(role, me.email, p, ref);
                const left = p.endDate ? daysBetween(t, p.endDate) : undefined;
                return (
                  <tr key={p.id} onClick={() => navigate(`projects/${p.id}`)}>
                    <td>
                      <span className="ts-swatch" style={{ background: p.color || '#c8d0d2' }} />
                      <span className="ts-subject">{p.title}</span>
                      {p.code && <span className="ts-muted ts-small"> · {p.code}</span>}
                      {!p.billable && <span className="ts-chip is-soft" style={{ marginLeft: 6 }}>Non-billable</span>}
                      {p.tmProjectId > 0 && <span className="ts-chip is-lime" style={{ marginLeft: 6 }} title="Linked to Brillienta Task Manager">Task Manager</span>}
                    </td>
                    <td>{p.client || <span className="ts-muted">—</span>}</td>
                    <td>
                      <StatusBadge status={p.status} />
                    </td>
                    <td className={left !== undefined && left <= 7 && p.status === 'Active' ? 'ts-status-warn' : 'ts-muted'}>
                      {p.endDate ? formatDay(p.endDate, true) : '—'}
                      {left !== undefined && p.status === 'Active' && left >= 0 && left <= 30 && <div className="ts-small">{left} days left</div>}
                    </td>
                    <td>
                      <div className="ts-small">
                        {hours(u[p.id] || 0)} h{p.plannedHours ? ` / ${hours(p.plannedHours)} h` : ''}
                      </div>
                      {p.plannedHours > 0 && <Progress value={u[p.id] || 0} max={p.plannedHours} />}
                    </td>
                    <td className="ts-small">{p.managerEmails.map((e) => (ref.users.filter((x) => x.email.toLowerCase() === e.toLowerCase())[0] || { name: e }).name).join(', ')}</td>
                    <td style={{ textAlign: 'right' }} onClick={(e) => e.stopPropagation()}>
                      {manage && (
                        <DefaultButton
                          text="Manage"
                          disabled={busy}
                          menuProps={{
                            items: [
                              { key: 'edit', text: 'Edit', iconProps: { iconName: 'Edit' }, onClick: () => setEditing(p) },
                              p.status !== 'Active'
                                ? { key: 'act', text: 'Make active', iconProps: { iconName: 'Play' }, onClick: () => void setProjectStatus(p, 'Active') }
                                : { key: 'done', text: 'Mark completed', iconProps: { iconName: 'CheckMark' }, onClick: () => void setProjectStatus(p, 'Completed') },
                              {
                                key: 'dis',
                                text: has('disableProjects') ? 'Disable (no new time)' : 'Disable (Enterprise)',
                                iconProps: { iconName: 'Blocked' },
                                disabled: !has('disableProjects') || p.status === 'Disabled',
                                onClick: () => void setProjectStatus(p, 'Disabled')
                              },
                              {
                                key: 'del',
                                text: has('deleteProjects') ? 'Delete' : 'Delete (Enterprise)',
                                iconProps: { iconName: 'Delete' },
                                disabled: !has('deleteProjects') || (!isAdmin && role !== 'programManager'),
                                onClick: () => void remove(p)
                              }
                            ]
                          }}
                        />
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      {editing && (
        <ProjectEditor
          project={editing}
          onDismiss={() => setEditing(undefined)}
          onSaved={(id) => {
            setEditing(undefined);
            void reloadRef().then(() => toast('Project saved.', 'success'));
            if (!editing.id) navigate(`projects/${id}`);
          }}
        />
      )}
      {bulk && (
        <BulkImport
          kind={bulk}
          onDismiss={() => setBulk(undefined)}
          onDone={() => {
            setBulk(undefined);
            void reloadRef();
          }}
        />
      )}
      {dialog}
    </>
  );
}
