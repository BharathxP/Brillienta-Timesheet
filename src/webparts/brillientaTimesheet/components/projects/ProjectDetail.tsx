import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, MessageBar, MessageBarType, PrimaryButton, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { Dimension, ITask } from '../../models/types';
import { canManageProject, sameEmail } from '../../services/access';
import { daysBetween, formatDay, today } from '../../services/dates';
import { hours, money } from '../../services/format';
import { totals } from '../../services/reports';
import { IPlannerPlan, IPlannerTask } from '../../services/GraphService';
import { Card, Empty, ErrorBox, Kpi, Loading, Locked, PageHeader, Person, Progress, Segmented, StatusBadge, Tabs, useBusy, useConfirm } from '../common/Ui';
import { ProjectEditor } from './ProjectEditor';
import { TaskEditor, newTask } from './TaskEditor';
import { BulkImport } from './BulkImport';
import { useFacts } from '../reports/useFacts';
import { ReportView } from '../reports/ReportView';

type Tab = 'overview' | 'tasks' | 'report' | 'team';

export function ProjectDetail(props: { id: number }): JSX.Element {
  const { ref, role, me, navigate, svc, reloadRef, has, config, toast, graph } = useApp();
  const project = ref.projects.filter((p) => p.id === props.id)[0];
  const [tab, setTab] = React.useState<Tab>('overview');
  const [editing, setEditing] = React.useState(false);
  const [task, setTask] = React.useState<ITask | undefined>();
  const [bulk, setBulk] = React.useState(false);
  const [planner, setPlanner] = React.useState(false);
  const [groupBy, setGroupBy] = React.useState<Dimension>('user');
  const [busy, run] = useBusy();
  const [ask, dialog] = useConfirm();
  const data = useFacts({}, { projectId: props.id });
  if (!project) return <Empty icon="ProjectCollection" title="This project does not exist (any more)." />;
  const manage = canManageProject(role, me.email, project, ref);
  const tasks = ref.tasks.filter((t) => t.projectId === project.id);
  const tot = totals(data.facts);
  const t0 = today();
  const left = project.endDate ? daysBetween(t0, project.endDate) : undefined;
  const hasMoney = has('projectCost') || has('userCost');
  const taskHours: Record<number, number> = {};
  data.facts.forEach((f) => (taskHours[f.taskId] = (taskHours[f.taskId] || 0) + f.hours));
  const people = data.facts.map((f) => f.userEmail.toLowerCase()).filter((v, i, a) => a.indexOf(v) === i);

  const removeTask = async (t: ITask): Promise<void> => {
    if (!(await ask(`Delete the task ${t.title}?`, 'Time already logged on it stays in reports.', 'Delete'))) return;
    await run(async () => {
      await svc.deleteTask(t.id);
      await reloadRef();
    }, 'Task deleted.');
  };

  return (
    <>
      <PageHeader
        back={() => navigate('projects')}
        eyebrow={[project.code, project.client].filter((x) => !!x).join(' · ') || 'Project'}
        title={
          <>
            <span className="ts-swatch" style={{ background: project.color || '#c8d0d2', width: 14, height: 14 }} />
            {project.title}
          </>
        }
        sub={
          <span className="ts-row">
            <StatusBadge status={project.status} />
            {!project.billable && <span className="ts-chip is-soft">Non-billable</span>}
            {project.startDate && <span>From {formatDay(project.startDate, true)}</span>}
            {project.endDate && <span>to {formatDay(project.endDate, true)}</span>}
          </span>
        }
        actions={manage && <DefaultButton iconProps={{ iconName: 'Edit' }} text="Edit project" onClick={() => setEditing(true)} />}
      />
      {left !== undefined && project.status === 'Active' && left <= 14 && (
        <MessageBar messageBarType={left < 0 ? MessageBarType.severeWarning : MessageBarType.warning} styles={{ root: { marginBottom: 12, borderRadius: 8 } }}>
          {left < 0 ? `The end date passed ${-left} days ago. Complete the project or move its end date.` : `The project ends in ${left} day${left === 1 ? '' : 's'}.`}
        </MessageBar>
      )}
      <Tabs
        value={tab}
        onChange={setTab}
        tabs={[
          { key: 'overview', text: 'Overview', icon: 'ViewDashboard' },
          { key: 'tasks', text: 'Tasks', icon: 'TaskManager', count: tasks.length },
          { key: 'report', text: 'Detailed report', icon: 'ReportDocument', locked: !has('projectReport') },
          { key: 'team', text: 'Team & rates', icon: 'People', count: people.length }
        ]}
      />
      {data.error && <ErrorBox error={data.error} onRetry={data.reload} />}
      {data.loading ? (
        <Loading />
      ) : tab === 'overview' ? (
        <>
          <div className="ts-kpis">
            <Kpi label="Logged hours" value={hours(tot.hours)} icon="Clock" tint={1} />
            {has('plannedHours') && <Kpi label="Planned hours" value={project.plannedHours ? hours(project.plannedHours) : '—'} icon="Bullseye" tint={2} tone={project.plannedHours && tot.hours > project.plannedHours ? 'danger' : undefined} />}
            <Kpi label="Billable" value={hours(tot.billableHours)} icon="Money" tint={5} hint={tot.hours ? `${Math.round((tot.billableHours / tot.hours) * 100)}% of hours` : undefined} />
            {hasMoney && <Kpi label="Cost" value={money(tot.cost, config.currency)} icon="Calculator" tint={3} />}
            {hasMoney && <Kpi label="Billable amount" value={money(tot.revenue, config.currency)} icon="Financial" tint={4} />}
            {hasMoney && <Kpi label="Margin" value={money(tot.margin, config.currency)} icon="Chart" tint={1} tone={tot.margin < 0 ? 'danger' : 'success'} />}
          </div>
          <div className="ts-grid ts-grid-2">
            <Card title="Planned hours" icon="Bullseye">
              {!has('plannedHours') ? (
                <Locked feature="plannedHours" />
              ) : project.plannedHours ? (
                <>
                  <div style={{ fontSize: 22, fontWeight: 700, fontFamily: 'var(--ts-heading)' }}>
                    {hours(tot.hours)} / {hours(project.plannedHours)} h
                  </div>
                  <Progress value={tot.hours} max={project.plannedHours} />
                  <div className="ts-muted ts-small" style={{ marginTop: 6 }}>
                    {tot.hours > project.plannedHours ? `${hours(tot.hours - project.plannedHours)} h over plan` : `${hours(project.plannedHours - tot.hours)} h left`}
                  </div>
                </>
              ) : (
                <div className="ts-muted">No planned hours set.</div>
              )}
            </Card>
            <Card title="Hours by person" icon="People">
              <ReportView facts={data.facts} def={{ groupBy: ['user'], measures: ['hours'], chart: 'bar' }} lookups={data.lookups} title={project.title} compact noExport />
            </Card>
            <Card title="People" icon="Contact">
              <div className="ts-stack" style={{ gap: 8 }}>
                <div>
                  <div className="ts-pill">Project managers</div>
                  {project.managerEmails.map((e) => (
                    <div key={e}>
                      <Person name={(ref.users.filter((u) => sameEmail(u.email, e))[0] || { name: e }).name} email={e} />
                    </div>
                  ))}
                </div>
                <div>
                  <div className="ts-pill">Members</div>
                  <div className="ts-small">{project.memberEmails.length ? `${project.memberEmails.length} people` : 'Open to everyone'}</div>
                </div>
              </div>
            </Card>
            <Card title="Details" icon="Info">
              <div className="ts-small">{project.description || <span className="ts-muted">No description.</span>}</div>
              {config.customColumns
                .filter((c) => c.level === 'project')
                .map((c) => (
                  <div key={c.key} className="ts-small">
                    <strong>{c.label}:</strong> {String(project.custom[c.key] === undefined ? '—' : project.custom[c.key])}
                  </div>
                ))}
              {hasMoney && (
                <div className="ts-small" style={{ marginTop: 8 }}>
                  Flat hourly cost {project.hourlyCost ? money(project.hourlyCost, config.currency) : '—'} · bill rate {project.billRate ? money(project.billRate, config.currency) : '—'}
                </div>
              )}
            </Card>
          </div>
        </>
      ) : tab === 'tasks' ? (
        <>
          {manage && (
            <div className="ts-row" style={{ marginBottom: 12 }}>
              <PrimaryButton iconProps={{ iconName: 'Add' }} text="New task" onClick={() => setTask(newTask(project.id))} />
              {has('bulkCreate') && <DefaultButton iconProps={{ iconName: 'BulkUpload' }} text="Bulk create tasks" onClick={() => setBulk(true)} />}
              {has('planner') && config.integrations.plannerEnabled && graph && <DefaultButton iconProps={{ iconName: 'PlannerLogo' }} text="Import from Planner" onClick={() => setPlanner(true)} />}
            </div>
          )}
          {!tasks.length ? (
            <Card>
              <Empty icon="TaskManager" title="No tasks yet">
                Time can be logged on the project without tasks.
              </Empty>
            </Card>
          ) : (
            <div className="ts-table-wrap">
              <table className="ts-table">
                <thead>
                  <tr>
                    <th>Task</th>
                    <th>Assignees</th>
                    <th>Status</th>
                    <th>Due</th>
                    <th style={{ width: 170 }}>Logged / planned</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {tasks.map((t) => (
                    <tr key={t.id} onClick={() => manage && setTask(t)} style={manage ? undefined : { cursor: 'default' }}>
                      <td>
                        <div className="ts-subject">{t.title}</div>
                        <div className="ts-row ts-small ts-muted">
                          {!t.billable && <span>Non-billable</span>}
                          {t.selfAssign && has('selfAssign') && <span>Self-assign</span>}
                          {t.plannerTaskId && <span>Planner</span>}
                        </div>
                      </td>
                      <td className="ts-small">{t.assigneeEmails.length ? t.assigneeEmails.map((e) => (ref.users.filter((u) => sameEmail(u.email, e))[0] || { name: e }).name).join(', ') : <span className="ts-muted">Anyone on the project</span>}</td>
                      <td>
                        <StatusBadge status={t.status} />
                      </td>
                      <td className="ts-muted">{t.dueDate ? formatDay(t.dueDate, true) : '—'}</td>
                      <td>
                        <div className="ts-small">
                          {hours(taskHours[t.id] || 0)} h{t.plannedHours ? ` / ${hours(t.plannedHours)} h` : ''}
                        </div>
                        {t.plannedHours > 0 && <Progress value={taskHours[t.id] || 0} max={t.plannedHours} />}
                      </td>
                      <td style={{ textAlign: 'right' }} onClick={(e) => e.stopPropagation()}>
                        {manage && <DefaultButton iconProps={{ iconName: 'Delete' }} ariaLabel="Delete task" onClick={() => void removeTask(t)} disabled={busy} styles={{ root: { minWidth: 32, padding: 0 } }} />}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      ) : tab === 'report' ? (
        !has('projectReport') ? (
          <Locked feature="projectReport" />
        ) : (
          <Card
            title="Detailed project report"
            icon="ReportDocument"
            actions={
              <Segmented
                value={groupBy}
                onChange={setGroupBy}
                options={[
                  { key: 'user', text: 'Person' },
                  { key: 'task', text: 'Task' },
                  { key: 'activity', text: 'Activity' },
                  { key: 'week', text: 'Week' },
                  { key: 'month', text: 'Month' },
                  { key: 'billable', text: 'Billable' }
                ]}
              />
            }
          >
            <ReportView
              facts={data.facts}
              def={{ groupBy: [groupBy, groupBy === 'user' ? 'task' : 'user'], measures: hasMoney ? ['hours', 'billableHours', 'nonBillableHours', 'cost', 'revenue'] : ['hours', 'billableHours', 'nonBillableHours'], chart: groupBy === 'week' || groupBy === 'month' ? 'column' : 'bar' }}
              lookups={data.lookups}
              title={`${project.title} report`}
            />
          </Card>
        )
      ) : (
        <Card title="Team & rates" icon="People">
          <ReportView facts={data.facts} def={{ groupBy: ['user'], measures: hasMoney ? ['hours', 'billableHours', 'cost', 'revenue', 'margin'] : ['hours', 'billableHours'], chart: 'table' }} lookups={data.lookups} title={`${project.title} team`} />
        </Card>
      )}
      {editing && (
        <ProjectEditor
          project={project}
          onDismiss={() => setEditing(false)}
          onSaved={() => {
            setEditing(false);
            void reloadRef().then(() => toast('Project saved.', 'success'));
          }}
        />
      )}
      {task && (
        <TaskEditor
          task={task}
          projectChoices={[project.id]}
          onDismiss={() => setTask(undefined)}
          onSaved={() => {
            setTask(undefined);
            void reloadRef();
          }}
        />
      )}
      {bulk && (
        <BulkImport
          kind="tasks"
          onDismiss={() => setBulk(false)}
          onDone={() => {
            setBulk(false);
            void reloadRef();
          }}
        />
      )}
      {planner && (
        <PlannerImport
          projectId={project.id}
          onDismiss={() => setPlanner(false)}
          onDone={() => {
            setPlanner(false);
            void reloadRef();
          }}
        />
      )}
      {dialog}
    </>
  );
}

/** Planner integration: import a plan's tasks (with assignees and due dates) into the project; re-import updates them. */
function PlannerImport(props: { projectId: number; onDismiss: () => void; onDone: () => void }): JSX.Element {
  const { graph, svc, ref, toast } = useApp();
  const [plans, setPlans] = React.useState<IPlannerPlan[] | undefined>();
  const [plan, setPlan] = React.useState('');
  const [tasks, setTasks] = React.useState<IPlannerTask[] | undefined>();
  const [skipDone, setSkipDone] = React.useState(true);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState('');
  React.useEffect(() => {
    if (!graph) return;
    graph
      .plans()
      .then(setPlans)
      .catch((e) => setError('Planner could not be read: ' + (e as Error).message + ' (an admin must approve the Tasks.ReadWrite and Group.Read.All permissions in the SharePoint admin center › API access).'));
  }, []);
  React.useEffect(() => {
    if (!graph || !plan) return;
    setTasks(undefined);
    graph
      .planTasks(plan)
      .then(setTasks)
      .catch((e) => setError((e as Error).message));
  }, [plan]);
  const chosen = (tasks || []).filter((t) => !skipDone || t.percentComplete < 100);
  const run = async (): Promise<void> => {
    if (!graph) return;
    setBusy(true);
    try {
      const emails = await graph.emailsOf(chosen.reduce<string[]>((a, t) => a.concat(t.assigneeIds), []));
      const existing = ref.tasks.filter((t) => t.projectId === props.projectId && t.plannerTaskId);
      const fresh = chosen.filter((t) => !existing.some((x) => x.plannerTaskId === t.id));
      const status = (p: number): ITask['status'] => (p >= 100 ? 'Done' : p > 0 ? 'In progress' : 'Open');
      await svc.addTasks(
        fresh.map((t) => ({
          title: t.title,
          projectId: props.projectId,
          assigneeEmails: t.assigneeIds.map((id) => emails[id]).filter((e) => !!e),
          dueDate: t.dueDate,
          status: status(t.percentComplete),
          billable: true,
          selfAssign: true,
          plannedHours: 0,
          description: '',
          plannerTaskId: t.id,
          plannerPlanId: plan
        }))
      );
      for (const t of chosen.filter((x) => existing.some((e) => e.plannerTaskId === x.id))) {
        const mine = existing.filter((e) => e.plannerTaskId === t.id)[0];
        await svc.saveTask({ id: mine.id, title: t.title, dueDate: t.dueDate, status: status(t.percentComplete), assigneeEmails: t.assigneeIds.map((id) => emails[id]).filter((e) => !!e) });
      }
      await svc.log('Planner tasks imported', `${fresh.length} new, ${chosen.length - fresh.length} updated`);
      toast(`${fresh.length} tasks imported, ${chosen.length - fresh.length} updated from Planner.`, 'success');
      props.onDone();
    } catch (e) {
      toast((e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };
  return (
    <Dialog hidden={false} onDismiss={props.onDismiss} dialogContentProps={{ type: DialogType.normal, title: 'Import tasks from Microsoft Planner', subText: 'Tasks are linked: importing again updates their names, dates, assignees and progress.' }} minWidth={520}>
      {error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>}
      {!plans && !error ? (
        <Loading label="Reading your plans…" />
      ) : (
        <Dropdown label="Plan" placeholder="Choose a plan" options={(plans || []).map((p) => ({ key: p.id, text: `${p.title} (${p.owner})` }))} selectedKey={plan || null} onChange={(_, o) => o && setPlan(String(o.key))} />
      )}
      <Toggle label="Skip completed tasks" checked={skipDone} onChange={(_, v) => setSkipDone(!!v)} styles={{ root: { marginTop: 10 } }} />
      {plan && !tasks && <Loading label="Reading tasks…" />}
      {tasks && <div className="ts-muted">{chosen.length} tasks to import.</div>}
      <DialogFooter>
        <PrimaryButton text="Import" onClick={() => void run()} disabled={busy || !chosen.length} />
        <DefaultButton text="Cancel" onClick={props.onDismiss} />
      </DialogFooter>
    </Dialog>
  );
}
