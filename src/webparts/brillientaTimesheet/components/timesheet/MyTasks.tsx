import * as React from 'react';
import { DefaultButton, Dropdown, PrimaryButton } from '@fluentui/react';
import { useApp } from '../AppContext';
import { ITask } from '../../models/types';
import { canManageProject, projectsFor, sameEmail, selfAssignable } from '../../services/access';
import { formatDay, today } from '../../services/dates';
import { hours } from '../../services/format';
import { Card, Empty, PageHeader, Progress, StatusBadge, Tabs, useBusy, useLoader } from '../common/Ui';
import { TaskEditor, newTask } from '../projects/TaskEditor';

/** My tasks: what I am assigned, open tasks I can take (self-assign), and my progress on each. */
export function MyTasks(): JSX.Element {
  const { ref, me, svc, reloadRef, config, has, role, navigate } = useApp();
  const [tab, setTab] = React.useState<'mine' | 'open'>('mine');
  const [editing, setEditing] = React.useState<ITask | undefined>();
  const [busy, run] = useBusy();
  const used = useLoader(() => svc.hoursByTask(), [ref.tasks.length]);
  const mine = ref.tasks.filter((t) => t.assigneeEmails.some((e) => sameEmail(e, me.email)));
  const open = has('selfAssign') && config.selfAssign ? selfAssignable(ref, me.email) : [];
  const myProjects = projectsFor(ref, me.email);
  const canCreate = myProjects.filter((p) => canManageProject(role, me.email, p, ref) || (config.membersCreateTasks && myProjects.indexOf(p) >= 0));
  const list = tab === 'mine' ? mine : open;
  const t0 = today();
  const project = (id: number): string => (ref.projects.filter((p) => p.id === id)[0] || { title: '(removed project)' }).title;

  const take = (t: ITask): Promise<void> =>
    run(async () => {
      await svc.saveTask({ id: t.id, assigneeEmails: t.assigneeEmails.concat(me.email) });
      await svc.log('Task self-assigned', `${t.title} → ${me.email}`);
      await reloadRef();
    }, `“${t.title}” is now yours.`);
  const drop = (t: ITask): Promise<void> =>
    run(async () => {
      await svc.saveTask({ id: t.id, assigneeEmails: t.assigneeEmails.filter((e) => !sameEmail(e, me.email)) });
      await reloadRef();
    }, 'Removed from your tasks.');
  const setStatus = (t: ITask, status: ITask['status']): Promise<void> =>
    run(async () => {
      await svc.saveTask({ id: t.id, status });
      await reloadRef();
    });

  return (
    <>
      <PageHeader
        title="My tasks"
        sub="Tasks assigned to you, and open tasks on your projects you can take on."
        actions={canCreate.length > 0 && <PrimaryButton iconProps={{ iconName: 'Add' }} text="New task" onClick={() => setEditing(newTask(canCreate[0].id))} />}
      />
      <Tabs
        value={tab}
        onChange={setTab}
        tabs={[
          { key: 'mine', text: 'Assigned to me', icon: 'TaskManager', count: mine.length },
          { key: 'open', text: 'Open to take', icon: 'AddFriend', count: open.length, locked: !has('selfAssign') }
        ]}
      />
      {!list.length ? (
        <Card>
          <Empty icon="TaskManager" title={tab === 'mine' ? 'No tasks assigned to you' : 'No open tasks to take'}>
            {tab === 'mine' ? 'You can still log time on your projects without a task.' : 'Project managers mark tasks as self-assignable.'}
          </Empty>
        </Card>
      ) : (
        <div className="ts-table-wrap">
          <table className="ts-table">
            <thead>
              <tr>
                <th>Task</th>
                <th>Project</th>
                <th>Due</th>
                <th>Status</th>
                <th style={{ width: 170 }}>Logged / planned</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {list.map((t) => {
                const u = (used.data || {})[t.id] || 0;
                return (
                  <tr key={t.id} style={{ cursor: 'default' }}>
                    <td>
                      <div className="ts-subject">{t.title}</div>
                      {t.description && <div className="ts-muted ts-small">{t.description}</div>}
                    </td>
                    <td>
                      <button className="ts-link" onClick={() => navigate(`projects/${t.projectId}`)}>
                        {project(t.projectId)}
                      </button>
                    </td>
                    <td className={t.dueDate && t.dueDate < t0 && t.status !== 'Done' ? 'ts-status-bad' : 'ts-muted'}>{t.dueDate ? formatDay(t.dueDate, true) : '—'}</td>
                    <td>
                      {tab === 'mine' ? (
                        <Dropdown selectedKey={t.status} options={['Open', 'In progress', 'Done'].map((s) => ({ key: s, text: s }))} onChange={(_, o) => o && void setStatus(t, o.key as ITask['status'])} disabled={busy} styles={{ root: { width: 130 } }} />
                      ) : (
                        <StatusBadge status={t.status} />
                      )}
                    </td>
                    <td>
                      <div className="ts-small">
                        {hours(u)} h{t.plannedHours ? ` / ${hours(t.plannedHours)} h` : ''}
                      </div>
                      {t.plannedHours > 0 && <Progress value={u} max={t.plannedHours} />}
                    </td>
                    <td style={{ textAlign: 'right', whiteSpace: 'nowrap' }}>
                      {tab === 'open' ? (
                        <PrimaryButton text="Assign to me" onClick={() => void take(t)} disabled={busy} />
                      ) : (
                        <>
                          <DefaultButton text="Log time" onClick={() => navigate('timesheet')} />{' '}
                          {t.selfAssign && <DefaultButton text="Drop" onClick={() => void drop(t)} disabled={busy} />}
                        </>
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
        <TaskEditor
          task={editing}
          projectChoices={canCreate.map((p) => p.id)}
          onDismiss={() => setEditing(undefined)}
          onSaved={() => {
            setEditing(undefined);
            void reloadRef();
          }}
        />
      )}
    </>
  );
}
