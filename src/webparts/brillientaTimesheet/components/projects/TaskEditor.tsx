import * as React from 'react';
import { DatePicker, DefaultButton, Dropdown, Panel, PanelType, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { ITask } from '../../models/types';
import { PeopleField } from '../common/Pickers';
import { formatDay, fromKey, toKey } from '../../services/dates';
import { sameEmail } from '../../services/access';

export function newTask(projectId: number): ITask {
  return { id: 0, title: '', projectId, description: '', assigneeEmails: [], billable: true, plannedHours: 0, status: 'Open', dueDate: '', selfAssign: true, plannerTaskId: '', plannerPlanId: '' };
}

/** Create or edit a task. Assignees who are new to the task are told (email / Teams / bell). */
export function TaskEditor(props: { task: ITask; projectChoices: number[]; onDismiss: () => void; onSaved: () => void }): JSX.Element {
  const { svc, ref, notifier, toast, me, has, graph, config } = useApp();
  const [t, setT] = React.useState<ITask>(props.task);
  const [busy, setBusy] = React.useState(false);
  const set = (patch: Partial<ITask>): void => setT({ ...t, ...patch });
  const project = ref.projects.filter((p) => p.id === t.projectId)[0];
  const save = async (): Promise<void> => {
    if (!t.title.trim() || !t.projectId) {
      toast('A task needs a name and a project.', 'warning');
      return;
    }
    setBusy(true);
    try {
      await svc.saveTask({ ...t, title: t.title.trim() });
      const added = t.assigneeEmails.filter((e) => !props.task.assigneeEmails.some((x) => sameEmail(x, e)) && !sameEmail(e, me.email));
      if (added.length) {
        await notifier.send({
          template: 'taskAssigned',
          to: added.map((e) => ({ email: e, name: (ref.users.filter((u) => sameEmail(u.email, e))[0] || { name: e }).name })),
          vars: { task: t.title, project: project ? project.title : '' },
          path: 'tasks',
          cta: 'Open my tasks'
        });
        notifier.event('task.assigned', { task: t.title, projectId: t.projectId, assignees: added });
      }
      // Planner write-back: a task linked to Planner is completed there too.
      if (t.plannerTaskId && graph && has('planner') && config.integrations.plannerWriteBack && t.status !== props.task.status) {
        await graph.setPlannerProgress(t.plannerTaskId, t.status === 'Done' ? 100 : t.status === 'In progress' ? 50 : 0).catch((e) => toast('Planner not updated: ' + (e as Error).message, 'warning'));
      }
      await svc.log(t.id ? 'Task updated' : 'Task created', `${t.title} (${project ? project.title : t.projectId})`);
      props.onSaved();
    } catch (e) {
      toast((e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };
  return (
    <Panel
      isOpen
      type={PanelType.medium}
      headerText={t.id ? 'Edit task' : 'New task'}
      onDismiss={props.onDismiss}
      isFooterAtBottom
      onRenderFooterContent={() => (
        <div className="ts-row">
          <PrimaryButton text="Save" onClick={() => void save()} disabled={busy} />
          <DefaultButton text="Cancel" onClick={props.onDismiss} />
        </div>
      )}
    >
      <div className="ts-stack">
        <Dropdown
          label="Project"
          required
          selectedKey={t.projectId || null}
          disabled={props.projectChoices.length <= 1 && !!t.projectId}
          options={ref.projects.filter((p) => props.projectChoices.indexOf(p.id) >= 0).map((p) => ({ key: p.id, text: p.title }))}
          onChange={(_, o) => o && set({ projectId: Number(o.key) })}
        />
        <TextField label="Task" required value={t.title} onChange={(_, v) => set({ title: v || '' })} />
        <TextField label="Description" multiline rows={3} value={t.description} onChange={(_, v) => set({ description: v || '' })} />
        <PeopleField label="Assignees" value={t.assigneeEmails} onChange={(v) => set({ assigneeEmails: v })} hint="Nobody assigned: everyone on the project can log time on it." onlyKnown />
        <div className="ts-form-grid">
          <Dropdown label="Status" selectedKey={t.status} options={['Open', 'In progress', 'Done'].map((s) => ({ key: s, text: s }))} onChange={(_, o) => o && set({ status: o.key as ITask['status'] })} />
          <TextField label="Planned hours" type="number" value={String(t.plannedHours || '')} onChange={(_, v) => set({ plannedHours: Number(v) || 0 })} />
          <DatePicker label="Due date" value={t.dueDate ? fromKey(t.dueDate) : undefined} onSelectDate={(d) => set({ dueDate: d ? toKey(d) : '' })} formatDate={(d) => (d ? formatDay(toKey(d), true) : '')} />
          <div />
          <Toggle label="Billable" checked={t.billable} onChange={(_, v) => set({ billable: !!v })} disabled={!!project && !project.billable} />
          {has('selfAssign') && <Toggle label="People can assign it to themselves" checked={t.selfAssign} onChange={(_, v) => set({ selfAssign: !!v })} />}
        </div>
        {t.plannerTaskId && <div className="ts-muted ts-small">Linked to a Microsoft Planner task.</div>}
      </div>
    </Panel>
  );
}
