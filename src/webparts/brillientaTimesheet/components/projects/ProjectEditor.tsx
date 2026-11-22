import * as React from 'react';
import { Checkbox, DatePicker, DefaultButton, Dropdown, Panel, PanelType, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { ICustomColumn, IProject } from '../../models/types';
import { PeopleField } from '../common/Pickers';
import { ColorInput, Field, Locked } from '../common/Ui';
import { formatDay, fromKey, toKey } from '../../services/dates';
import { PALETTE } from '../../schema/defaults';
import { sameEmail } from '../../services/access';

export function newProject(managerEmail: string, color: string): IProject {
  return {
    id: 0,
    title: '',
    code: '',
    client: '',
    description: '',
    programId: 0,
    managerEmails: managerEmail ? [managerEmail] : [],
    memberEmails: [],
    billable: true,
    status: 'Active',
    startDate: '',
    endDate: '',
    plannedHours: 0,
    hourlyCost: 0,
    billRate: 0,
    memberRates: {},
    color,
    custom: {},
    endNotices: []
  };
}

export function CustomInput(props: { column: ICustomColumn; value: string | number | boolean | undefined; onChange: (v: string | number | boolean) => void }): JSX.Element {
  const c = props.column;
  switch (c.type) {
    case 'boolean':
      return <Checkbox label={c.label} checked={!!props.value} onChange={(_, v) => props.onChange(!!v)} />;
    case 'choice':
      return <Dropdown label={c.label} required={c.required} options={[{ key: '', text: '—' }].concat(c.choices.map((x) => ({ key: x, text: x })))} selectedKey={String(props.value || '')} onChange={(_, o) => o && props.onChange(String(o.key))} />;
    case 'number':
      return <TextField label={c.label} required={c.required} type="number" value={props.value === undefined ? '' : String(props.value)} onChange={(_, v) => props.onChange(v ? Number(v) : '')} />;
    case 'date':
      return <TextField label={c.label} required={c.required} type="date" value={String(props.value || '')} onChange={(_, v) => props.onChange(v || '')} />;
    default:
      return <TextField label={c.label} required={c.required} value={String(props.value || '')} onChange={(_, v) => props.onChange(v || '')} />;
  }
}

/** Create or edit a project: details, people, dates and planned hours, cost and bill rates, custom columns. */
export function ProjectEditor(props: { project: IProject; onDismiss: () => void; onSaved: (id: number) => void }): JSX.Element {
  const { svc, ref, config, has, toast, notifier, role, me } = useApp();
  const [p, setP] = React.useState<IProject>(props.project);
  const [busy, setBusy] = React.useState(false);
  const set = (patch: Partial<IProject>): void => setP({ ...p, ...patch });
  const columns = has('customColumns') ? config.customColumns.filter((c) => c.level === 'project') : [];
  const programs = role === 'admin' ? ref.programs : ref.programs.filter((g) => g.managerEmails.some((e) => sameEmail(e, me.email)) || g.id === p.programId);
  const save = async (): Promise<void> => {
    if (!p.title.trim()) {
      toast('A project needs a name.', 'warning');
      return;
    }
    const missing = columns.filter((c) => c.required && (p.custom[c.key] === undefined || p.custom[c.key] === ''));
    if (missing.length) {
      toast(`Fill in ${missing.map((c) => c.label).join(', ')}.`, 'warning');
      return;
    }
    if (p.startDate && p.endDate && p.endDate < p.startDate) {
      toast('The end date is before the start date.', 'warning');
      return;
    }
    if (!p.id && ref.projects.some((x) => p.code && x.code && x.code.toLowerCase() === p.code.toLowerCase())) {
      toast(`Another project already has the code ${p.code}.`, 'warning');
      return;
    }
    setBusy(true);
    try {
      // A new end date starts its notices again.
      const endChanged = props.project.endDate !== p.endDate;
      const id = await svc.saveProject({ ...p, title: p.title.trim(), endNotices: endChanged ? p.endNotices.filter((n) => n === 'planned') : p.endNotices });
      await svc.log(p.id ? 'Project updated' : 'Project created', p.title);
      if (!p.id) notifier.event('project.created', { projectId: id, title: p.title, code: p.code, client: p.client });
      props.onSaved(id);
    } catch (e) {
      toast((e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };
  const rateRows = p.memberEmails.length ? p.memberEmails : Object.keys(p.memberRates);
  return (
    <Panel
      isOpen
      type={PanelType.large}
      headerText={p.id ? `Edit ${props.project.title}` : 'New project'}
      onDismiss={props.onDismiss}
      isFooterAtBottom
      onRenderFooterContent={() => (
        <div className="ts-row">
          <PrimaryButton text="Save project" onClick={() => void save()} disabled={busy} />
          <DefaultButton text="Cancel" onClick={props.onDismiss} />
        </div>
      )}
    >
      <div className="ts-stack" style={{ paddingTop: 8 }}>
        <div className="ts-form-grid">
          <TextField label="Project name" required value={p.title} onChange={(_, v) => set({ title: v || '' })} className="is-wide" />
          <TextField label="Code" value={p.code} onChange={(_, v) => set({ code: v || '' })} placeholder="e.g. ACM-001" />
          <TextField label="Client" value={p.client} onChange={(_, v) => set({ client: v || '' })} list="ts-clients" />
          <datalist id="ts-clients">
            {ref.projects
              .map((x) => x.client)
              .filter((c, i, a) => !!c && a.indexOf(c) === i)
              .map((c) => (
                <option key={c} value={c} />
              ))}
          </datalist>
          {has('programManager') && (
            <Dropdown label="Programme" selectedKey={p.programId} options={[{ key: 0, text: '— none —' }].concat(programs.map((g) => ({ key: g.id, text: g.title })))} onChange={(_, o) => o && set({ programId: Number(o.key) })} />
          )}
          <Dropdown
            label="Status"
            selectedKey={p.status}
            options={[
              { key: 'Active', text: 'Active' },
              { key: 'Completed', text: 'Completed' },
              { key: 'Disabled', text: 'Disabled (no new time)', disabled: !has('disableProjects') }
            ]}
            onChange={(_, o) => o && set({ status: o.key as IProject['status'] })}
          />
          <TextField label="Description" multiline rows={3} value={p.description} onChange={(_, v) => set({ description: v || '' })} className="is-wide" />
        </div>
        <h3 className="ts-card-title" style={{ marginTop: 10 }}>People</h3>
        <PeopleField label="Project managers" value={p.managerEmails} onChange={(v) => set({ managerEmails: v })} hint="Project managers edit the project, its tasks and approve time logged on it." />
        <PeopleField label="Members" value={p.memberEmails} onChange={(v) => set({ memberEmails: v })} hint="Nobody listed: everyone can log time on this project." />
        <h3 className="ts-card-title" style={{ marginTop: 10 }}>Dates and planned hours</h3>
        <div className="ts-form-grid">
          <DatePicker label="Start date" value={p.startDate ? fromKey(p.startDate) : undefined} onSelectDate={(d) => set({ startDate: d ? toKey(d) : '' })} formatDate={(d) => (d ? formatDay(toKey(d), true) : '')} />
          {has('projectEndDates') ? (
            <DatePicker label="End date" value={p.endDate ? fromKey(p.endDate) : undefined} onSelectDate={(d) => set({ endDate: d ? toKey(d) : '' })} formatDate={(d) => (d ? formatDay(toKey(d), true) : '')} />
          ) : (
            <Field label="End date">
              <Locked feature="projectEndDates">Project end dates and notifications.</Locked>
            </Field>
          )}
          {has('plannedHours') ? (
            <TextField label="Planned hours" type="number" value={String(p.plannedHours || '')} onChange={(_, v) => set({ plannedHours: Number(v) || 0 })} description="Managers are told when logged hours pass the plan." />
          ) : (
            <Field label="Planned hours">
              <Locked feature="plannedHours">Planned project hours tracking.</Locked>
            </Field>
          )}
          <Field label="Colour">
            <div className="ts-row">
              <ColorInput value={p.color} onChange={(v) => set({ color: v })} />
              {PALETTE.slice(0, 8).map((c) => (
                <button key={c} className="ts-iconbtn" style={{ background: c, width: 22, height: 22 }} onClick={() => set({ color: c })} aria-label={c} />
              ))}
            </div>
          </Field>
        </div>
        <h3 className="ts-card-title" style={{ marginTop: 10 }}>Billing and cost</h3>
        <div className="ts-form-grid">
          <Toggle label="Billable project" checked={p.billable} onChange={(_, v) => set({ billable: !!v })} />
          <div />
          <TextField label={`Flat hourly cost (${config.currency})`} type="number" value={String(p.hourlyCost || '')} onChange={(_, v) => set({ hourlyCost: Number(v) || 0 })} description="Used for cost when a person has no hourly cost of their own (Admin › Costs & rates)." />
          <TextField label={`Hourly bill rate (${config.currency})`} type="number" value={String(p.billRate || '')} onChange={(_, v) => set({ billRate: Number(v) || 0 })} />
        </div>
        {has('userRates') && rateRows.length > 0 && (
          <>
            <div className="ms-Label">Person-specific bill rates on this project</div>
            <div className="ts-table-wrap">
              <table className="ts-table">
                <tbody>
                  {rateRows.map((e) => (
                    <tr key={e} style={{ cursor: 'default' }}>
                      <td>{(ref.users.filter((u) => sameEmail(u.email, e))[0] || { name: e }).name}</td>
                      <td style={{ width: 160 }}>
                        <TextField
                          type="number"
                          placeholder="Default"
                          value={p.memberRates[e] ? String(p.memberRates[e]) : ''}
                          onChange={(_, v) => {
                            const r = { ...p.memberRates };
                            if (Number(v) > 0) r[e] = Number(v);
                            else delete r[e];
                            set({ memberRates: r });
                          }}
                          ariaLabel={`Rate for ${e}`}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
        {columns.length > 0 && (
          <>
            <h3 className="ts-card-title" style={{ marginTop: 10 }}>More details</h3>
            <div className="ts-form-grid">
              {columns.map((c) => (
                <CustomInput key={c.key} column={c} value={p.custom[c.key]} onChange={(v) => set({ custom: { ...p.custom, [c.key]: v } })} />
              ))}
            </div>
          </>
        )}
      </div>
    </Panel>
  );
}
