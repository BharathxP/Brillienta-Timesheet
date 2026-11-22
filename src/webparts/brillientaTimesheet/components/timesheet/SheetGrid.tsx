import * as React from 'react';
import { Checkbox, Dropdown, Icon, IconButton, IDropdownOption, TextField, TooltipHost } from '@fluentui/react';
import { DateKey, ICustomColumn, ITimeEntry, ITimeOff, ITsUser } from '../../models/types';
import { useApp } from '../AppContext';
import { DAY_SHORT, dayOfWeek, fromKey, today } from '../../services/dates';
import { dayTotals, holidayOn, isTrackingDay, leaveHours, limitsFor, parseHours, snapHours } from '../../services/rules';
import { defaultBillable, projectsFor, tasksFor } from '../../services/access';
import { entryTotal } from '../../services/mappers';
import { hours } from '../../services/format';

export interface ISheetGridProps {
  days: DateKey[];
  lines: ITimeEntry[];
  owner: { email: string; user?: ITsUser };
  readOnly: boolean;
  timeOff: ITimeOff[];
  onChange?: (lines: ITimeEntry[]) => void;
}

export function blankLine(email: string, weekStart: string): ITimeEntry {
  return { id: 0, sheetId: 0, userEmail: email, weekStart, projectId: 0, taskId: 0, activityId: 0, billable: true, hours: {}, notes: {}, total: 0, custom: {}, costRate: 0, billRate: 0 };
}

/** The week's lines as an editable grid: project / task / activity / billable / custom columns, hours per day, notes. */
export function SheetGrid(props: ISheetGridProps): JSX.Element {
  const { ref, config, has } = useApp();
  const { days, lines, readOnly, owner } = props;
  const [notesOpen, setNotesOpen] = React.useState<Record<number, boolean>>({});
  const [drafts, setDrafts] = React.useState<Record<string, string>>({});
  const activitiesOn = has('activityTracking') && ref.activities.some((a) => a.active && !a.isTimeOff);
  const billableOn = has('billable');
  const columns: ICustomColumn[] = has('customColumns') ? config.customColumns.filter((c) => c.level === 'entry') : [];
  const limits = limitsFor(config, owner.user);
  const totals = dayTotals(lines, days);
  const timeOffActivity = ref.activities.filter((a) => a.isTimeOff)[0];
  const t = today();

  const myProjects = projectsFor(ref, owner.email, days[0]);
  const projectOptions: IDropdownOption[] = myProjects.map((p) => ({ key: p.id, text: p.code ? `${p.code} · ${p.title}` : p.title }));
  const activityOptions: IDropdownOption[] = [{ key: 0, text: '—' }].concat(ref.activities.filter((a) => a.active).map((a) => ({ key: a.id, text: a.title })));

  const update = (i: number, patch: Partial<ITimeEntry>): void => {
    if (!props.onChange) return;
    const next = lines.slice();
    next[i] = { ...next[i], ...patch };
    next[i].total = entryTotal(next[i]);
    props.onChange(next);
  };
  const setProject = (i: number, projectId: number): void => {
    const p = ref.projects.filter((x) => x.id === projectId)[0];
    const a = ref.activities.filter((x) => x.id === lines[i].activityId)[0];
    update(i, { projectId, taskId: 0, billable: billableOn ? defaultBillable(p, undefined, a ? a.billable : undefined) : false });
  };
  const setTask = (i: number, taskId: number): void => {
    const p = ref.projects.filter((x) => x.id === lines[i].projectId)[0];
    const task = ref.tasks.filter((x) => x.id === taskId)[0];
    update(i, { taskId, billable: billableOn ? defaultBillable(p, task, undefined) : false });
  };
  const setActivity = (i: number, activityId: number): void => {
    const a = ref.activities.filter((x) => x.id === activityId)[0];
    const p = ref.projects.filter((x) => x.id === lines[i].projectId)[0];
    const task = ref.tasks.filter((x) => x.id === lines[i].taskId)[0];
    update(i, { activityId, billable: billableOn ? (a && a.isTimeOff ? false : defaultBillable(p, task, a ? a.billable : undefined)) : false });
  };
  const commitHours = (i: number, day: DateKey, text: string): void => {
    const key = `${i}|${day}`;
    const n = parseHours(text);
    setDrafts((d) => {
      const c = { ...d };
      delete c[key];
      return c;
    });
    if (isNaN(n)) return;
    update(i, { hours: { ...lines[i].hours, [day]: snapHours(n, config.increment) } });
  };

  const projectLabel = (e: ITimeEntry): { title: string; sub: string; color: string } => {
    const p = ref.projects.filter((x) => x.id === e.projectId)[0];
    const task = ref.tasks.filter((x) => x.id === e.taskId)[0];
    const a = ref.activities.filter((x) => x.id === e.activityId)[0];
    return {
      title: p ? (p.code ? `${p.code} · ${p.title}` : p.title) : a && a.isTimeOff ? 'Time off' : e.projectId ? '(removed project)' : '—',
      sub: [task ? task.title : '', a ? a.title : ''].filter((x) => !!x).join(' · '),
      color: (p && p.color) || (a && a.color) || '#c8d0d2'
    };
  };

  return (
    <div className="ts-sheet-wrap">
      <table className="ts-sheet">
        <thead>
          <tr>
            <th className="ts-col-line">Project · task{activitiesOn ? ' · activity' : ''}</th>
            {billableOn && <th style={{ width: 70, textAlign: 'center' }}>Billable</th>}
            {columns.map((c) => (
              <th key={c.key} style={{ minWidth: 110 }}>
                {c.label}
                {c.required ? ' *' : ''}
              </th>
            ))}
            {days.map((d) => {
              const off = !isTrackingDay(config, d, owner.user);
              const hol = holidayOn(config, d);
              const leave = leaveHours(props.timeOff, d, config.standardDay);
              return (
                <th key={d} className={'ts-day' + (d === t ? ' is-today' : '') + (off ? ' is-off' : '') + (leave ? ' is-leave' : '')} title={hol || (leave ? 'Approved time off' : off ? 'Not a working day' : '')}>
                  {DAY_SHORT[dayOfWeek(d)]}
                  <strong>{fromKey(d).getDate()}</strong>
                  {hol ? <span className="ts-small">Holiday</span> : leave ? <span className="ts-small">Leave</span> : null}
                </th>
              );
            })}
            <th className="ts-col-total">Total</th>
            {!readOnly && <th className="ts-col-actions" />}
          </tr>
        </thead>
        <tbody>
          {lines.map((e, i) => {
            const label = projectLabel(e);
            const isTimeOffLine = !!timeOffActivity && e.activityId === timeOffActivity.id;
            const taskOptions: IDropdownOption[] = [{ key: 0, text: e.projectId ? '— no task —' : 'Choose a project first' }].concat(
              tasksFor(ref, e.projectId, owner.email).map((x) => ({ key: x.id, text: x.title }))
            );
            if (e.taskId && !taskOptions.some((o) => o.key === e.taskId)) {
              const task = ref.tasks.filter((x) => x.id === e.taskId)[0];
              taskOptions.push({ key: e.taskId, text: task ? task.title : '(removed task)' });
            }
            const projOpts = projectOptions.slice();
            if (e.projectId && !projOpts.some((o) => o.key === e.projectId)) projOpts.unshift({ key: e.projectId, text: label.title });
            const notesShown = !!notesOpen[i];
            const hasNotes = Object.keys(e.notes).some((k) => !!e.notes[k]);
            return (
              <React.Fragment key={i}>
                <tr>
                  <td className="ts-col-line">
                    {readOnly ? (
                      <div className="ts-line-static">
                        <div className="ts-line-project">
                          <span className="ts-swatch" style={{ background: label.color }} />
                          {label.title}
                        </div>
                        {label.sub && <div className="ts-line-sub">{label.sub}</div>}
                        {hasNotes && (
                          <button className="ts-link ts-small" onClick={() => setNotesOpen({ ...notesOpen, [i]: !notesShown })}>
                            {notesShown ? 'Hide notes' : 'Notes'}
                          </button>
                        )}
                      </div>
                    ) : (
                      <div className="ts-line">
                        {!isTimeOffLine && <Dropdown placeholder="Choose a project" options={projOpts} selectedKey={e.projectId || null} onChange={(_, o) => o && setProject(i, Number(o.key))} ariaLabel="Project" />}
                        {isTimeOffLine && (
                          <div className="ts-line-static">
                            <div className="ts-line-project">
                              <span className="ts-swatch" style={{ background: label.color }} />
                              Time off (from Time off Manager 365)
                            </div>
                          </div>
                        )}
                        <div className="ts-row" style={{ flexWrap: 'nowrap' }}>
                          {!isTimeOffLine && (
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <Dropdown options={taskOptions} selectedKey={e.taskId} disabled={!e.projectId} onChange={(_, o) => o && setTask(i, Number(o.key))} ariaLabel="Task" />
                            </div>
                          )}
                          {activitiesOn && !isTimeOffLine && (
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <Dropdown options={activityOptions.filter((o) => !timeOffActivity || o.key !== timeOffActivity.id)} selectedKey={e.activityId} placeholder="Activity" onChange={(_, o) => o && setActivity(i, Number(o.key))} ariaLabel="Activity" />
                            </div>
                          )}
                        </div>
                      </div>
                    )}
                  </td>
                  {billableOn && (
                    <td style={{ textAlign: 'center' }}>
                      {readOnly || !config.allowBillableOverride || isTimeOffLine ? (
                        <TooltipHost content={e.billable ? 'Billable' : 'Non-billable'}>
                          <Icon iconName={e.billable ? 'Money' : 'Blocked2'} style={{ color: e.billable ? 'var(--ts-success)' : 'var(--ts-faint)' }} />
                        </TooltipHost>
                      ) : (
                        <Checkbox checked={e.billable} onChange={(_, v) => update(i, { billable: !!v })} ariaLabel="Billable" styles={{ root: { justifyContent: 'center' } }} />
                      )}
                    </td>
                  )}
                  {columns.map((c) => (
                    <td key={c.key}>
                      <CustomCell column={c} value={e.custom[c.key]} readOnly={readOnly} onChange={(v) => update(i, { custom: { ...e.custom, [c.key]: v } })} />
                    </td>
                  ))}
                  {days.map((d) => {
                    const key = `${i}|${d}`;
                    const v = e.hours[d] || 0;
                    const off = !isTrackingDay(config, d, owner.user);
                    const blocked = off && !limits.allowNonWorkingDays && !v;
                    return (
                      <td key={d} className={'ts-day' + (off ? ' is-off' : '')}>
                        <input
                          className={'ts-cell' + (e.notes[d] ? ' has-note' : '') + (!v ? ' is-empty' : '')}
                          value={drafts[key] !== undefined ? drafts[key] : v ? hours(v) : readOnly ? '–' : ''}
                          disabled={readOnly || blocked}
                          inputMode="decimal"
                          aria-label={`Hours on ${d}`}
                          title={e.notes[d] || (blocked ? 'Time cannot be logged on this day' : '')}
                          onChange={(ev) => setDrafts({ ...drafts, [key]: ev.target.value })}
                          onBlur={(ev) => commitHours(i, d, ev.target.value)}
                          onKeyDown={(ev) => {
                            if (ev.key === 'Enter') (ev.target as HTMLInputElement).blur();
                          }}
                        />
                      </td>
                    );
                  })}
                  <td className="ts-col-total">{hours(entryTotal(e))}</td>
                  {!readOnly && (
                    <td className="ts-col-actions">
                      <IconButton iconProps={{ iconName: 'EditNote' }} title="Notes for each day" ariaLabel="Notes" checked={notesShown || hasNotes} onClick={() => setNotesOpen({ ...notesOpen, [i]: !notesShown })} />
                      <IconButton
                        iconProps={{ iconName: 'More' }}
                        title="More"
                        ariaLabel="More actions"
                        menuProps={{
                          items: [
                            { key: 'dup', text: 'Duplicate line', iconProps: { iconName: 'Copy' }, onClick: () => props.onChange && props.onChange(lines.slice(0, i + 1).concat([{ ...e, id: 0 }], lines.slice(i + 1))) },
                            { key: 'clear', text: 'Clear hours', iconProps: { iconName: 'EraseTool' }, onClick: () => update(i, { hours: {}, notes: {} }) },
                            { key: 'del', text: 'Remove line', iconProps: { iconName: 'Delete' }, onClick: () => props.onChange && props.onChange(lines.filter((_, j) => j !== i)) }
                          ]
                        }}
                        onRenderMenuIcon={() => null}
                      />
                    </td>
                  )}
                </tr>
                {notesShown && (
                  <tr>
                    <td colSpan={1 + (billableOn ? 1 : 0) + columns.length} className="ts-small ts-muted" style={{ textAlign: 'right' }}>
                      Notes
                    </td>
                    {days.map((d) => (
                      <td key={d} className="ts-day">
                        {readOnly ? (
                          <span className="ts-small">{e.notes[d] || ''}</span>
                        ) : (
                          <TextField
                            value={e.notes[d] || ''}
                            onChange={(_, v) => update(i, { notes: { ...e.notes, [d]: v || '' } })}
                            ariaLabel={`Note for ${d}`}
                            styles={{ fieldGroup: { height: 30, minWidth: 62 }, field: { fontSize: 12, padding: '0 6px' } }}
                          />
                        )}
                      </td>
                    ))}
                    <td colSpan={readOnly ? 1 : 2} />
                  </tr>
                )}
              </React.Fragment>
            );
          })}
          <tr className="ts-sheet-total">
            <td colSpan={1 + (billableOn ? 1 : 0) + columns.length}>Total</td>
            {days.map((d) => {
              const leave = leaveHours(props.timeOff, d, config.standardDay);
              const onSheet = !!timeOffActivity && lines.some((e) => e.activityId === timeOffActivity.id && (e.hours[d] || 0) > 0);
              const v = totals[d] + (onSheet ? 0 : leave);
              const over = limits.dailyMax > 0 && v > limits.dailyMax;
              return (
                <td key={d} className={'ts-day' + (over ? ' is-over' : '')} title={over ? `More than ${limits.dailyMax} h` : undefined}>
                  {hours(v)}
                </td>
              );
            })}
            <td className="ts-col-total">{hours(lines.reduce((s, e) => s + entryTotal(e), 0))}</td>
            {!readOnly && <td />}
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function CustomCell(props: { column: ICustomColumn; value: string | number | boolean | undefined; readOnly: boolean; onChange: (v: string | number | boolean) => void }): JSX.Element {
  const c = props.column;
  if (props.readOnly) {
    const v = props.value;
    return <span>{v === undefined || v === '' ? '–' : typeof v === 'boolean' ? (v ? 'Yes' : 'No') : String(v)}</span>;
  }
  switch (c.type) {
    case 'boolean':
      return <Checkbox checked={!!props.value} onChange={(_, v) => props.onChange(!!v)} ariaLabel={c.label} />;
    case 'choice':
      return <Dropdown options={[{ key: '', text: '—' }].concat(c.choices.map((x) => ({ key: x, text: x })))} selectedKey={String(props.value || '')} onChange={(_, o) => o && props.onChange(String(o.key))} ariaLabel={c.label} />;
    case 'number':
      return <TextField type="number" value={props.value === undefined ? '' : String(props.value)} onChange={(_, v) => props.onChange(v === '' || v === undefined ? '' : Number(v))} ariaLabel={c.label} />;
    case 'date':
      return <TextField type="date" value={String(props.value || '')} onChange={(_, v) => props.onChange(v || '')} ariaLabel={c.label} />;
    default:
      return <TextField value={String(props.value || '')} onChange={(_, v) => props.onChange(v || '')} ariaLabel={c.label} />;
  }
}
