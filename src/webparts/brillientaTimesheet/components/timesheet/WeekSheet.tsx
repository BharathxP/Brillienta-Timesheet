import * as React from 'react';
import { ChoiceGroup, DatePicker, DayOfWeek, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, Icon, MessageBar, MessageBarType, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { useApp, buildQuery } from '../AppContext';
import { ITimeEntry, ITimeOff, ITimesheet } from '../../models/types';
import { addDays, formatDay, formatWeek, fromKey, isoWeek, toKey, today, weekDays, weekStartOf } from '../../services/dates';
import { IIssue, blocking, isTrackingDay, leaveHours, limitsFor, validateWeek } from '../../services/rules';
import { entryTotal } from '../../services/mappers';
import { hours, formatDateTime } from '../../services/format';
import { projectsFor } from '../../services/access';
import { Card, ErrorBox, Kpi, Loading, PageHeader, StatusBadge, useConfirm } from '../common/Ui';
import { SheetGrid, blankLine } from './SheetGrid';
import { Comments } from './Comments';

interface ILoaded {
  sheet?: ITimesheet;
  entries: ITimeEntry[];
  timeOff: ITimeOff[];
}

/** My timesheet: one week, editable until submitted. */
export function WeekSheet(): JSX.Element {
  const app = useApp();
  const { config, svc, me, myUser, route, navigate, workflow, toast, ref, has, role } = app;
  const week = route.query.week && /^\d{4}-\d{2}-\d{2}$/.test(route.query.week) ? weekStartOf(route.query.week, config.weekStart) : weekStartOf(today(), config.weekStart);
  const days = weekDays(week);
  const [loaded, setLoaded] = React.useState<ILoaded | undefined>();
  const [error, setError] = React.useState<string | undefined>();
  const [lines, setLines] = React.useState<ITimeEntry[]>([]);
  const [dirty, setDirty] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [submitIssues, setSubmitIssues] = React.useState<IIssue[] | undefined>();
  const [cloneOpen, setCloneOpen] = React.useState(false);
  const [submitOpen, setSubmitOpen] = React.useState(false);
  const [comment, setComment] = React.useState('');
  const [confirm, confirmDialog] = useConfirm();
  const [tick, setTick] = React.useState(0);
  const timeOffActivity = ref.activities.filter((a) => a.isTimeOff)[0];

  React.useEffect(() => {
    let alive = true;
    setLoaded(undefined);
    setError(undefined);
    setSubmitIssues(undefined);
    (async () => {
      const [{ sheet, entries }, timeOff] = await Promise.all([
        svc.getSheet(me.email, week),
        has('timeOff') && config.timeOff.enabled ? app.timeOff.load(config.timeOff, week, addDays(week, 6), me.email).catch(() => [] as ITimeOff[]) : Promise.resolve([] as ITimeOff[])
      ]);
      if (!alive) return;
      let start = entries.length ? entries : [blankLine(me.email, week)];
      // Approved leave goes on the sheet on the time-off activity (Time off Manager 365, "put leave on the timesheet").
      const editable = !sheet || sheet.status === 'Draft' || sheet.status === 'Rejected';
      if (editable && timeOff.length && config.timeOff.autoFill && timeOffActivity && !entries.some((e) => e.activityId === timeOffActivity.id)) {
        const h: Record<string, number> = {};
        days.forEach((d) => {
          const l = leaveHours(timeOff, d, config.standardDay);
          if (l && isTrackingDay(config, d, myUser)) h[d] = l;
        });
        if (Object.keys(h).length) {
          start = start.filter((e) => e.projectId || entryTotal(e) > 0).concat([{ ...blankLine(me.email, week), activityId: timeOffActivity.id, billable: false, hours: h }]);
          if (!start.some((e) => e.activityId !== timeOffActivity.id)) start.unshift(blankLine(me.email, week));
          setDirty(true);
        } else setDirty(false);
      } else setDirty(false);
      setLines(start);
      setLoaded({ sheet, entries, timeOff });
    })().catch((e) => alive && setError((e as Error).message));
    return () => {
      alive = false;
    };
  }, [week, me.email, tick]);

  // Unsaved changes: warn before leaving the page.
  React.useEffect(() => {
    if (!dirty) return;
    const onLeave = (e: BeforeUnloadEvent): void => {
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', onLeave);
    return () => window.removeEventListener('beforeunload', onLeave);
  }, [dirty]);

  const goWeek = async (w: string): Promise<void> => {
    if (dirty && !(await confirm('Discard your changes?', 'This week has changes that are not saved.', 'Discard'))) return;
    navigate('timesheet' + buildQuery({ week: w }));
  };

  if (error) return <ErrorBox error={error} onRetry={() => setTick((t) => t + 1)} />;
  if (!loaded) return <Loading label="Loading your week…" />;

  const sheet = loaded.sheet;
  const status = sheet ? sheet.status : 'Draft';
  const editable = status === 'Draft' || status === 'Rejected';
  const limits = limitsFor(config, myUser);
  const liveIssues = validateWeek({ cfg: config, user: myUser, days, entries: lines, timeOff: loaded.timeOff, timeOffActivityId: timeOffActivity ? timeOffActivity.id : 0 });
  const issues = submitIssues || liveIssues;
  const total = lines.reduce((s, e) => s + entryTotal(e), 0);
  const billable = lines.filter((e) => e.billable).reduce((s, e) => s + entryTotal(e), 0);
  const trackingDays = days.filter((d) => isTrackingDay(config, d, myUser)).length;
  const target = limits.weeklyMin || trackingDays * config.standardDay;
  const actor = { email: me.email, name: me.title, role };

  const save = async (): Promise<void> => {
    setBusy(true);
    try {
      const r = await workflow.save({ email: me.email, name: me.title }, week, sheet, lines, loaded.entries);
      setLoaded({ ...loaded, sheet: r.sheet, entries: r.entries });
      setLines(r.entries.length ? r.entries : [blankLine(me.email, week)]);
      setDirty(false);
      toast('Saved.', 'success');
    } catch (e) {
      toast('Could not save: ' + (e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };

  const startSubmit = (): void => {
    const found = workflow.check(myUser, week, lines, loaded.timeOff, timeOffActivity ? timeOffActivity.id : 0);
    setSubmitIssues(found);
    if (app.overLimit) {
      toast('Your organisation has reached its user limit; ask an administrator to add users to the licence.', 'warning');
      return;
    }
    if (blocking(found)) {
      toast('Fix the problems shown before submitting.', 'warning');
      return;
    }
    setComment('');
    setSubmitOpen(true);
  };

  const submit = async (): Promise<void> => {
    setSubmitOpen(false);
    setBusy(true);
    try {
      const r = await workflow.submit(actor, week, sheet, lines, loaded.entries, submitIssues || [], comment.trim());
      setLoaded({ ...loaded, sheet: r.sheet, entries: r.entries });
      setLines(r.entries);
      setDirty(false);
      setSubmitIssues(undefined);
      toast(r.autoApproved ? 'Submitted and approved automatically.' : `Submitted for approval${r.approvers.length ? ' to ' + r.approvers.length + ' approver' + (r.approvers.length === 1 ? '' : 's') : ''}.`, 'success');
      app.bumpNotifications();
    } catch (e) {
      toast('Could not submit: ' + (e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };

  const recall = async (): Promise<void> => {
    if (!sheet) return;
    const ok = await confirm('Recall this timesheet?', 'It goes back to draft so you can change it, and must be submitted again.', 'Recall');
    if (!ok) return;
    setBusy(true);
    try {
      const next = await workflow.reopen(actor, sheet, 'Recalled by the owner', loaded.entries);
      setLoaded({ ...loaded, sheet: next });
      toast('Back to draft.', 'success');
    } catch (e) {
      toast((e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };

  // Time logged in Brillienta Task Manager this week, as lines (already in Task Manager, so never written back).
  const importTm = async (): Promise<void> => {
    setBusy(true);
    try {
      const r = await app.taskManager.importWeek(me.email, week, addDays(week, 6), ref, lines);
      if (!r.lines.length) {
        toast(r.skipped ? `${r.skipped} entries are on Task Manager projects that are not linked yet (ask an administrator to sync).` : 'Nothing new logged in Task Manager this week.', 'info');
        return;
      }
      const keep = lines.filter((e) => e.projectId || entryTotal(e) > 0);
      change(keep.concat(r.lines));
      const h = r.lines.reduce((s, l) => s + l.total, 0);
      toast(`${hours(h)} h imported from Task Manager on ${r.lines.length} line${r.lines.length === 1 ? '' : 's'}${r.skipped ? ` (${r.skipped} entries skipped: project not linked)` : ''}. Save or submit to keep them.`, 'success');
    } catch (e) {
      toast('Task Manager could not be read: ' + (e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };

  const change = (next: ITimeEntry[]): void => {
    setLines(next);
    setDirty(true);
    setSubmitIssues(undefined);
  };

  return (
    <div onKeyDown={(e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && editable) {
        e.preventDefault();
        void save();
      }
    }}>
      <PageHeader
        eyebrow={`Week ${isoWeek(week)}`}
        title="My timesheet"
        sub={`${formatWeek(week)} · ${limits.source}${limits.weeklyMax ? ` · max ${limits.weeklyMax} h/week` : ''}${limits.dailyMax && limits.dailyMax < 24 ? ` · max ${limits.dailyMax} h/day` : ''}`}
        actions={
          <>
            {editable && has('cloneWeek') && <DefaultButton iconProps={{ iconName: 'Copy' }} text="Clone a week" onClick={() => setCloneOpen(true)} disabled={busy} />}
            {editable && has('taskManager') && config.taskManager.enabled && config.taskManager.importTime && (
              <DefaultButton iconProps={{ iconName: 'TaskManager' }} text="Import from Task Manager" onClick={() => void importTm()} disabled={busy} />
            )}
            {editable && <DefaultButton iconProps={{ iconName: 'Save' }} text={dirty ? 'Save draft' : 'Saved'} onClick={() => void save()} disabled={busy || !dirty} />}
            {editable && <PrimaryButton iconProps={{ iconName: 'Send' }} text={status === 'Rejected' ? 'Resubmit' : 'Submit'} onClick={startSubmit} disabled={busy} />}
            {status === 'Submitted' && <DefaultButton iconProps={{ iconName: 'Undo' }} text="Recall" onClick={() => void recall()} disabled={busy} />}
          </>
        }
      />
      <div className="ts-weekbar">
        <DefaultButton iconProps={{ iconName: 'ChevronLeft' }} ariaLabel="Previous week" onClick={() => void goWeek(addDays(week, -7))} styles={{ root: { minWidth: 36, padding: 0 } }} />
        <DefaultButton iconProps={{ iconName: 'ChevronRight' }} ariaLabel="Next week" onClick={() => void goWeek(addDays(week, 7))} styles={{ root: { minWidth: 36, padding: 0 } }} />
        <DefaultButton text="This week" onClick={() => void goWeek(weekStartOf(today(), config.weekStart))} disabled={week === weekStartOf(today(), config.weekStart)} />
        <DatePicker
          ariaLabel="Go to a week"
          placeholder="Go to a date"
          firstDayOfWeek={config.weekStart as DayOfWeek}
          value={fromKey(week)}
          onSelectDate={(d) => d && void goWeek(weekStartOf(toKey(d), config.weekStart))}
          formatDate={(d) => (d ? formatDay(toKey(d), true) : '')}
          styles={{ root: { width: 170 } }}
        />
        <span className="ts-spacer" />
        <StatusBadge status={status} title={sheet && sheet.autoApproved ? 'Approved automatically' : undefined} />
        {sheet && sheet.autoApproved && <span className="ts-chip is-lime">Auto-approved</span>}
        {dirty && <span className="ts-chip is-soft">Unsaved changes</span>}
      </div>
      {status === 'Rejected' && sheet && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline styles={{ root: { marginBottom: 12, borderRadius: 8 } }}>
          {sheet.decidedBy} sent this week back
          {(() => {
            const c = sheet.comments.filter((x) => x.action === 'rejected').pop();
            return c && c.text ? `: “${c.text}”` : '.';
          })()}{' '}
          Correct it and resubmit.
        </MessageBar>
      )}
      {status === 'Approved' && (
        <MessageBar messageBarType={MessageBarType.success} styles={{ root: { marginBottom: 12, borderRadius: 8 } }}>
          Approved by {sheet && sheet.decidedBy} {sheet && sheet.decidedAt ? `on ${formatDateTime(sheet.decidedAt)}` : ''}. Ask an approver to reopen it if something needs to change.
        </MessageBar>
      )}
      <div className="ts-kpis">
        <Kpi label="Total this week" value={`${hours(total)} h`} icon="Clock" tint={1} hint={`${trackingDays} working day${trackingDays === 1 ? '' : 's'} · ${config.standardDay} h day`} />
        <Kpi label="Target" value={`${hours(target)} h`} icon="Bullseye" tint={2} tone={total >= target ? 'success' : undefined} hint={total >= target ? 'Complete' : `${hours(target - total)} h to go`} />
        <Kpi label="Billable" value={`${hours(billable)} h`} icon="Money" tint={5} />
        <Kpi label="Non-billable" value={`${hours(total - billable)} h`} icon="Blocked2" tint={3} />
      </div>
      <div>
          <SheetGrid days={days} lines={lines} owner={{ email: me.email, user: myUser }} readOnly={!editable} timeOff={loaded.timeOff} onChange={change} />
          {editable && (
            <div className="ts-row" style={{ marginTop: 10 }}>
              <DefaultButton iconProps={{ iconName: 'Add' }} text="Add line" onClick={() => change(lines.concat(blankLine(me.email, week)))} />
              {!projectsFor(ref, me.email, week).length && <span className="ts-muted ts-small">You are not on any open project yet. Ask a project manager to add you.</span>}
              <span className="ts-spacer" />
              <span className="ts-muted ts-small">Type 7.5, 7:30 or 45m · Ctrl+S saves</span>
            </div>
          )}
          {issues.length > 0 && (
            <ul className="ts-issues" aria-label="Problems with this week">
              {issues.map((i, n) => (
                <li key={n} className={i.level === 'error' ? 'is-error' : ''}>
                  <Icon iconName={i.level === 'error' ? 'ErrorBadge' : 'Warning'} />
                  {i.message}
                </li>
              ))}
            </ul>
          )}
        </div>
      <div className="ts-grid ts-grid-2" style={{ marginTop: 16 }}>
        {loaded.timeOff.length > 0 && (
          <Card title="Time off this week" icon="Vacation">
            {loaded.timeOff.map((t, i) => (
              <div key={i} className="ts-small">
                <strong>{t.type}</strong> · {formatDay(t.start)} – {formatDay(t.end)}
              </div>
            ))}
            <div className="ts-small ts-muted" style={{ marginTop: 6 }}>
              From Time off Manager 365. {config.timeOff.autoFill ? 'Put on your timesheet automatically.' : 'Counts towards your minimum hours.'}
            </div>
          </Card>
        )}
        {sheet && sheet.comments.length > 0 && (
          <Card title="History" icon="History">
            <Comments comments={sheet.comments} />
          </Card>
        )}
      </div>
      {cloneOpen && (
        <CloneDialog
          week={week}
          onDismiss={() => setCloneOpen(false)}
          onClone={(cloned, replace) => {
            setCloneOpen(false);
            const keep = replace ? [] : lines.filter((e) => e.projectId || entryTotal(e) > 0);
            change(keep.concat(cloned).length ? keep.concat(cloned) : [blankLine(me.email, week)]);
            toast(`${cloned.length} line${cloned.length === 1 ? '' : 's'} copied. Save or submit to keep them.`, 'success');
          }}
        />
      )}
      {submitOpen && (
        <Dialog hidden={false} onDismiss={() => setSubmitOpen(false)} dialogContentProps={{ type: DialogType.normal, title: `Submit ${formatWeek(week)}?`, subText: `${hours(total)} hours (${hours(billable)} billable) go to your approver.` }} minWidth={460}>
          {(submitIssues || []).length > 0 && (
            <MessageBar messageBarType={MessageBarType.warning} isMultiline>
              {(submitIssues || []).length} warning{(submitIssues || []).length === 1 ? '' : 's'}: {(submitIssues || [])[0].message}
            </MessageBar>
          )}
          <TextField label="Comment for the approver (optional)" multiline rows={3} value={comment} onChange={(_, v) => setComment(v || '')} />
          <DialogFooter>
            <PrimaryButton text="Submit" onClick={() => void submit()} />
            <DefaultButton text="Cancel" onClick={() => setSubmitOpen(false)} />
          </DialogFooter>
        </Dialog>
      )}
      {confirmDialog}
    </div>
  );
}

/** Clone weekly timesheets: copy the lines (and optionally the hours) of an earlier week. */
function CloneDialog(props: { week: string; onDismiss: () => void; onClone: (lines: ITimeEntry[], replace: boolean) => void }): JSX.Element {
  const { svc, me, workflow, ref, toast, config, myUser } = useApp();
  const [sheets, setSheets] = React.useState<ITimesheet[] | undefined>();
  const [source, setSource] = React.useState(addDays(props.week, -7));
  const [withHours, setWithHours] = React.useState(true);
  const [mode, setMode] = React.useState<'add' | 'replace'>('add');
  const [busy, setBusy] = React.useState(false);
  React.useEffect(() => {
    svc
      .sheetsOf(me.email, 30)
      .then((s) => setSheets(s.filter((x) => x.weekStart !== props.week && x.totalHours > 0)))
      .catch(() => setSheets([]));
  }, []);
  const options = [{ key: addDays(props.week, -7), text: `Last week (${formatWeek(addDays(props.week, -7))})` }].concat(
    (sheets || []).filter((s) => s.weekStart !== addDays(props.week, -7)).map((s) => ({ key: s.weekStart, text: `${formatWeek(s.weekStart)} · ${hours(s.totalHours)} h · ${s.status}` }))
  );
  const run = async (): Promise<void> => {
    setBusy(true);
    try {
      const { entries } = await svc.getSheet(me.email, source);
      if (!entries.length) {
        toast('That week has no lines to copy.', 'warning');
        return;
      }
      const open = ref.projects.filter((p) => p.status === 'Active').map((p) => p.id);
      const timeOff = ref.activities.filter((a) => a.isTimeOff).map((a) => a.id);
      const lines = workflow.cloneLines(
        entries.filter((e) => timeOff.indexOf(e.activityId) < 0),
        source,
        props.week,
        withHours,
        open
      );
      // Hours never land on days that are not tracked this week (holidays, days off).
      lines.forEach((l) => Object.keys(l.hours).forEach((d) => !isTrackingDay(config, d, myUser) && delete l.hours[d]));
      props.onClone(lines, mode === 'replace');
    } catch (e) {
      toast((e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };
  return (
    <Dialog hidden={false} onDismiss={props.onDismiss} dialogContentProps={{ type: DialogType.normal, title: 'Clone a week', subText: `Copy projects, tasks and activities into ${formatWeek(props.week)}.` }} minWidth={480}>
      <Dropdown label="Copy from" options={options} selectedKey={source} onChange={(_, o) => o && setSource(String(o.key))} />
      <Toggle label="Copy the hours too" checked={withHours} onChange={(_, v) => setWithHours(!!v)} styles={{ root: { marginTop: 12 } }} />
      <ChoiceGroup
        selectedKey={mode}
        onChange={(_, o) => o && setMode(o.key as 'add' | 'replace')}
        options={[
          { key: 'add', text: 'Add to the lines already here' },
          { key: 'replace', text: 'Replace the lines already here' }
        ]}
      />
      <DialogFooter>
        <PrimaryButton text="Clone" onClick={() => void run()} disabled={busy} />
        <DefaultButton text="Cancel" onClick={props.onDismiss} />
      </DialogFooter>
    </Dialog>
  );
}
