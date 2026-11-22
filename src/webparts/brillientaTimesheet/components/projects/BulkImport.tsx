import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, DialogType, MessageBar, MessageBarType, PrimaryButton, TextField } from '@fluentui/react';
import { useApp } from '../AppContext';
import { BULK_COLUMNS, BulkKind, IBulkRow, bulkTemplate, parseBulk } from '../../services/bulk';
import { PALETTE } from '../../schema/defaults';
import { download } from '../../services/format';
import { canManageProject } from '../../services/access';

/**
 * Bulk project / task / activity creation: paste rows from Excel (or a CSV), check them, create them all in one go.
 * Projects and tasks are written in $batch requests of up to 100.
 */
export function BulkImport(props: { kind: BulkKind; onDismiss: () => void; onDone: () => void }): JSX.Element {
  const { svc, ref, toast, me, role } = useApp();
  const [text, setText] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const rows: IBulkRow[] = React.useMemo(() => (text.trim() ? parseBulk(props.kind, text) : []), [text, props.kind]);
  const findProject = (v: string): number => {
    const n = v.trim().toLowerCase();
    const p = ref.projects.filter((x) => x.code.toLowerCase() === n || x.title.toLowerCase() === n)[0];
    return p && canManageProject(role, me.email, p, ref) ? p.id : 0;
  };
  // Cross-row checks the parser cannot do (projects that exist, duplicates).
  const checked = rows.map((r) => {
    const errors = r.errors.slice();
    if (props.kind === 'tasks' && r.values.project && !findProject(String(r.values.project))) errors.push(`No project "${r.values.project}" you manage`);
    if (props.kind === 'projects' && r.values.code && ref.projects.some((p) => p.code && p.code.toLowerCase() === String(r.values.code).toLowerCase())) errors.push(`Code ${r.values.code} already used`);
    if (props.kind === 'activities' && ref.activities.some((a) => a.title.toLowerCase() === String(r.values.title || '').toLowerCase())) errors.push('Already exists');
    return { ...r, errors };
  });
  const good = checked.filter((r) => !r.errors.length);
  const bool = (v: unknown, d: boolean): boolean => (v === '' || v === undefined ? d : !!v);
  const run = async (): Promise<void> => {
    setBusy(true);
    try {
      if (props.kind === 'projects') {
        await svc.addProjects(
          good.map((r, i) => {
            const v = r.values;
            const program = ref.programs.filter((g) => g.title.toLowerCase() === String(v.program || '').toLowerCase())[0];
            const managers = (v.managers as string[]) || [];
            return {
              title: String(v.title),
              code: String(v.code || ''),
              client: String(v.client || ''),
              managerEmails: managers.length ? managers : [me.email],
              memberEmails: (v.members as string[]) || [],
              billable: bool(v.billable, true),
              status: 'Active',
              startDate: String(v.start || ''),
              endDate: String(v.end || ''),
              plannedHours: Number(v.planned) || 0,
              hourlyCost: Number(v.cost) || 0,
              billRate: Number(v.rate) || 0,
              programId: program ? program.id : 0,
              color: PALETTE[(ref.projects.length + i) % PALETTE.length],
              memberRates: {},
              custom: {},
              endNotices: [],
              description: ''
            };
          })
        );
      } else if (props.kind === 'tasks') {
        await svc.addTasks(
          good.map((r) => ({
            title: String(r.values.title),
            projectId: findProject(String(r.values.project)),
            assigneeEmails: (r.values.assignees as string[]) || [],
            billable: bool(r.values.billable, true),
            plannedHours: Number(r.values.planned) || 0,
            dueDate: String(r.values.due || ''),
            selfAssign: bool(r.values.selfAssign, true),
            status: 'Open',
            description: '',
            plannerTaskId: '',
            plannerPlanId: ''
          }))
        );
      } else {
        await svc.addActivities(good.map((r, i) => ({ title: String(r.values.title), billable: bool(r.values.billable, true), active: true, isTimeOff: false, color: PALETTE[(ref.activities.length + i) % PALETTE.length] })));
      }
      await svc.log(`Bulk ${props.kind} created`, `${good.length} rows`);
      toast(`${good.length} ${props.kind} created.`, 'success');
      props.onDone();
    } catch (e) {
      toast((e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };
  const cols = BULK_COLUMNS[props.kind];
  return (
    <Dialog hidden={false} onDismiss={props.onDismiss} dialogContentProps={{ type: DialogType.largeHeader, title: `Bulk create ${props.kind}` }} minWidth={760} maxWidth={960} modalProps={{ isBlocking: true }}>
      <p className="ts-muted" style={{ marginTop: 0 }}>
        Paste rows from Excel (one {props.kind === 'activities' ? 'activity' : props.kind.slice(0, -1)} per row) or a CSV. Columns, in this order or with a header row:{' '}
        {cols.map((c, i) => (
          <span key={c.key}>
            <strong>{c.label}</strong>
            {c.required ? '*' : ''}
            {c.hint ? <span className="ts-small"> ({c.hint})</span> : null}
            {i < cols.length - 1 ? ', ' : ''}
          </span>
        ))}
        .
      </p>
      <DefaultButton text="Download a template" iconProps={{ iconName: 'Download' }} onClick={() => download(`${props.kind}-template.csv`, bulkTemplate(props.kind).replace(/\t/g, ',') + '\r\n', 'text/csv')} />
      <TextField multiline rows={8} value={text} onChange={(_, v) => setText(v || '')} placeholder={bulkTemplate(props.kind)} styles={{ root: { marginTop: 12 }, field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }} />
      {checked.length > 0 && (
        <>
          <MessageBar messageBarType={good.length === checked.length ? MessageBarType.success : MessageBarType.warning} styles={{ root: { marginTop: 10 } }}>
            {good.length} of {checked.length} rows ready.{good.length < checked.length ? ' Rows with problems are skipped.' : ''}
          </MessageBar>
          <div className="ts-table-wrap" style={{ maxHeight: 240, overflowY: 'auto', marginTop: 8 }}>
            <table className="ts-table">
              <thead>
                <tr>
                  <th>#</th>
                  {cols.slice(0, 5).map((c) => (
                    <th key={c.key}>{c.label}</th>
                  ))}
                  <th>Check</th>
                </tr>
              </thead>
              <tbody>
                {checked.map((r) => (
                  <tr key={r.line} style={{ cursor: 'default' }}>
                    <td className="ts-number">{r.line}</td>
                    {cols.slice(0, 5).map((c) => {
                      const v = r.values[c.key];
                      return <td key={c.key}>{Array.isArray(v) ? v.join('; ') : v === true ? 'Yes' : v === false ? 'No' : String(v === undefined ? '' : v)}</td>;
                    })}
                    <td className={r.errors.length ? 'ts-status-bad ts-small' : 'ts-status-ok ts-small'}>{r.errors.length ? r.errors.join('; ') : 'OK'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
      <DialogFooter>
        <PrimaryButton text={`Create ${good.length || ''} ${props.kind}`} onClick={() => void run()} disabled={busy || !good.length} />
        <DefaultButton text="Cancel" onClick={props.onDismiss} />
      </DialogFooter>
    </Dialog>
  );
}
