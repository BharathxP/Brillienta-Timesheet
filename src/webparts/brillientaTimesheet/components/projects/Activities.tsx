import * as React from 'react';
import { DefaultButton, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { IActivity } from '../../models/types';
import { Card, ColorInput, Empty, Locked, PageHeader, useBusy, useConfirm } from '../common/Ui';
import { BulkImport } from './BulkImport';
import { PALETTE } from '../../schema/defaults';

/** Activity tracking: the kinds of work people pick on each timesheet line (Pro). */
export function Activities(): JSX.Element {
  const { ref, svc, reloadRef, has, config, saveConfig } = useApp();
  const [name, setName] = React.useState('');
  const [bulk, setBulk] = React.useState(false);
  const [busy, run] = useBusy();
  const [ask, dialog] = useConfirm();
  const on = has('activityTracking');
  const add = (): Promise<void> =>
    run(async () => {
      if (!name.trim()) return;
      if (ref.activities.some((a) => a.title.toLowerCase() === name.trim().toLowerCase())) throw new Error('That activity exists already.');
      await svc.saveActivity({ title: name.trim(), billable: true, active: true, isTimeOff: false, color: PALETTE[ref.activities.length % PALETTE.length] });
      setName('');
      await reloadRef();
    }, 'Activity added.');
  const update = (a: IActivity, patch: Partial<IActivity>): Promise<void> =>
    run(async () => {
      await svc.saveActivity({ id: a.id, ...patch });
      await reloadRef();
    });
  return (
    <>
      <PageHeader
        title="Activities"
        sub="Kinds of work tracked on each timesheet line (development, meetings, travel…), with their default billable setting."
        actions={on && has('bulkActivities') && <DefaultButton iconProps={{ iconName: 'BulkUpload' }} text="Bulk create" onClick={() => setBulk(true)} />}
      />
      {!on && <Locked feature="activityTracking">Track the kind of work on each line and report hours by activity.</Locked>}
      {on && (
        <Card>
          <div className="ts-row" style={{ marginBottom: 14 }}>
            <TextField placeholder="New activity" value={name} onChange={(_, v) => setName(v || '')} onKeyDown={(e) => e.key === 'Enter' && void add()} styles={{ root: { width: 280 } }} />
            <PrimaryButton text="Add" onClick={() => void add()} disabled={busy || !name.trim()} />
            <span className="ts-spacer" />
            <Toggle inlineLabel label="Require an activity on every line" checked={config.requireActivity} onChange={(_, v) => void saveConfig({ ...config, requireActivity: !!v }, 'Require activity')} />
          </div>
          {!ref.activities.length ? (
            <Empty title="No activities" />
          ) : (
            <div className="ts-table-wrap">
              <table className="ts-table">
                <thead>
                  <tr>
                    <th style={{ width: 60 }}>Colour</th>
                    <th>Activity</th>
                    <th>Billable by default</th>
                    <th>Active</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {ref.activities.map((a) => (
                    <tr key={a.id} style={{ cursor: 'default' }}>
                      <td>
                        <ColorInput value={a.color} onChange={(v) => void update(a, { color: v })} />
                      </td>
                      <td>
                        <span className="ts-subject">{a.title}</span>
                        {a.isTimeOff && <span className="ts-chip is-lime" style={{ marginLeft: 8 }}>Time off</span>}
                      </td>
                      <td>
                        <Toggle checked={a.billable} disabled={busy || a.isTimeOff} onChange={(_, v) => void update(a, { billable: !!v })} />
                      </td>
                      <td>
                        <Toggle checked={a.active} disabled={busy} onChange={(_, v) => void update(a, { active: !!v })} />
                      </td>
                      <td style={{ textAlign: 'right' }}>
                        {!a.isTimeOff && (
                          <DefaultButton
                            text="Delete"
                            disabled={busy}
                            onClick={async () => {
                              if (await ask(`Delete ${a.title}?`, 'Lines already logged keep it in reports as "(no activity)". Switch it off instead to keep history readable.', 'Delete'))
                                await run(async () => {
                                  await svc.deleteActivity(a.id);
                                  await reloadRef();
                                }, 'Deleted.');
                            }}
                          />
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </Card>
      )}
      {bulk && (
        <BulkImport
          kind="activities"
          onDismiss={() => setBulk(false)}
          onDone={() => {
            setBulk(false);
            void reloadRef();
          }}
        />
      )}
      {dialog}
    </>
  );
}
