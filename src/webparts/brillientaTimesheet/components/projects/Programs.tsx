import * as React from 'react';
import { DefaultButton, Panel, PanelType, PrimaryButton, TextField } from '@fluentui/react';
import { useApp } from '../AppContext';
import { IProgram } from '../../models/types';
import { PeopleField } from '../common/Pickers';
import { Card, ColorInput, Empty, Field, Locked, PageHeader, useBusy, useConfirm, useLoader } from '../common/Ui';
import { hours } from '../../services/format';
import { sameEmail } from '../../services/access';

/** Programmes group projects; their programme managers manage, approve and report on all of them (Pro). */
export function Programs(): JSX.Element {
  const { ref, has, svc, reloadRef, isAdmin, me, navigate, syncSecurityGroups } = useApp();
  const [editing, setEditing] = React.useState<IProgram | undefined>();
  const [busy, run] = useBusy();
  const [ask, dialog] = useConfirm();
  const used = useLoader(() => svc.hoursByProject(), []);
  if (!has('programManager')) return (
    <>
      <PageHeader title="Programmes" />
      <Locked feature="programManager">Group projects into programmes run by Programme Managers, who manage, approve and report on all of them.</Locked>
    </>
  );
  const list = isAdmin ? ref.programs : ref.programs.filter((g) => g.managerEmails.some((e) => sameEmail(e, me.email)));
  const save = (g: IProgram): Promise<void> =>
    run(async () => {
      if (!g.title.trim()) throw new Error('A programme needs a name.');
      await svc.saveProgram(g);
      await svc.log(g.id ? 'Programme updated' : 'Programme created', g.title);
      // Programme managers need the role to act on it.
      for (const e of g.managerEmails) {
        const u = ref.users.filter((x) => sameEmail(x.email, e))[0];
        if (u && (u.role === 'member' || u.role === 'teamManager' || u.role === 'projectManager')) await svc.saveUser({ id: u.id, role: 'programManager' });
      }
      setEditing(undefined);
      await reloadRef();
      await syncSecurityGroups();
    }, 'Programme saved.');
  return (
    <>
      <PageHeader title="Programmes" sub="Groups of projects with programme managers." actions={isAdmin && <PrimaryButton iconProps={{ iconName: 'Add' }} text="New programme" onClick={() => setEditing({ id: 0, title: '', description: '', managerEmails: [], color: '#8a5cd1' })} />} />
      {!list.length ? (
        <Card>
          <Empty icon="Org" title="No programmes yet" />
        </Card>
      ) : (
        <div className="ts-tiles">
          {list.map((g) => {
            const projects = ref.projects.filter((p) => p.programId === g.id);
            const h = projects.reduce((s, p) => s + ((used.data || {})[p.id] || 0), 0);
            return (
              <div key={g.id} className="ts-tile" style={{ cursor: 'default' }}>
                <span className="ts-tile-icon" style={{ background: g.color || undefined, color: '#fff' }}>
                  <i className="ms-Icon ms-Icon--Org" />
                </span>
                <div className="ts-tile-title">{g.title}</div>
                <div className="ts-tile-desc">{g.description || `${projects.length} projects`}</div>
                <div className="ts-small" style={{ marginTop: 6 }}>
                  {g.managerEmails.map((e) => (ref.users.filter((u) => sameEmail(u.email, e))[0] || { name: e }).name).join(', ') || 'No programme manager'}
                </div>
                <div className="ts-chips" style={{ marginTop: 8 }}>
                  {projects.map((p) => (
                    <button key={p.id} className="ts-chip is-link" onClick={() => navigate(`projects/${p.id}`)}>
                      {p.title}
                    </button>
                  ))}
                </div>
                <div className="ts-tile-link">
                  <span>{hours(h)} h logged</span>
                  {isAdmin && (
                    <span className="ts-row">
                      <button className="ts-link" onClick={() => setEditing(g)}>
                        Edit
                      </button>
                      <button
                        className="ts-link"
                        disabled={busy}
                        onClick={async () => {
                          if (await ask(`Delete ${g.title}?`, 'Its projects stay; they are no longer in a programme.', 'Delete'))
                            await run(async () => {
                              for (const p of projects) await svc.saveProject({ id: p.id, programId: 0 });
                              await svc.deleteProgram(g.id);
                              await reloadRef();
                            }, 'Programme deleted.');
                        }}
                      >
                        Delete
                      </button>
                    </span>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}
      {editing && <ProgramEditor program={editing} onDismiss={() => setEditing(undefined)} onSave={(g) => void save(g)} busy={busy} />}
      {dialog}
    </>
  );
}

function ProgramEditor(props: { program: IProgram; onDismiss: () => void; onSave: (g: IProgram) => void; busy: boolean }): JSX.Element {
  const [g, setG] = React.useState(props.program);
  return (
    <Panel
      isOpen
      type={PanelType.medium}
      headerText={g.id ? 'Edit programme' : 'New programme'}
      onDismiss={props.onDismiss}
      isFooterAtBottom
      onRenderFooterContent={() => (
        <div className="ts-row">
          <PrimaryButton text="Save" onClick={() => props.onSave(g)} disabled={props.busy} />
          <DefaultButton text="Cancel" onClick={props.onDismiss} />
        </div>
      )}
    >
      <div className="ts-stack">
        <TextField label="Programme" required value={g.title} onChange={(_, v) => setG({ ...g, title: v || '' })} />
        <TextField label="Description" multiline rows={3} value={g.description} onChange={(_, v) => setG({ ...g, description: v || '' })} />
        <PeopleField label="Programme managers" value={g.managerEmails} onChange={(v) => setG({ ...g, managerEmails: v })} hint="They get the Programme Manager role." onlyKnown />
        <Field label="Colour">
          <ColorInput value={g.color} onChange={(v) => setG({ ...g, color: v })} />
        </Field>
        <div className="ts-muted ts-small">Add projects to the programme from each project (Edit project › Programme).</div>
      </div>
    </Panel>
  );
}
