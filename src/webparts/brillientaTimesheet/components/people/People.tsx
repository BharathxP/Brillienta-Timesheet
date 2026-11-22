import * as React from 'react';
import { Checkbox, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, MessageBar, MessageBarType, Panel, PanelType, PrimaryButton, SearchBox, TextField, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { ITsUser, ROLES, Role } from '../../models/types';
import { ROLE_HELP, ROLE_LABELS, reportsOf, sameEmail } from '../../services/access';
import { DAY_SHORT } from '../../services/dates';
import { money } from '../../services/format';
import { planInfo, userLimit } from '../../services/licensing/plans';
import { Card, Empty, Kpi, Locked, PageHeader, Person, RoleBadge, useBusy } from '../common/Ui';
import { ExportButton, PersonPicker } from '../common/Pickers';

function blankUser(): ITsUser {
  return { id: 0, name: '', email: '', role: 'member', managerEmail: '', department: '', jobTitle: '', costRate: 0, billRate: 0, external: false, active: true, ruleId: '', workDays: [] };
}

/** People who track time: role, manager, department, user-specific hourly cost and bill rate, external users. */
export function People(): JSX.Element {
  const { ref, reloadRef, has, config, license, graph } = useApp();
  const [q, setQ] = React.useState('');
  const [roleFilter, setRoleFilter] = React.useState('');
  const [showInactive, setShowInactive] = React.useState(false);
  const [editing, setEditing] = React.useState<ITsUser | undefined>();
  const [external, setExternal] = React.useState(false);
  const [groupImport, setGroupImport] = React.useState(false);
  const limit = userLimit(license.plan, license.users);
  const active = ref.users.filter((u) => u.active);
  const n = q.trim().toLowerCase();
  const list = ref.users.filter(
    (u) => (showInactive || u.active) && (!roleFilter || u.role === roleFilter) && (!n || u.name.toLowerCase().indexOf(n) >= 0 || u.email.toLowerCase().indexOf(n) >= 0 || u.department.toLowerCase().indexOf(n) >= 0)
  );
  const name = (e: string): string => (ref.users.filter((u) => sameEmail(u.email, e))[0] || { name: e }).name;
  const plan = license.plan ? planInfo(license.plan) : undefined;
  return (
    <>
      <PageHeader
        title="People & rates"
        sub="Roles, managers (who approves whom), hourly cost and bill rates."
        actions={
          <>
            <ExportButton
              name="people"
              sheets={() => [
                {
                  name: 'People',
                  columns: ['Name', 'Email', 'Role', 'Manager', 'Department', 'Job title', 'Hourly cost', 'Bill rate', 'External', 'Active'],
                  rows: ref.users.map((u) => [u.name, u.email, ROLE_LABELS[u.role], u.managerEmail, u.department, u.jobTitle, u.costRate, u.billRate, u.external, u.active])
                }
              ]}
            />
            {graph && <DefaultButton iconProps={{ iconName: 'Group' }} text="Add a group" onClick={() => setGroupImport(true)} />}
            <DefaultButton iconProps={{ iconName: 'AddFriend' }} text="Add external user" onClick={() => setExternal(true)} />
            <PrimaryButton iconProps={{ iconName: 'Add' }} text="Add person" onClick={() => setEditing(blankUser())} />
          </>
        }
      />
      <div className="ts-kpis">
        <Kpi label="Active users" value={`${active.length}${limit ? ' / ' + limit : ''}`} icon="People" tint={1} tone={limit && active.length > limit ? 'danger' : undefined} hint={plan && plan.extraUserPrice ? `Extra users ₹${plan.extraUserPrice}/user` : undefined} />
        <Kpi label="Approvers" value={ref.users.filter((u) => u.active && (u.role !== 'member' || reportsOf(ref.users, u.email, true).length > 0)).length} icon="DocumentApproval" tint={2} />
        <Kpi label="External users" value={active.filter((u) => u.external).length} icon="Globe" tint={3} />
        <Kpi label="Without a cost rate" value={active.filter((u) => !u.costRate).length} icon="Money" tint={5} tone={has('userCost') && active.some((u) => !u.costRate) ? 'warning' : undefined} />
      </div>
      {limit > 0 && active.length > limit && (
        <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 12, borderRadius: 8 } }}>
          {active.length - limit} people beyond the {limit} users of your licence cannot submit timesheets (the most recently added). Deactivate people who no longer track time, or ask Brillienta for more users.
        </MessageBar>
      )}
      <div className="ts-filters">
        <SearchBox placeholder="Search people" value={q} onChange={(_, v) => setQ(v || '')} styles={{ root: { width: 240 } }} />
        <Dropdown selectedKey={roleFilter} onChange={(_, o) => o && setRoleFilter(String(o.key))} options={[{ key: '', text: 'All roles' }].concat(ROLES.map((r) => ({ key: r, text: ROLE_LABELS[r] })))} />
        <Checkbox label="Show inactive" checked={showInactive} onChange={(_, v) => setShowInactive(!!v)} styles={{ root: { alignSelf: 'center' } }} />
      </div>
      {!list.length ? (
        <Card>
          <Empty icon="People" title="Nobody here yet">
            People are added automatically the first time they open the timesheet, or add them here.
          </Empty>
        </Card>
      ) : (
        <div className="ts-table-wrap">
          <table className="ts-table">
            <thead>
              <tr>
                <th>Person</th>
                <th>Role</th>
                <th>Manager</th>
                <th>Department</th>
                {has('userCost') && <th style={{ textAlign: 'right' }}>Hourly cost</th>}
                {has('userRates') && <th style={{ textAlign: 'right' }}>Bill rate</th>}
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {list.map((u) => (
                <tr key={u.id} onClick={() => setEditing(u)}>
                  <td>
                    <Person name={u.name} email={u.email} sub={u.jobTitle} />
                    {u.external && <span className="ts-chip is-soft" style={{ marginLeft: 6 }}>External</span>}
                  </td>
                  <td>
                    <RoleBadge role={u.role} />
                  </td>
                  <td className="ts-small">{u.managerEmail ? name(u.managerEmail) : <span className="ts-muted">—</span>}</td>
                  <td className="ts-small">{u.department || <span className="ts-muted">—</span>}</td>
                  {has('userCost') && <td className="ts-number" style={{ textAlign: 'right' }}>{u.costRate ? money(u.costRate, config.currency) : '—'}</td>}
                  {has('userRates') && <td className="ts-number" style={{ textAlign: 'right' }}>{u.billRate ? money(u.billRate, config.currency) : '—'}</td>}
                  <td>{u.active ? <span className="ts-status is-approved">Active</span> : <span className="ts-status">Inactive</span>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {editing && (
        <UserEditor
          user={editing}
          onDismiss={() => setEditing(undefined)}
          onSaved={() => {
            setEditing(undefined);
            void reloadRef();
          }}
        />
      )}
      {external && (
        <ExternalDialog
          onDismiss={() => setExternal(false)}
          onDone={() => {
            setExternal(false);
            void reloadRef();
          }}
        />
      )}
      {groupImport && (
        <GroupImport
          onDismiss={() => setGroupImport(false)}
          onDone={() => {
            setGroupImport(false);
            void reloadRef();
          }}
        />
      )}
    </>
  );
}

function UserEditor(props: { user: ITsUser; onDismiss: () => void; onSaved: () => void }): JSX.Element {
  const { svc, ref, has, config, graph, toast } = useApp();
  const [u, setU] = React.useState<ITsUser>(props.user);
  const [busy, run] = useBusy();
  const set = (patch: Partial<ITsUser>): void => setU({ ...u, ...patch });
  const roles = ROLES.filter((r) => r !== 'programManager' || has('programManager'));
  const save = (): Promise<void> =>
    run(async () => {
      if (!u.email.trim() || !u.name.trim()) throw new Error('A person needs a name and an email.');
      if (!u.id && ref.users.some((x) => sameEmail(x.email, u.email))) throw new Error('This person is already in the list.');
      if (sameEmail(u.managerEmail, u.email)) throw new Error('A person cannot be their own manager.');
      await svc.saveUser(u);
      await svc.log(u.id ? 'Person updated' : 'Person added', `${u.email} · ${ROLE_LABELS[u.role]}`);
      props.onSaved();
    }, 'Saved.');
  const fill = async (email: string, name: string): Promise<void> => {
    const patch: Partial<ITsUser> = { email, name };
    if (graph) {
      const p = await graph.person(email).catch(() => undefined);
      if (p) Object.assign(patch, { name: p.name || name, department: p.department, jobTitle: p.jobTitle, external: p.external, managerEmail: p.managerEmail || u.managerEmail });
    }
    setU({ ...u, ...patch });
    if (!graph) toast('Microsoft 365 details are not available here; fill them in by hand.', 'info');
  };
  return (
    <Panel
      isOpen
      type={PanelType.medium}
      headerText={u.id ? u.name : 'Add a person'}
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
        {!u.id && <PersonPicker label="Find in Microsoft 365" onPick={(p) => void fill(p.email, p.name)} />}
        <TextField label="Name" required value={u.name} onChange={(_, v) => set({ name: v || '' })} />
        <TextField label="Email" required value={u.email} onChange={(_, v) => set({ email: (v || '').trim() })} disabled={!!u.id} />
        <Dropdown label="Role" selectedKey={u.role} options={roles.map((r) => ({ key: r, text: ROLE_LABELS[r] }))} onChange={(_, o) => o && set({ role: o.key as Role })} />
        <div className="ts-muted ts-small">{ROLE_HELP[u.role]}</div>
        <Dropdown
          label="Manager (approves their timesheets)"
          selectedKey={u.managerEmail.toLowerCase()}
          options={[{ key: '', text: '— none —' }].concat(ref.users.filter((x) => x.active && !sameEmail(x.email, u.email)).map((x) => ({ key: x.email.toLowerCase(), text: x.name })))}
          onChange={(_, o) => o && set({ managerEmail: String(o.key) })}
        />
        <div className="ts-form-grid">
          <TextField label="Department" value={u.department} onChange={(_, v) => set({ department: v || '' })} />
          <TextField label="Job title" value={u.jobTitle} onChange={(_, v) => set({ jobTitle: v || '' })} />
          {has('userCost') ? (
            <TextField label={`Hourly cost (${config.currency})`} type="number" value={u.costRate ? String(u.costRate) : ''} onChange={(_, v) => set({ costRate: Number(v) || 0 })} description="What an hour of this person costs." />
          ) : (
            <Locked feature="userCost" />
          )}
          {has('userRates') ? (
            <TextField label={`Bill rate (${config.currency})`} type="number" value={u.billRate ? String(u.billRate) : ''} onChange={(_, v) => set({ billRate: Number(v) || 0 })} description="Default rate this person is billed at." />
          ) : (
            <Locked feature="userRates" />
          )}
        </div>
        {has('hourRules') && config.rulesEnabled && config.rules.length > 0 && (
          <Dropdown label="Hour rule" selectedKey={u.ruleId} options={[{ key: '', text: 'Matched automatically' }].concat(config.rules.map((r) => ({ key: r.id, text: r.name })))} onChange={(_, o) => o && set({ ruleId: String(o.key) })} />
        )}
        {has('trackingDays') && (
          <div>
            <div className="ms-Label">Days this person tracks time</div>
            <div className="ts-row">
              {[0, 1, 2, 3, 4, 5, 6].map((d) => (
                <Checkbox
                  key={d}
                  label={DAY_SHORT[d]}
                  checked={(u.workDays.length ? u.workDays : config.workDays).indexOf(d) >= 0}
                  onChange={(_, v) => {
                    const base = u.workDays.length ? u.workDays : config.workDays;
                    set({ workDays: (v ? base.concat(d) : base.filter((x) => x !== d)).sort() });
                  }}
                />
              ))}
            </div>
            {u.workDays.length > 0 && (
              <button className="ts-link ts-small" onClick={() => set({ workDays: [] })}>
                Use the organisation&apos;s working days
              </button>
            )}
          </div>
        )}
        <Toggle label="External user (Microsoft 365 guest)" checked={u.external} onChange={(_, v) => set({ external: !!v })} disabled={!has('externalUsers') && !u.external} />
        <Toggle label="Active" checked={u.active} onChange={(_, v) => set({ active: !!v })} onText="Tracks time" offText="Inactive (kept in reports)" />
      </div>
    </Panel>
  );
}

/** Add Office 365 external users: invite a guest (Microsoft Graph) and add them to People. */
function ExternalDialog(props: { onDismiss: () => void; onDone: () => void }): JSX.Element {
  const { has, graph, svc, linkBase, ref } = useApp();
  const [email, setEmail] = React.useState('');
  const [name, setName] = React.useState('');
  const [invite, setInvite] = React.useState(true);
  const [busy, run] = useBusy();
  if (!has('externalUsers'))
    return (
      <Dialog hidden={false} onDismiss={props.onDismiss} dialogContentProps={{ title: 'External users' }} minWidth={480}>
        <Locked feature="externalUsers">Microsoft 365 guests (contractors, client staff) can track time. Included in Enterprise, optional on Pro.</Locked>
        <DialogFooter>
          <DefaultButton text="Close" onClick={props.onDismiss} />
        </DialogFooter>
      </Dialog>
    );
  const go = (): Promise<void> =>
    run(async () => {
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Enter a valid email address.');
      if (ref.users.some((u) => sameEmail(u.email, email))) throw new Error('This person is already in People.');
      if (invite && graph) {
        await graph.inviteGuest(email, name || email, linkBase, `You have been invited to log your time in ${'Brillienta Timesheet'}.`);
      }
      await svc.saveUser({ name: name || email, email, role: 'member', external: true, active: true });
      await svc.log('External user added', email);
      props.onDone();
    }, invite ? 'Invitation sent and person added.' : 'Person added.');
  return (
    <Dialog hidden={false} onDismiss={props.onDismiss} dialogContentProps={{ type: DialogType.normal, title: 'Add an external user', subText: 'A Microsoft 365 guest who tracks time (contractor, client staff).' }} minWidth={480}>
      <TextField label="Email" required value={email} onChange={(_, v) => setEmail((v || '').trim())} />
      <TextField label="Name" value={name} onChange={(_, v) => setName(v || '')} />
      <Toggle label="Send a Microsoft 365 guest invitation" checked={invite && !!graph} disabled={!graph} onChange={(_, v) => setInvite(!!v)} />
      <div className="ts-muted ts-small">
        The invitation needs the User.Invite.All permission (or guest-inviter rights). Also share this SharePoint site with the guest, then run Admin › Schema &amp; security › Apply security so they can use the lists.
      </div>
      <DialogFooter>
        <PrimaryButton text="Add" onClick={() => void go()} disabled={busy} />
        <DefaultButton text="Cancel" onClick={props.onDismiss} />
      </DialogFooter>
    </Dialog>
  );
}

/** Add everyone in a Microsoft 365 group or team at once (with their managers and departments). */
function GroupImport(props: { onDismiss: () => void; onDone: () => void }): JSX.Element {
  const { graph, svc, ref, toast } = useApp();
  const [groups, setGroups] = React.useState<Array<{ id: string; name: string }> | undefined>();
  const [group, setGroup] = React.useState('');
  const [busy, run] = useBusy();
  React.useEffect(() => {
    if (graph)
      graph
        .myGroups()
        .then(setGroups)
        .catch((e) => {
          toast('Groups could not be read: ' + (e as Error).message, 'error');
          setGroups([]);
        });
  }, []);
  const go = (): Promise<void> =>
    run(async () => {
      if (!graph) return;
      const members = (await graph.groupMembers(group)).filter((m) => !ref.users.some((u) => sameEmail(u.email, m.email)));
      const withManagers = [];
      for (const m of members.slice(0, 300)) withManagers.push((await graph.person(m.email)) || m);
      await svc.addUsers(withManagers.map((m) => ({ name: m.name, email: m.email, department: m.department, jobTitle: m.jobTitle, managerEmail: m.managerEmail || '', external: m.external, role: 'member', active: true })));
      await svc.log('People added from a group', `${withManagers.length} people`);
      toast(`${withManagers.length} people added.`, 'success');
      props.onDone();
    });
  return (
    <Dialog hidden={false} onDismiss={props.onDismiss} dialogContentProps={{ type: DialogType.normal, title: 'Add a Microsoft 365 group', subText: 'Everyone in the group who is not in People yet is added as a member, with their manager and department.' }} minWidth={480}>
      <Dropdown label="Group or team" placeholder={groups ? 'Choose' : 'Loading…'} options={(groups || []).map((g) => ({ key: g.id, text: g.name }))} selectedKey={group || null} onChange={(_, o) => o && setGroup(String(o.key))} />
      <DialogFooter>
        <PrimaryButton text="Add people" onClick={() => void go()} disabled={busy || !group} />
        <DefaultButton text="Cancel" onClick={props.onDismiss} />
      </DialogFooter>
    </Dialog>
  );
}
