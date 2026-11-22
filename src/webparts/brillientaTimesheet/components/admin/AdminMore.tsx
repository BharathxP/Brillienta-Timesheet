import * as React from 'react';
import { Checkbox, ChoiceGroup, DefaultButton, Dropdown, MessageBar, MessageBarType, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { useApp } from '../AppContext';
import { TemplateKey } from '../../models/types';
import { TEMPLATE_LABELS, defaultTemplates } from '../../schema/defaults';
import { DAY_NAMES } from '../../services/dates';
import { download, formatDateTime, relativeTime } from '../../services/format';
import { WEBHOOK_EVENTS } from '../../services/Notifier';
import { flowPackageZip } from '../../services/flowPackage';
import { copilotManifest, declarativeAgent, endpoints, pbids, teamsManifest } from '../../services/packages';
import { zipFiles } from '../../services/zip';
import { clockTileSvg } from '../../assets/defaultLogo';
import { LogLevel, Provisioner, ISchemaReport } from '../../services/Provisioner';
import { SETTING_KEYS } from '../../services/SettingsStore';
import { describeStatus } from '../../services/licensing/license';
import { planHas, planInfo, userLimit } from '../../services/licensing/plans';
import { listTitle } from '../../schema/schema';
import { Card, Kpi, Loading, Tabs, useBusy, useConfirm, useLoader } from '../common/Ui';
import { Gate, SaveBar, useDraft } from './AdminPage';
import { KeyEntry, PlanCards } from './LicenseAdmin';

const WEB_PART_ID = 'ff5b1593-a3b3-45f5-b3a5-8d11821fc2a1';

/** Render an SVG to PNG bytes (Teams / Copilot package icons). */
async function svgToPng(svg: string, size: number): Promise<Uint8Array> {
  const img = new Image();
  img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
  await new Promise<void>((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error('Icon could not be drawn'));
  });
  const c = document.createElement('canvas');
  c.width = c.height = size;
  const ctx = c.getContext('2d');
  if (!ctx) throw new Error('No canvas');
  ctx.drawImage(img, 0, 0, size, size);
  const blob: Blob = await new Promise((resolve, reject) => c.toBlob((b) => (b ? resolve(b) : reject(new Error('PNG failed'))), 'image/png'));
  return new Uint8Array(await blob.arrayBuffer());
}

async function appIcons(): Promise<Array<{ path: string; content: Uint8Array }>> {
  const outline = clockTileSvg('transparent', 'transparent', '#ffffff', '#ffffff').replace('fill="transparent"/><circle cx="60" cy="62" r="34" fill="transparent"', 'fill="transparent"/><circle cx="60" cy="62" r="34" fill="none" stroke="#ffffff" stroke-width="7"');
  return [
    { path: 'color.png', content: await svgToPng(clockTileSvg(), 192) },
    { path: 'outline.png', content: await svgToPng(outline, 32) }
  ];
}

function guid(): string {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (ch) => {
    const r = Math.floor(Math.random() * 16);
    return (ch === 'x' ? r : (r & 0x3) | 0x8).toString(16);
  });
}

// ---------------------------------------------------------------------------------------------- notifications

export function NotificationsAdmin(): JSX.Element {
  const draft = useDraft();
  const { d, set } = draft;
  const { has, notifier, svc, sp, prefix, reminders, ref, me, toast } = useApp();
  const [tpl, setTpl] = React.useState<TemplateKey>('sheetSubmitted');
  const [busy, run] = useBusy();
  const queue = useLoader(() => svc.emailQueue(20), []);
  const r = d.reminders;
  const t = d.templates[tpl];
  const html = React.useMemo(() => {
    const vars: Record<string, string> = { recipient: 'Priya', person: 'Arjun Mehta', approver: 'Priya Nair', week: '16 – 22 Nov 2026', hours: '40', billable: '34', comment: 'Thanks!', project: 'Website redesign', days: '7', endDate: '30 Nov 2026', used: '312', planned: '400', task: 'Design review', subject: 'Question', email: 'arjun@contoso.com', priority: 'Normal', category: 'Question', details: '…' };
    try {
      return notifier.compose(tpl, vars, 'timesheet', 'Open', d).html;
    } catch {
      return '';
    }
  }, [tpl, d, notifier]);
  const flow = (): Promise<void> =>
    run(async () => {
      const list = await sp.get<{ Id: string }>(`${sp.list(listTitle(prefix, 'emailQueue'))}?$select=Id`);
      download('Brillienta-Timesheet-email-flow.zip', flowPackageZip({ siteUrl: sp.webUrl, listId: list.Id, listTitle: listTitle(prefix, 'emailQueue'), displayName: 'Brillienta Timesheet – send emails', sharedMailbox: '', allowedDomains: [], useReplyTo: false }), 'application/zip');
    });
  return (
    <div className="ts-stack">
      <Card title="Pending weekly timesheet reminders" icon="Ringer">
        <Gate feature="reminders" text="Remind people whose week is not submitted (Pro).">
          <Toggle label="Send reminders" checked={r.enabled} onChange={(_, v) => set({ reminders: { ...r, enabled: !!v } })} />
          <div className="ts-form-grid">
            <Dropdown label="For the past week, on" selectedKey={r.day} options={DAY_NAMES.map((n, i) => ({ key: i, text: n }))} onChange={(_, o) => o && set({ reminders: { ...r, day: Number(o.key) } })} />
            <Dropdown label="From" selectedKey={r.hour} options={Array.from({ length: 24 }, (_, i) => ({ key: i, text: `${i}:00` }))} onChange={(_, o) => o && set({ reminders: { ...r, hour: Number(o.key) } })} />
            <Dropdown label="Chase weeks up to" selectedKey={r.lookbackWeeks} options={[1, 2, 3, 4, 8].map((n) => ({ key: n, text: `${n} week${n === 1 ? '' : 's'} back` }))} onChange={(_, o) => o && set({ reminders: { ...r, lookbackWeeks: Number(o.key) } })} />
            <div className="ts-stack" style={{ gap: 6, paddingTop: 28 }}>
              <Checkbox label="Also on the last working day of the current week" checked={r.currentWeek} onChange={(_, v) => set({ reminders: { ...r, currentWeek: !!v } })} />
              <Checkbox label="Copy the person's manager" checked={r.ccManager} onChange={(_, v) => set({ reminders: { ...r, ccManager: !!v } })} />
            </div>
          </div>
          <div className="ts-row" style={{ marginTop: 10 }}>
            <DefaultButton
              text="Send due reminders now"
              disabled={busy}
              onClick={() =>
                void run(async () => {
                  const res = await reminders.run({ cfg: d, users: ref.users, projects: ref.projects, actor: me.title, remindersOn: has('reminders'), endDatesOn: has('projectEndDates'), force: true });
                  toast(`${res ? res.reminders : 0} reminders and ${res ? res.notices : 0} project notices sent.`, 'success');
                })
              }
            />
            <span className="ts-muted ts-small">Reminders go out when an administrator opens the app (at most hourly). For unattended sending, schedule the Power Automate flow below.</span>
          </div>
        </Gate>
      </Card>
      <Card title="Project end dates and planned hours" icon="Calendar">
        <Gate feature="projectEndDates">
          <TextField label="Tell project managers this many days before the end date" value={d.projectEndNotice.join(', ')} onChange={(_, v) => set({ projectEndNotice: (v || '').split(/[,\s]+/).map(Number).filter((x) => x > 0) })} description="e.g. 14, 7, 1 — and on the end date itself." />
        </Gate>
        <Gate feature="plannedHours">
          <Toggle label="Tell project managers when logged hours pass the planned hours" checked={d.plannedHoursAlerts} onChange={(_, v) => set({ plannedHoursAlerts: !!v })} styles={{ root: { marginTop: 10 } }} />
        </Gate>
      </Card>
      <Card title="How email is sent" icon="Mail">
        <ChoiceGroup
          selectedKey={has('teamsOutlook') ? d.integrations.mailMode : 'flow'}
          options={[
            { key: 'graph', text: 'Outlook — from the mailbox of the person who acts (Microsoft Graph Mail.Send)', disabled: !has('teamsOutlook') },
            { key: 'flow', text: 'Power Automate — the email queue list, sent by a flow you import (one sender)' }
          ]}
          onChange={(_, o) => o && set({ integrations: { ...d.integrations, mailMode: o.key as 'graph' | 'flow' } })}
        />
        <Gate feature="teamsOutlook" text="Teams chat messages and Outlook email (Enterprise).">
          <Toggle label="Also post notifications to people in Teams chat" checked={d.integrations.teamsChat} onChange={(_, v) => set({ integrations: { ...d.integrations, teamsChat: !!v } })} styles={{ root: { marginTop: 8 } }} />
        </Gate>
        <div className="ts-row" style={{ marginTop: 8 }}>
          <DefaultButton iconProps={{ iconName: 'Download' }} text="Download the email flow" onClick={() => void flow()} disabled={busy} />
          <span className="ts-muted ts-small">Power Automate › My flows › Import › Import Package (Legacy). Choose your SharePoint and Outlook connections.</span>
        </div>
        {queue.data && queue.data.length > 0 && (
          <div className="ts-table-wrap" style={{ marginTop: 12, maxHeight: 220, overflowY: 'auto' }}>
            <table className="ts-table">
              <tbody>
                {queue.data.map((q) => (
                  <tr key={q.id} style={{ cursor: 'default' }}>
                    <td className="ts-small">{q.subject}</td>
                    <td className="ts-small ts-muted">{q.to}</td>
                    <td className="ts-small">{q.status}</td>
                    <td className="ts-small ts-muted">{relativeTime(q.created)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
      <Card title="Email templates" icon="MailOptions">
        <Gate feature="emailTemplates" text="Edit the subject and text of every email and Teams message (Enterprise). The standard templates are used meanwhile.">
          <div className="ts-tpl-head">
            <Dropdown label="Template" selectedKey={tpl} options={(Object.keys(TEMPLATE_LABELS) as TemplateKey[]).map((k) => ({ key: k, text: TEMPLATE_LABELS[k].label }))} onChange={(_, o) => o && setTpl(o.key as TemplateKey)} styles={{ root: { width: 280 } }} />
            <div className="ts-tpl-switches">
              <Toggle label="Send" checked={t.enabled} onChange={(_, v) => set({ templates: { ...d.templates, [tpl]: { ...t, enabled: !!v } } })} />
              <Toggle label="Also in Teams" checked={t.teams} onChange={(_, v) => set({ templates: { ...d.templates, [tpl]: { ...t, teams: !!v } } })} />
            </div>
          </div>
          <div className="ts-muted ts-small" style={{ margin: '6px 0' }}>
            {TEMPLATE_LABELS[tpl].when} Placeholders:{' '}
            {TEMPLATE_LABELS[tpl].placeholders.map((p) => (
              <span key={p} className="ts-ph">{`{{${p}}}`}</span>
            ))}
          </div>
          <div className="ts-tpl-grid">
            <div className="ts-stack">
              <TextField label="Subject" value={t.subject} onChange={(_, v) => set({ templates: { ...d.templates, [tpl]: { ...t, subject: v || '' } } })} />
              <TextField label="Body (HTML)" multiline rows={12} value={t.body} onChange={(_, v) => set({ templates: { ...d.templates, [tpl]: { ...t, body: v || '' } } })} styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }} />
              <div>
                <DefaultButton text="Restore the standard text" onClick={() => set({ templates: { ...d.templates, [tpl]: defaultTemplates()[tpl] } })} />
              </div>
            </div>
            <iframe className="ts-tpl-preview" title="Preview" sandbox="" srcDoc={html} />
          </div>
        </Gate>
      </Card>
      <SaveBar draft={draft} what="Notifications" />
    </div>
  );
}

// ---------------------------------------------------------------------------------------------- integrations

type ITab = 'timeoff' | 'teams' | 'planner' | 'copilot' | 'automate' | 'api' | 'powerbi';

export function IntegrationsAdmin(): JSX.Element {
  const draft = useDraft();
  const { d, set } = draft;
  const { has, timeOff, sp, prefix, appVersion, config, toast, notifier } = useApp();
  const [tab, setTab] = React.useState<ITab>('timeoff');
  const [busy, run] = useBusy();
  const [test, setTest] = React.useState<string[] | undefined>();
  const to = d.timeOff;
  const i = d.integrations;
  const host = new URL(sp.webUrl).hostname;
  const copilotPkg = (): Promise<void> =>
    run(async () => {
      const o = { siteUrl: sp.webUrl, orgName: config.orgName, appName: config.appName, prefix, appId: guid(), version: appVersion.split('.').slice(0, 3).join('.') };
      const zip = zipFiles([{ path: 'manifest.json', content: JSON.stringify(copilotManifest(o), null, 2) }, { path: 'declarativeAgent.json', content: JSON.stringify(declarativeAgent(o), null, 2) }, ...(await appIcons())]);
      download('Brillienta-Timesheet-Copilot-agent.zip', zip, 'application/zip');
    });
  const teamsPkg = (): Promise<void> =>
    run(async () => {
      const zip = zipFiles([{ path: 'manifest.json', content: JSON.stringify(teamsManifest({ appId: guid(), webPartId: WEB_PART_ID, version: appVersion.split('.').slice(0, 3).join('.'), tenantHost: host, appName: config.appName }), null, 2) }, ...(await appIcons())]);
      download('Brillienta-Timesheet-Teams-app.zip', zip, 'application/zip');
    });
  return (
    <div className="ts-stack">
      <Tabs
        value={tab}
        onChange={setTab}
        tabs={[
          { key: 'timeoff', text: 'Time off Manager 365', icon: 'Vacation', locked: !has('timeOff') },
          { key: 'teams', text: 'Teams & Outlook', icon: 'TeamsLogo', locked: !has('teamsOutlook') },
          { key: 'planner', text: 'Planner', icon: 'PlannerLogo', locked: !has('planner') },
          { key: 'copilot', text: 'Copilot', icon: 'Robot', locked: !has('copilot') },
          { key: 'automate', text: 'Power Automate', icon: 'Flow', locked: !has('powerAutomate') },
          { key: 'api', text: 'API', icon: 'Code', locked: !has('api') },
          { key: 'powerbi', text: 'Power BI', icon: 'PowerBILogo', locked: !has('powerBi') }
        ]}
      />
      {tab === 'timeoff' && (
        <Card title="Time off Manager 365" icon="Vacation">
          <Gate feature="timeOff">
            <p className="ts-muted" style={{ marginTop: 0 }}>
              Approved leave from Time off Manager 365 (or any SharePoint leave list) shows on the timesheet, counts towards minimum hours and can be filled in automatically on the Time off activity.
            </p>
            <Toggle label="Read approved leave" checked={to.enabled} onChange={(_, v) => set({ timeOff: { ...to, enabled: !!v } })} />
            <div className="ts-form-grid">
              <TextField label="Site of Time off Manager 365" placeholder="Empty = this site" value={to.siteUrl} onChange={(_, v) => set({ timeOff: { ...to, siteUrl: (v || '').trim() } })} className="is-wide" />
              <TextField label="Leave list" value={to.listTitle} onChange={(_, v) => set({ timeOff: { ...to, listTitle: v || '' } })} />
              <TextField label="Employee email column" value={to.emailField} onChange={(_, v) => set({ timeOff: { ...to, emailField: v || '' } })} />
              <TextField label="Start date column" value={to.startField} onChange={(_, v) => set({ timeOff: { ...to, startField: v || '' } })} />
              <TextField label="End date column" value={to.endField} onChange={(_, v) => set({ timeOff: { ...to, endField: v || '' } })} />
              <TextField label="Status column" value={to.statusField} onChange={(_, v) => set({ timeOff: { ...to, statusField: v || '' } })} />
              <TextField label="Approved value" value={to.approvedValue} onChange={(_, v) => set({ timeOff: { ...to, approvedValue: v || '' } })} />
              <TextField label="Leave type column" value={to.typeField} onChange={(_, v) => set({ timeOff: { ...to, typeField: v || '' } })} />
              <Toggle label="Put approved leave on the timesheet automatically" checked={to.autoFill} onChange={(_, v) => set({ timeOff: { ...to, autoFill: !!v } })} />
            </div>
            <div className="ts-row" style={{ marginTop: 10 }}>
              <DefaultButton text="Test the connection" disabled={busy} onClick={() => void run(async () => setTest(await timeOff.test(to)))} />
              {test && (test.length ? <span className="ts-status-bad ts-small">{test.join(' ')}</span> : <span className="ts-status-ok ts-small">The leave list and its columns were found.</span>)}
            </div>
            <div className="ts-muted ts-small" style={{ marginTop: 6 }}>Column names are internal names (List settings › column › Field= in the address).</div>
          </Gate>
        </Card>
      )}
      {tab === 'teams' && (
        <Card title="Teams & Outlook" icon="TeamsLogo">
          <Gate feature="teamsOutlook">
            <div className="ts-check-list">
              <div>
                <strong>Teams app.</strong> In the SharePoint app catalog select the solution and choose <em>Sync to Teams</em>, or upload this package in Teams admin center › Manage apps. People get Timesheet as a personal app; teams can add it as a tab.
              </div>
              <div>
                <DefaultButton iconProps={{ iconName: 'Download' }} text="Download the Teams app package" onClick={() => void teamsPkg()} disabled={busy} />
              </div>
              <div>
                <strong>Outlook email and Teams chat</strong> are switched on under Notifications &amp; email. They use the Mail.Send, Chat.Create and ChatMessage.Send permissions the package requests: approve them in the SharePoint admin center › Advanced › API access.
              </div>
              <div>
                <strong>Teams tabs on another site:</strong> a tab finds the timesheet through the web part&apos;s &quot;Timesheet site address&quot; setting.
              </div>
            </div>
          </Gate>
        </Card>
      )}
      {tab === 'planner' && (
        <Card title="Microsoft Planner" icon="PlannerLogo">
          <Gate feature="planner">
            <Toggle label="Import tasks from Planner plans (Project › Tasks › Import from Planner)" checked={i.plannerEnabled} onChange={(_, v) => set({ integrations: { ...i, plannerEnabled: !!v } })} />
            <Toggle label="When a linked task is marked done (or reopened), update it in Planner" checked={i.plannerWriteBack} onChange={(_, v) => set({ integrations: { ...i, plannerWriteBack: !!v } })} />
            <div className="ts-muted ts-small">Needs Tasks.ReadWrite and Group.Read.All approved in API access. Assignees are matched by email; re-importing updates linked tasks.</div>
          </Gate>
        </Card>
      )}
      {tab === 'copilot' && (
        <Card title="Microsoft 365 Copilot" icon="Robot">
          <Gate feature="copilot">
            <p style={{ marginTop: 0 }}>A Copilot agent that answers questions about projects, tasks, timesheets and approvals from this site — “How many hours did my team log on Website redesign last month?”</p>
            <Toggle label="Copilot integration on" checked={i.copilotEnabled} onChange={(_, v) => set({ integrations: { ...i, copilotEnabled: !!v } })} />
            <ol className="ts-small">
              <li>Download the agent package (it is grounded on {sp.webUrl}).</li>
              <li>Teams admin center › Manage apps › Upload new app (or Teams › Apps › Upload a custom app).</li>
              <li>People with a Microsoft 365 Copilot licence find it in Copilot Chat. It only sees what each person may see in SharePoint.</li>
            </ol>
            <DefaultButton iconProps={{ iconName: 'Download' }} text="Download the Copilot agent" onClick={() => void copilotPkg()} disabled={busy || !i.copilotEnabled} />
          </Gate>
        </Card>
      )}
      {tab === 'automate' && (
        <Card title="Customise with Power Automate" icon="Flow">
          <Gate feature="powerAutomate">
            <p className="ts-muted" style={{ marginTop: 0 }}>
              Two ways: build flows on the lists themselves (SharePoint trigger “When an item is created or modified” on {listTitle(prefix, 'timesheets')}), or receive events from the app with the trigger “When an HTTP request is received” and paste its URL here.
            </p>
            <TextField label="Flow HTTP trigger URL" value={i.webhookUrl} onChange={(_, v) => set({ integrations: { ...i, webhookUrl: (v || '').trim() } })} placeholder="https://prod-00.westeurope.logic.azure.com/workflows/…" />
            <div className="ms-Label">Events sent</div>
            <div className="ts-row">
              {WEBHOOK_EVENTS.map((e) => (
                <Checkbox key={e.key} label={e.text} checked={i.webhookEvents.indexOf(e.key) >= 0} onChange={(_, v) => set({ integrations: { ...i, webhookEvents: v ? i.webhookEvents.concat(e.key) : i.webhookEvents.filter((x) => x !== e.key) } })} />
              ))}
            </div>
            <pre className="ts-code" style={{ marginTop: 10 }}>{`{ "event": "timesheet.submitted", "at": "2026-11-22T10:00:00Z", "org": "${config.orgName}",\n  "data": { "sheetId": 12, "person": "name@contoso.com", "weekStart": "2026-11-16", "hours": 40, "billableHours": 34 } }`}</pre>
            <DefaultButton
              text="Send a test event"
              disabled={!i.webhookUrl}
              onClick={() => {
                notifier.event(i.webhookEvents[0] || 'timesheet.submitted', { test: true });
                toast('Test event sent (save first if you just changed the URL).', 'info');
              }}
            />
          </Gate>
        </Card>
      )}
      {tab === 'api' && (
        <Card title="API connectivity" icon="Code">
          <Gate feature="api">
            <p className="ts-muted" style={{ marginTop: 0 }}>
              Every record is in a SharePoint list, so any system can read and write it through the SharePoint REST API or Microsoft Graph (<code>/sites/{'{site-id}'}/lists/{'{list}'}/items?expand=fields</code>) with an Entra ID app registration
              (Sites.Selected recommended). People are stored by email; dates as yyyy-mm-dd; hours per day as JSON.
            </p>
            <div className="ts-table-wrap">
              <table className="ts-table">
                <thead>
                  <tr>
                    <th>List</th>
                    <th>REST endpoint</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {endpoints(sp.webUrl, prefix).map((e) => (
                    <tr key={e.key} style={{ cursor: 'default' }}>
                      <td>
                        <div className="ts-subject">{e.title}</div>
                        <div className="ts-muted ts-small">{e.description}</div>
                      </td>
                      <td className="ts-small" style={{ wordBreak: 'break-all' }}>
                        <code>{e.rest}</code>
                      </td>
                      <td>
                        <DefaultButton
                          text="Copy"
                          onClick={() =>
                            void navigator.clipboard
                              .writeText(e.rest)
                              .then(() => toast('Copied.', 'success'))
                              .catch(() => undefined)
                          }
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Gate>
        </Card>
      )}
      {tab === 'powerbi' && (
        <Card title="Connect with Power BI" icon="PowerBILogo">
          <Gate feature="powerBi">
            <ol className="ts-small">
              <li>Download the data source file and open it with Power BI Desktop (it connects to {sp.webUrl} with the SharePoint Online list connector).</li>
              <li>Choose {listTitle(prefix, 'entries')}, {listTitle(prefix, 'timesheets')}, {listTitle(prefix, 'projects')}, {listTitle(prefix, 'tasks')}, {listTitle(prefix, 'activities')} and {listTitle(prefix, 'users')}.</li>
              <li>Relate TimeEntries.ProjectId → Projects.ID, TaskId → Tasks.ID, ActivityId → Activities.ID, SheetId → Timesheets.ID, UserEmail → Users.Email.</li>
              <li>Expand DayHours with Transform › Parse › JSON, then Record › To Table, to get one row per day (Date, Hours).</li>
              <li>Publish and schedule a refresh in the Power BI service.</li>
            </ol>
            <DefaultButton iconProps={{ iconName: 'Download' }} text="Download Power BI data source (.pbids)" onClick={() => download('Brillienta-Timesheet.pbids', pbids(sp.webUrl), 'application/json')} />
          </Gate>
        </Card>
      )}
      <SaveBar draft={draft} what="Integrations" />
    </div>
  );
}

// ---------------------------------------------------------------------------------------------- licence

export function LicenceAdmin(): JSX.Element {
  const { license, tenantId, settings, refreshLicense, ref, toast } = useApp();
  const limit = userLimit(license.plan, license.users);
  const active = ref.users.filter((u) => u.active).length;
  const plan = license.plan ? planInfo(license.plan) : undefined;
  return (
    <div className="ts-stack">
      <Card title="Licence" icon="Certificate">
        <div className="ts-kpis">
          <Kpi label="Status" value={license.state === 'trial' ? 'Trial' : plan ? plan.name : 'None'} icon="Certificate" tint={1} hint={describeStatus(license)} />
          <Kpi label="Users" value={`${active} / ${limit || '∞'}`} icon="People" tint={2} tone={limit && active > limit ? 'danger' : undefined} />
          <Kpi label="Extra users" value={plan && plan.extraUserPrice ? `₹${plan.extraUserPrice}/user` : '—'} icon="AddFriend" tint={5} />
          <Kpi label="Support" value={plan && plan.tickets ? `${plan.tickets} tickets/year` : 'Unlimited'} icon="Headset" tint={3} />
        </div>
        {license.license && (
          <div className="ts-small ts-muted">
            Licensed to {license.license.customer} · {license.license.id} · issued {license.license.issued}
            {license.license.expires ? ` · until ${license.license.expires}` : ' · perpetual'}
            {license.addons.length ? ` · add-ons: ${license.addons.join(', ')}` : ''}
          </div>
        )}
        {license.keyProblem && <MessageBar messageBarType={MessageBarType.warning}>The stored key was not accepted: {license.keyProblem}</MessageBar>}
        <div style={{ marginTop: 12 }}>
          <KeyEntry
            tenantId={tenantId}
            settings={settings}
            onChanged={() =>
              void refreshLicense().then(() => {
                toast('Licence activated.', 'success');
              })
            }
          />
        </div>
      </Card>
      <PlanCards current={license.state === 'trial' ? undefined : license.plan} />
      {!planHas(license.plan, 'externalUsers', license.addons) && license.plan === 'pro' && <div className="ts-muted ts-small">External users are an option on Pro: ask Brillienta to add them to your licence.</div>}
    </div>
  );
}

// ---------------------------------------------------------------------------------------------- schema & security

export function SecurityAdmin(): JSX.Element {
  const { sp, prefix, ref, admins, tenantId, settings, security, setSecurity, svc, me, toast } = useApp();
  const [log, setLog] = React.useState<Array<{ m: string; l: LogLevel }>>([]);
  const [report, setReport] = React.useState<ISchemaReport | undefined>();
  const [busy, run] = useBusy();
  const [ask, dialog] = useConfirm();
  const audit = useLoader(() => svc.audit(100), []);
  const add = (m: string, l: LogLevel): void => setLog((x) => x.concat({ m, l }));
  const prov = new Provisioner(sp, prefix);
  const managers = ref.users.filter((u) => u.active && u.role !== 'member' && u.role !== 'admin').map((u) => u.email);
  const teamManagers = ref.users.filter((u) => u.active && ref.users.some((x) => x.managerEmail.toLowerCase() === u.email.toLowerCase())).map((u) => u.email);
  return (
    <div className="ts-stack">
      <Card title="Lists" icon="Database">
        <p className="ts-muted" style={{ marginTop: 0 }}>
          Brillienta Timesheet keeps everything in hidden lists prefixed <code>{prefix}</code> on {sp.webUrl}. Check them, or repair missing columns (data is never deleted).
        </p>
        <div className="ts-row">
          <DefaultButton text="Check" disabled={busy} onClick={() => void run(async () => setReport(await prov.ensureSchema(add, false)))} />
          <PrimaryButton text="Repair" disabled={busy} onClick={() => void run(async () => setReport(await prov.ensureSchema(add, true)), 'Lists repaired.')} />
        </div>
        {report && (
          <MessageBar messageBarType={report.ok ? MessageBarType.success : MessageBarType.warning} styles={{ root: { marginTop: 10 } }} isMultiline>
            {report.ok ? 'Every list and column is in place.' : report.lists.filter((l) => !l.exists || l.missingFields.length || l.errors.length).map((l) => `${l.title}: ${!l.exists ? 'missing' : l.missingFields.length ? 'missing ' + l.missingFields.join(', ') : l.errors[0]}`).join(' · ')}
          </MessageBar>
        )}
      </Card>
      <Card title="Security" icon="Lock">
        <p style={{ marginTop: 0 }}>
          Applies the permission model: everyone (and each external user) reads projects and settings and writes only their own timesheets; approvers ({managers.length + teamManagers.length}) see and edit all timesheets and manage projects;
          administrators ({admins.length}) edit everything. Run it again after adding approvers, administrators or external users.
        </p>
        {security && security.appliedAt && (
          <div className="ts-small ts-muted">
            Last applied {formatDateTime(security.appliedAt)} by {security.by}.
          </div>
        )}
        <PrimaryButton
          text="Apply security"
          disabled={busy}
          styles={{ root: { marginTop: 10 } }}
          onClick={() =>
            void run(async () => {
              const problems = await prov.applySecurity({ tenantId, admins: admins.concat(me.email), managers: managers.concat(teamManagers), externals: ref.users.filter((u) => u.external && u.active).map((u) => u.email) }, add);
              const meta = { appliedAt: new Date().toISOString(), by: me.title };
              await settings.setJson(SETTING_KEYS.security, meta);
              setSecurity(meta);
              await svc.log('Security applied', problems.join('; '));
              if (problems.length) toast(`Applied with ${problems.length} problem(s): ${problems[0]}`, 'warning');
              else toast('Security applied.', 'success');
            })
          }
        />
      </Card>
      {log.length > 0 && (
        <div className="ts-log">
          {log.map((l, i) => (
            <div key={i} className={'is-' + l.l}>
              {l.m}
            </div>
          ))}
        </div>
      )}
      <Card title="Audit log" icon="History">
        {audit.loading ? (
          <Loading />
        ) : (
          <div className="ts-table-wrap" style={{ maxHeight: 360, overflowY: 'auto' }}>
            <table className="ts-table">
              <tbody>
                {(audit.data || []).map((a) => (
                  <tr key={a.id} style={{ cursor: 'default' }}>
                    <td className="ts-small ts-muted" style={{ whiteSpace: 'nowrap' }}>{formatDateTime(a.created)}</td>
                    <td className="ts-small">
                      <strong>{a.action}</strong> {a.details}
                    </td>
                    <td className="ts-small">{a.by}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
      <Card title="Remove Brillienta Timesheet" icon="Delete">
        <p className="ts-muted" style={{ marginTop: 0 }}>Deletes every timesheet list and all their data from this site (they go to the site recycle bin).</p>
        <DefaultButton
          text="Delete all timesheet lists"
          disabled={busy}
          onClick={async () => {
            const typed = await ask('Delete all timesheet data?', 'Type DELETE to confirm. This removes every list and its data from this site.', 'Delete everything', 'Type DELETE', true);
            if (typed !== 'DELETE') return;
            await run(async () => {
              await prov.removeAll(add);
              window.location.reload();
            });
          }}
        />
      </Card>
      {dialog}
    </div>
  );
}
