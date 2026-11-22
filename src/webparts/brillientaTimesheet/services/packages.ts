// Files the admin downloads from Admin › Integrations. Pure (tests/packages.test.cjs).
//
// - Copilot: a Microsoft 365 Copilot declarative agent (Teams app package) grounded on the timesheet site, so people
//   can ask Copilot about projects, tasks and timesheets in natural language. Upload it in Teams admin center or
//   Teams › Apps › Manage your apps › Upload an app.
// - Power BI: a .pbids data source file that opens Power BI Desktop connected to the site's SharePoint lists.
// - Teams: a Teams app package with the web part as a personal app and a channel tab (SharePoint-hosted).
// - API: the REST and Microsoft Graph endpoints of every list.

import { LISTS, listTitle } from '../schema/schema';

export interface ICopilotOptions {
  siteUrl: string;
  orgName: string;
  appName: string;
  prefix: string;
  appId: string;
  version: string;
}

export function declarativeAgent(o: ICopilotOptions): Record<string, unknown> {
  const lists = ['projects', 'tasks', 'timesheets', 'entries', 'activities', 'users'].map((k) => listTitle(o.prefix, k as never));
  return {
    $schema: 'https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.2/schema.json',
    version: 'v1.2',
    name: `${o.appName} assistant`,
    description: `Answers questions about ${o.orgName || 'your'} projects, tasks, timesheets and hours from Brillienta Timesheet.`,
    instructions: [
      `You are the ${o.appName} assistant for ${o.orgName || 'this organisation'}. Answer only from the SharePoint site ${o.siteUrl} and its lists: ${lists.join(', ')}.`,
      `${lists[0]} holds projects (code, client, status, start and end dates, planned hours, managers). ${lists[1]} holds tasks (ProjectId, assignees, status, due date, planned hours).`,
      `${lists[2]} holds one row per person per week (UserEmail, WeekStart yyyy-mm-dd, SheetStatus Draft/Submitted/Approved/Rejected, TotalHours, BillableHours).`,
      `${lists[3]} holds timesheet lines (ProjectId, TaskId, ActivityId, IsBillable, DayHours JSON of date to hours, TotalHours).`,
      'When asked for totals, add up TotalHours and say which weeks and statuses you included. Never reveal hourly cost or bill rates unless the user asks about their own.',
      'If something is not in the data, say so. To log or change time, point the user to the Brillienta Timesheet page.'
    ].join('\n'),
    capabilities: [{ name: 'OneDriveAndSharePoint', items_by_url: [{ url: o.siteUrl }] }],
    conversation_starters: [
      { title: 'My week', text: 'How many hours have I logged this week, and is my timesheet submitted?' },
      { title: 'Project status', text: 'Which projects end in the next 30 days and how many of their planned hours are used?' },
      { title: 'Pending approvals', text: 'Which timesheets are waiting for approval?' },
      { title: 'Billable hours', text: 'What share of last month’s hours were billable, by project?' }
    ]
  };
}

export function copilotManifest(o: ICopilotOptions): Record<string, unknown> {
  return {
    $schema: 'https://developer.microsoft.com/json-schemas/teams/v1.19/MicrosoftTeams.schema.json',
    manifestVersion: '1.19',
    version: o.version,
    id: o.appId,
    developer: { name: 'Brillienta', websiteUrl: 'https://brillienta.com/timesheet', privacyUrl: 'https://brillienta.com/privacy', termsOfUseUrl: 'https://brillienta.com/terms' },
    icons: { color: 'color.png', outline: 'outline.png' },
    name: { short: `${o.appName} assistant`.substring(0, 30), full: `${o.appName} assistant for Microsoft 365 Copilot` },
    description: { short: 'Ask Copilot about projects, tasks and timesheets.', full: 'A Microsoft 365 Copilot agent that answers questions about projects, tasks, hours and approvals from Brillienta Timesheet.' },
    accentColor: '#038F62',
    copilotAgents: { declarativeAgents: [{ id: 'timesheetAgent', file: 'declarativeAgent.json' }] }
  };
}

/** Power BI Desktop data source file for the timesheet site's lists (SharePoint Online list connector, v2). */
export function pbids(siteUrl: string): string {
  return JSON.stringify({ version: '0.1', connections: [{ details: { protocol: 'sharepoint-list', address: { url: siteUrl } }, options: { Implementation: '2.0', ViewMode: 'All' }, mode: 'Import' }] }, null, 2);
}

/** Teams app package manifest: the web part as a personal app and a configurable channel tab. */
export function teamsManifest(o: { appId: string; webPartId: string; version: string; tenantHost: string; appName: string }): Record<string, unknown> {
  const page = `https://${o.tenantHost}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=${o.webPartId}%26forceLocale={locale}`;
  return {
    $schema: 'https://developer.microsoft.com/json-schemas/teams/v1.19/MicrosoftTeams.schema.json',
    manifestVersion: '1.19',
    version: o.version,
    id: o.appId,
    developer: { name: 'Brillienta', websiteUrl: 'https://brillienta.com/timesheet', privacyUrl: 'https://brillienta.com/privacy', termsOfUseUrl: 'https://brillienta.com/terms' },
    icons: { color: 'color.png', outline: 'outline.png' },
    name: { short: o.appName.substring(0, 30), full: o.appName },
    description: { short: 'Timesheets, approvals and project reports.', full: 'Weekly timesheets, approvals, projects, billable hours and reports from Brillienta Timesheet.' },
    accentColor: '#038F62',
    staticTabs: [{ entityId: 'timesheet', name: 'Timesheet', contentUrl: page.replace('openPropertyPane=true%26', ''), scopes: ['personal'] }],
    configurableTabs: [{ configurationUrl: page, canUpdateConfiguration: true, scopes: ['team'] }],
    validDomains: ['*.login.microsoftonline.com', '*.sharepoint.com', 'resourceseng.blob.core.windows.net'],
    webApplicationInfo: { resource: `https://${o.tenantHost}`, id: '00000003-0000-0ff1-ce00-000000000000' }
  };
}

export interface IEndpoint {
  key: string;
  title: string;
  rest: string;
  description: string;
}

/** REST endpoints of every list, for the API page. */
export function endpoints(siteUrl: string, prefix: string): IEndpoint[] {
  return LISTS.map((l) => ({
    key: l.key,
    title: listTitle(prefix, l.key),
    rest: `${siteUrl.replace(/\/$/, '')}/_api/web/lists/getbytitle('${listTitle(prefix, l.key)}')/items`,
    description: l.description
  }));
}
