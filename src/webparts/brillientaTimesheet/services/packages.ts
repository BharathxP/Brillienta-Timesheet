// Files the admin downloads from Admin › Integrations. Pure (tests/packages.test.cjs).
//
// - Copilot: a Microsoft 365 Copilot declarative agent (Teams app package) grounded on the timesheet site, so people
//   can ask Copilot about projects, tasks and timesheets in natural language. Upload it in Teams admin center or
//   Teams › Apps › Manage your apps › Upload an app.
// - Power BI: a .pbids data source file that opens Power BI Desktop connected to the site's SharePoint lists.
// - Teams: a Teams app package with the web part as a personal app and a channel tab (SharePoint-hosted).
// - API: the REST and Microsoft Graph endpoints of every list.

import { LISTS, listTitle } from '../schema/schema';
import { PRODUCT, SHAREPOINT_APP_ID, TEAMS_MANIFEST_VERSION, WEB_PART_ID } from './product';

export interface ICopilotOptions {
  siteUrl: string;
  orgName: string;
  appName: string;
  prefix: string;
  /** Leave empty for the stable id derived from the site. */
  appId?: string;
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

/** A stable UUID for a text (FNV-1a based, version-4 layout): the same site always gets the same Copilot app id. */
export function stableId(text: string): string {
  let hex = '';
  for (let round = 0; round < 4; round++) {
    let h = 0x811c9dc5 ^ round;
    for (let i = 0; i < text.length; i++) {
      h ^= text.charCodeAt(i);
      h = Math.imul(h, 0x01000193) >>> 0;
    }
    hex += ('0000000' + h.toString(16)).slice(-8);
  }
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-${'89ab'[parseInt(hex[16], 16) % 4]}${hex.slice(17, 20)}-${hex.slice(20, 32)}`;
}

/** Icon file names: named after the app id, as SharePoint's Sync to Teams expects for SPFx apps. */
export function teamsIconNames(id: string): { color: string; outline: string } {
  return { color: `${id}_color.png`, outline: `${id}_outline.png` };
}

const developer = (): Record<string, string> => ({ ...PRODUCT.developer });

/**
 * The Copilot agent's Teams app manifest. Its id is derived from the site, so downloading again and uploading updates
 * the same agent rather than adding a second one.
 */
export function copilotManifest(o: ICopilotOptions): Record<string, unknown> {
  const id = o.appId || stableId('copilot|' + o.siteUrl.toLowerCase());
  return {
    $schema: `https://developer.microsoft.com/en-us/json-schemas/teams/v${TEAMS_MANIFEST_VERSION}/MicrosoftTeams.schema.json`,
    manifestVersion: TEAMS_MANIFEST_VERSION,
    version: o.version,
    id,
    developer: developer(),
    icons: teamsIconNames(id),
    name: { short: `${o.appName} assistant`.substring(0, 30), full: `${o.appName} assistant for Microsoft 365 Copilot` },
    description: {
      short: 'Ask Copilot about your projects, tasks, hours and approvals.',
      full: `A Microsoft 365 Copilot agent for ${o.orgName || 'your organisation'} that answers questions about projects, tasks, hours and approvals from Brillienta Timesheet. It only sees what each person may see in SharePoint.`
    },
    accentColor: PRODUCT.accentColor,
    copilotAgents: { declarativeAgents: [{ id: 'timesheetAgent', file: 'declarativeAgent.json' }] }
  };
}

/** Power BI Desktop data source file for the timesheet site's lists (SharePoint Online list connector, v2). */
export function pbids(siteUrl: string): string {
  return JSON.stringify({ version: '0.1', connections: [{ details: { protocol: 'sharepoint-list', address: { url: siteUrl } }, options: { Implementation: '2.0', ViewMode: 'All' }, mode: 'Import' }] }, null, 2);
}

/**
 * The Teams app manifest: the web part as a personal app and a channel tab, with the SPFx URLs Teams fills in
 * ({teamSiteDomain}, {teamSitePath}, {locale}). The same manifest ships in the package (teams/manifest.json, used by
 * Sync to Teams) and is offered for download in Admin › Integrations › Teams.
 */
export function teamsManifest(version: string): Record<string, unknown> {
  const id = WEB_PART_ID;
  const personal = `https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamshostedapp.aspx%3Fteams%26personal%26componentId=${id}%26forceLocale={locale}`;
  const tab = `https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=${id}%26forceLocale={locale}`;
  return {
    $schema: `https://developer.microsoft.com/en-us/json-schemas/teams/v${TEAMS_MANIFEST_VERSION}/MicrosoftTeams.schema.json`,
    manifestVersion: TEAMS_MANIFEST_VERSION,
    version,
    // The web part id, as Sync to Teams uses: the app is updated in place, never added twice.
    id,
    developer: developer(),
    name: { short: PRODUCT.name, full: PRODUCT.name },
    description: { short: PRODUCT.shortDescription, full: PRODUCT.longDescription },
    icons: teamsIconNames(id),
    accentColor: PRODUCT.accentColor,
    staticTabs: [{ entityId: id, name: 'Timesheet', contentUrl: personal, websiteUrl: PRODUCT.developer.websiteUrl, scopes: ['personal'] }],
    configurableTabs: [{ configurationUrl: tab, canUpdateConfiguration: true, scopes: ['team'], context: ['channelTab'] }],
    // Works in standard, private and shared channels: the tab doesn't use channel or team membership (tier1).
    supportsChannelFeatures: 'tier1',
    permissions: ['identity'],
    validDomains: ['*.login.microsoftonline.com', '*.sharepoint.com', '*.sharepoint-df.com', 'spoppe-a.akamaihd.net', 'spoprod-a.akamaihd.net', 'resourceseng.blob.core.windows.net', 'msft.spoppe.com'],
    webApplicationInfo: { resource: 'https://{teamSiteDomain}', id: SHAREPOINT_APP_ID }
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
