// The product as it is published: name, publisher, links and store descriptions. One source for the SharePoint package
// (config/package-solution.json must match: tests/teams.test.cjs), the Teams app manifest (teams/manifest.json, made
// by tools/teams/make-app-package.cjs) and the packages admins download in the app (Copilot agent, Teams app).

/** The web part's id (BrillientaTimesheetWebPart.manifest.json); Teams uses it as the Teams app id. */
export const WEB_PART_ID = 'ff5b1593-a3b3-45f5-b3a5-8d11821fc2a1';

export const PRODUCT = {
  /** The marketplace name (AppSource, SharePoint Store, Teams): with a space, as in the listing. */
  name: 'Brillienta Timesheet',
  developer: {
    name: 'Brillienta',
    websiteUrl: 'https://brillienta.com/timesheet',
    privacyUrl: 'https://brillienta.com/privacy',
    termsOfUseUrl: 'https://brillienta.com/terms',
    mpnId: '7194411'
  },
  helpUrl: 'https://kb.brillienta.com/timesheet/',
  contactUrl: 'https://brillienta.com/contact',
  supportEmail: 'support@brillienta.com',
  /** Brillienta green (the logo tile). */
  accentColor: '#038F62',
  /** ≤ 80 characters (Teams). */
  shortDescription: 'Timesheets, approvals, projects and cost reports for SharePoint and Teams.',
  /** ≤ 4,000 characters (Teams); ends with Get started, Help and Contact us (Marketplace). */
  longDescription:
    'Brillienta Timesheet is a time tracking app that runs inside SharePoint Online and Microsoft Teams. All data stays in SharePoint lists on your own site. ' +
    'People fill in weekly timesheets with projects, tasks, activities, billable and non-billable hours and notes, clone earlier weeks and submit them for approval. ' +
    'Managers and project managers approve or send back weeks, one at a time or in bulk, with role-based auto-approval and daily and weekly hour limits, including rules by role, department or external user. ' +
    'It includes projects and programmes with planned hours, end dates and notifications; bulk creation of projects, tasks and activities; self-assigned tasks; custom columns; user and project hourly cost and bill rates; ' +
    'project, team, billable, cost and utilisation reports with Excel export; custom reports and dashboards; pending timesheet reminders; email templates; and branded SharePoint security groups. ' +
    'Integrations: Microsoft Teams and Outlook notifications, Microsoft Planner, Microsoft 365 Copilot, Power Automate, Power BI, the SharePoint REST API, Time off Manager 365 and Brillienta Task Manager. ' +
    'Includes a 14-day free trial of every feature; after the trial a licence key from Brillienta is required (Standard, Pro or Enterprise). ' +
    'Microsoft Graph permissions, all optional: User.Read.All (people search), Group.Read.All and Tasks.ReadWrite (Planner), Mail.Send, Chat.Create and ChatMessage.Send (Outlook and Teams notifications) and User.Invite.All (inviting external users). ' +
    'To use it in Microsoft Teams, a SharePoint admin selects the app in the tenant app catalog and selects Sync to Teams; it then appears in Teams under Apps, Built for your org, as Brillienta Timesheet. ' +
    'Get started: start your free 14-day trial at https://brillienta.com/timesheet. Help: read the product guide at https://kb.brillienta.com/timesheet/. Contact us: https://brillienta.com/contact or support@brillienta.com.'
};

/** The Teams app manifest version: 1.25+ is required for apps with channel tabs (supportsChannelFeatures). */
export const TEAMS_MANIFEST_VERSION = '1.25';
/** SharePoint Online's first-party app id: required for SharePoint Framework apps in Teams. */
export const SHAREPOINT_APP_ID = '00000003-0000-0ff1-ce00-000000000000';
