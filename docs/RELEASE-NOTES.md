# Release notes

## 1.0.0

First release of **Brillienta Timesheet** for SharePoint Online and Microsoft Teams.

- **Timesheets**: weekly grid with projects, tasks, activities, billable flags, ticket numbers, notes and custom
  columns; clone a week; typed hours (`7.5`, `7:30`, `45m`); daily and weekly limits by organisation, role, department or
  external user; public holidays reduce the weekly target; time off from Time off Manager fills in automatically.
- **Approvals**: by manager or project manager, one at a time or in bulk, with comments, send back and reopen;
  role-based auto-approval; pending timesheet reminders.
- **Projects**: programmes, clients, planned hours, end dates and over-plan alerts, bulk creation of projects, tasks and
  activities, self-assigned tasks, hourly cost and bill rates per person and project.
- **Reports**: project, team, billable, cost and margin, utilisation and custom reports, dashboards, Excel and CSV
  export.
- **Integrations**: Teams and Outlook notifications, Microsoft Planner, Microsoft 365 Copilot agent, Power Automate,
  Power BI, the SharePoint REST API, Time off Manager 365 and Brillienta Task Manager.
- **Security**: item-level permissions, branded SharePoint groups (`BrillTS Administrators`, `BrillTS Approvers`, …),
  audit log, external users. Lists are named `BrillTS_…`; operational lists are kept out of search.
- **Licensing**: 14-day trial of every feature, then Standard, Pro or Enterprise keys.

Publication:

- The app is named **Brillienta Timesheet**, with a space, in the app catalog, the web part picker and Teams, matching
  the marketplace listing; the picker shows the app logo. The long description ends with Get started, Help and Contact
  us links. Five 1366 × 768 screenshots are in the package.
- Microsoft Teams: the package includes its own Teams app manifest (`teams/manifest.json`, v1.25, with
  `supportsChannelFeatures: "tier1"`, which Teams now requires for apps with channel tabs), so **Sync to Teams** adds
  the app as **Brillienta Timesheet** with the personal app and the channel tab. The Teams app id is the web part id, so
  later versions update the same app. The download under Admin › Integrations › Teams is the same manifest.
- Teams finds the timesheet by itself: when the team's site has none and nothing points to one, the app finds the
  timesheet sites the person can open with SharePoint search. With several, it opens the one marked **Use this site in
  Teams** (Admin › Integrations › Teams), or asks once and remembers.
- The Copilot agent download keeps the same app id for the same site, so downloading it again updates the agent.

Fixes found in end-to-end testing:

- Date fields (holidays, project and task dates, report periods, custom date columns) use the app's date format
  ("25 Nov 2026") and start the calendar on the organisation's first day of the week. Before, the holiday picker was
  34 px wide and report dates showed in the browser's format.
- The weekly minimum is reduced by public holidays in the week ("40 h less public holidays").
