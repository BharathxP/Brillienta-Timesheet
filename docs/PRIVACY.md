# Brillienta Timesheet — Privacy statement

Brillienta Timesheet is a SharePoint Framework app that runs entirely inside your Microsoft 365 tenant, in the browser
of the person using it (on a SharePoint page or in Microsoft Teams).

## What data is used

- **Timesheet data**: weeks, time entries (hours per day, notes, ticket numbers, custom columns), approvals and their
  comments, projects, programmes, tasks, activities, holidays and the hour limits your admins set. Stored in SharePoint
  lists on your timesheet site (their names start with `BrillTS_`).
- **People and rates**: names, email addresses, managers, departments, roles, and the hourly cost and bill rates your
  admins enter. People are found with Microsoft Graph people search, using the signed-in person's delegated permission
  (`User.Read.All`), and stored in the same lists.
- **Notifications**, when an admin switches them on: emails (`Mail.Send`) and Teams chat messages (`Chat.Create`,
  `ChatMessage.Send`) are sent with the signed-in person's own permissions, or added to a SharePoint list and sent by a
  Power Automate flow your admins create in your tenant, from the mailbox they choose.
- **Microsoft Planner** (optional): plans and tasks are read and updated with the person's permissions
  (`Group.Read.All`, `Tasks.ReadWrite`).
- **External users** (optional): an admin can invite a guest (`User.Invite.All`); their name and email are stored in the
  people list.
- **Other Brillienta apps** (optional): with Time off Manager or Brillienta Task Manager connected, the app reads their
  SharePoint lists on your sites (leave, projects, tasks, time) and, if an admin chooses, writes approved hours back.
- **Audit log**: who changed what and when (settings, approvals, rates, people), in a SharePoint list on your site.

## Where data is stored and who can see it

All data stays in your tenant: in SharePoint lists on the site that hosts the timesheet, protected by SharePoint
permissions and the app's branded security groups (`BrillTS Administrators`, `BrillTS Approvers` and others). People
see their own timesheets; managers and project managers see the weeks they approve; costs and rates are shown only to
roles allowed to see them. Operational lists (audit log, notifications, email queue, support tickets, saved reports)
are kept out of SharePoint search.

The publisher operates **no server** and receives **no timesheet data, usage data or telemetry**. The app calls only
your SharePoint sites' REST API, Microsoft Graph, Microsoft's CDN for the Fluent UI icon font, and any webhook URL your
admins enter themselves (Admin › Integrations).

The one exception is a **support ticket**: when someone sends one from the Support page, its subject, details, category,
priority and the sender's name and email are emailed to Brillienta support so we can answer.

## Retention and removal

Data is kept until your admins delete it. Removing the app and deleting the `BrillTS_` lists and `BrillTS` groups
removes all of its data. The licence key is stored in the settings list; it holds your organisation name and plan, not
personal data.

## Contact

Questions about how your organisation uses the timesheet go to your Microsoft 365 administrators. Questions for the
publisher: support@brillienta.com or https://brillienta.com/contact.
