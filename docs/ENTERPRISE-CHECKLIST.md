# Enterprise plan — feature checklist

Audit of every Enterprise item (and the Standard and Pro items Enterprise includes) against the build. "Browser" means
checked in the preview (real app, in-memory SharePoint, signed Enterprise licence key); "Unit" means covered by
`tests/*.test.cjs`; "Tenant" means it calls Microsoft 365 services that only a real tenant has, so it is built and
wired but must be confirmed on a tenant (see the last section).

| # | Feature | Status | Verified by |
| --- | --- | --- | --- |
| 1 | Select days for time tracking (holidays, per-person working days) | ✅ | Browser (holiday editor; person's days saved as `1,2,3,4,5,6`), Unit |
| 2 | Disable projects | ✅ | Browser (status → Disabled; no new time) |
| 3 | Additional users at ₹300/user | ✅ | Browser (licence page: 30 users, ₹300/user); user count per licence key |
| 4 | Configure start day of the week | ✅ | Browser (unlocked; warning when changed after weeks exist), Unit |
| 5 | Integrate with Time off Manager 365 | ✅ | Browser (approved leave put on the sheet Wed–Fri, not on the weekend, rejected leave ignored), Unit |
| 6 | Teams & Outlook connectivity | ✅ / Tenant | Browser (settings, Teams app package with personal app and tab); sending needs Graph |
| 7 | Planner integration | ✅ / Tenant | Code review (import, re-import updates, write-back); needs Graph |
| 8 | Copilot integration | ✅ / Tenant | Browser (agent package: manifest, declarativeAgent.json grounded on the site, icons), Unit |
| 9 | Project end date & notifications | ✅ | Browser (notice sent 6 days before the end), Unit |
| 10 | Planned project hours tracking | ✅ | Browser (logged / planned bars, over-plan alert), Unit |
| 11 | Restrict daily & weekly hours based on rules | ✅ | Browser (rule editor), Unit (first match, pinned rule, block / warn, weekends) |
| 12 | Custom reports & dashboards | ✅ | Browser (custom report saved; starter dashboard with 6 widgets saved) |
| 13 | Email templates | ✅ | Browser (editor with live preview of every template) |
| 14 | Delete projects | ✅ | Browser (project and its tasks deleted; logged time stays) |
| 15 | Add Office 365 external users | ✅ / Tenant | Browser (guest added and marked external); the invitation itself needs Graph |
| 16 | Customise with Power Automate | ✅ | Browser (HTTP trigger and events; importable email flow package) |
| 17 | API connectivity | ✅ | Browser (REST endpoints of all 13 lists), Unit |
| 18 | Connect with Power BI | ✅ | Browser (.pbids data source file), Unit |
| 19 | Dedicated account manager | ✅ | Browser (shown on Support from the licence key) |
| 20 | Brillienta Task Manager integration (added) | ✅ | Browser (sync, import, write-back on approval, removal on reopen), Unit |

Included from Pro — all ✅ in the browser: up to 30 users, unlimited support, user-specific rates, flat project
cost, Programme Manager role (sees its programme, approves time on its projects), role-based auto approval, daily and
weekly hour restrictions, my team report, pending weekly reminders (sent once per person and week), bulk activity
creation, activity tracking, select week days, external users.

Included from Standard — all ✅: clone weekly timesheets, detailed project report, billable / non-billable tracking,
four user roles, user-specific hourly cost, bulk project / task creation, custom columns, self-assign tasks
(support tickets: 4 a year on Standard only).

Plan gating: with a Pro key every Enterprise integration and the start-day setting show as locked; with a Standard key
the Pro features do (browser and unit tests).

## Fixed during the audit

- Approved leave spanning a weekend was also counted on Saturday and Sunday, inflating the week's total and possibly
  breaking a weekly maximum. Leave now counts only on the person's working days.
- Cloning a week now respects the person's own working days, not only the organisation's.
- Changing the start day of the week now warns that weeks already filled in keep their old dates.
- Adding an external user without Microsoft Graph said "Invitation sent"; it now says the person was added and how to
  give them access.

## To confirm on a real tenant

These call Microsoft 365 services the preview does not have. After approving the Graph permissions in the SharePoint
admin center (API access):

1. **Outlook email** — submit a week; the approver gets the email from the submitter's mailbox.
2. **Teams chat** — switch on "Also post notifications in Teams chat"; the approver gets a one-to-one chat message.
3. **Planner** — Project › Tasks › Import from Planner; then mark a linked task done and check Planner shows 100%.
4. **External users** — add one with the invitation switched on; they receive the Microsoft 365 guest invitation.
   Share the site with them and run Admin › Schema & security › Apply security.
5. **Copilot** — upload the agent package in Teams admin center; ask it about a project (needs Copilot licences).
6. **Power BI** — open the .pbids file in Power BI Desktop and load the lists.
7. **Time off Manager 365** — enter its site, list and column names (internal names) and use *Test the connection*;
   the defaults (`TOM_LeaveRequests`, `EmployeeEmail`, `StartDate`, `EndDate`, `RequestStatus` = `Approved`,
   `LeaveType`) are placeholders for your installation.
8. **Brillienta Task Manager** — *Test the connection*, *Sync now*; approve a week with a line on a linked task and
   check the task's time spent in Task Manager.
9. **Apply security** — check Site settings › People and groups shows BrillTS Administrators, BrillTS Approvers and
   BrillTS External Users with the right people; sign in as a member and confirm they see only their own timesheets.
