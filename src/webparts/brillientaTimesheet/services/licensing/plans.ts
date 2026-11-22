// Plans and the features each one includes. Pure functions, unit tested (tests/licensing.test.cjs).
//
// The active plan is applied on top of the saved configuration every time it loads: a feature the plan does not
// include is switched off in the configuration the app runs with, whatever the Settings list says, so editing the
// list cannot unlock anything. Saved values are kept (restoreLocked), so upgrading brings earlier settings back.

import { ITimesheetConfig } from '../../models/types';
import { defaultConfig, defaultTemplates } from '../../schema/defaults';

export type Plan = 'standard' | 'pro' | 'enterprise';

export type Feature =
  // Standard
  | 'timesheets'
  | 'cloneWeek'
  | 'projectReport'
  | 'billable'
  | 'projectCost'
  | 'roles'
  | 'teamReport'
  | 'userCost'
  | 'bulkCreate'
  | 'customColumns'
  | 'selfAssign'
  | 'export'
  // Pro
  | 'userRates'
  | 'programManager'
  | 'autoApproval'
  | 'hourLimits'
  | 'reminders'
  | 'bulkActivities'
  | 'activityTracking'
  | 'workDays'
  | 'unlimitedSupport'
  // Enterprise
  | 'trackingDays'
  | 'disableProjects'
  | 'weekStart'
  | 'timeOff'
  | 'teamsOutlook'
  | 'planner'
  | 'copilot'
  | 'projectEndDates'
  | 'plannedHours'
  | 'hourRules'
  | 'customReports'
  | 'emailTemplates'
  | 'deleteProjects'
  | 'externalUsers'
  | 'powerAutomate'
  | 'api'
  | 'powerBi'
  | 'accountManager';

export interface IPlanInfo {
  plan: Plan;
  name: string;
  tagline: string;
  /** Users included (a licence can set its own number). */
  users: number;
  /** Price of each extra user (INR), 0 = not offered. */
  extraUserPrice: number;
  /** Support tickets per year (0 = unlimited). */
  tickets: number;
}

export const PLANS: IPlanInfo[] = [
  { plan: 'standard', name: 'Standard', tagline: 'Weekly timesheets, projects, billable hours, rates, roles and reports', users: 10, extraUserPrice: 0, tickets: 4 },
  { plan: 'pro', name: 'Pro', tagline: 'Programme managers, auto-approval, hour limits, reminders and activity tracking', users: 30, extraUserPrice: 200, tickets: 0 },
  { plan: 'enterprise', name: 'Enterprise', tagline: 'Time off, Teams, Outlook, Planner, Copilot, custom dashboards, Power BI and API', users: 30, extraUserPrice: 300, tickets: 0 }
];

const RANK: Record<Plan, number> = { standard: 1, pro: 2, enterprise: 3 };

export const FEATURES: Array<{ feature: Feature; label: string; plan: Plan }> = [
  { feature: 'timesheets', label: 'Weekly timesheets with approvals', plan: 'standard' },
  { feature: 'cloneWeek', label: 'Clone weekly timesheets', plan: 'standard' },
  { feature: 'projectReport', label: 'Detailed project report', plan: 'standard' },
  { feature: 'billable', label: 'Billable / non-billable hours tracking', plan: 'standard' },
  { feature: 'projectCost', label: 'Flat hourly cost for projects', plan: 'standard' },
  { feature: 'roles', label: 'Four user roles (Admin, Project Manager, Team Manager, Member)', plan: 'standard' },
  { feature: 'teamReport', label: 'My team report', plan: 'standard' },
  { feature: 'userCost', label: 'User-specific hourly cost', plan: 'standard' },
  { feature: 'bulkCreate', label: 'Bulk project / task creation', plan: 'standard' },
  { feature: 'customColumns', label: 'Custom columns', plan: 'standard' },
  { feature: 'selfAssign', label: 'Self-assign tasks', plan: 'standard' },
  { feature: 'export', label: 'Excel and CSV export', plan: 'standard' },
  { feature: 'userRates', label: 'User-specific bill rates (per project)', plan: 'pro' },
  { feature: 'programManager', label: 'Programme Manager role', plan: 'pro' },
  { feature: 'autoApproval', label: 'Role-based auto timesheet approval', plan: 'pro' },
  { feature: 'hourLimits', label: 'Daily and weekly hour restrictions', plan: 'pro' },
  { feature: 'reminders', label: 'Pending weekly timesheet reminders', plan: 'pro' },
  { feature: 'bulkActivities', label: 'Bulk activity creation', plan: 'pro' },
  { feature: 'activityTracking', label: 'Activity tracking', plan: 'pro' },
  { feature: 'workDays', label: 'Select week days for time tracking', plan: 'pro' },
  { feature: 'unlimitedSupport', label: 'Unlimited support and free updates', plan: 'pro' },
  { feature: 'trackingDays', label: 'Select days for time tracking (holidays, per-person days)', plan: 'enterprise' },
  { feature: 'disableProjects', label: 'Disable projects', plan: 'enterprise' },
  { feature: 'weekStart', label: 'Configure the start day of the week', plan: 'enterprise' },
  { feature: 'timeOff', label: 'Integrate with Time off Manager 365', plan: 'enterprise' },
  { feature: 'teamsOutlook', label: 'Teams and Outlook connectivity', plan: 'enterprise' },
  { feature: 'planner', label: 'Planner integration', plan: 'enterprise' },
  { feature: 'copilot', label: 'Copilot integration', plan: 'enterprise' },
  { feature: 'projectEndDates', label: 'Project end dates and notifications', plan: 'enterprise' },
  { feature: 'plannedHours', label: 'Planned project hours tracking', plan: 'enterprise' },
  { feature: 'hourRules', label: 'Daily and weekly hours based on rules', plan: 'enterprise' },
  { feature: 'customReports', label: 'Custom reports and dashboards', plan: 'enterprise' },
  { feature: 'emailTemplates', label: 'Email templates', plan: 'enterprise' },
  { feature: 'deleteProjects', label: 'Delete projects', plan: 'enterprise' },
  { feature: 'externalUsers', label: 'Microsoft 365 external users', plan: 'enterprise' },
  { feature: 'powerAutomate', label: 'Customise with Power Automate', plan: 'enterprise' },
  { feature: 'api', label: 'API connectivity', plan: 'enterprise' },
  { feature: 'powerBi', label: 'Connect with Power BI', plan: 'enterprise' },
  { feature: 'accountManager', label: 'Dedicated account manager', plan: 'enterprise' }
];

/** Add-ons a licence can include on a lower plan. External users are an option on Pro. */
export type Addon = 'externalUsers';
export const ADDONS: Array<{ addon: Addon; label: string; plans: Plan[] }> = [{ addon: 'externalUsers', label: 'Microsoft 365 external users', plans: ['pro'] }];

export function planInfo(plan: Plan): IPlanInfo {
  return PLANS.filter((p) => p.plan === plan)[0];
}

export function featurePlan(feature: Feature): Plan {
  const f = FEATURES.filter((x) => x.feature === feature)[0];
  return f ? f.plan : 'enterprise';
}

/** Whether a plan (plus add-ons) includes a feature. No plan (trial over, licence ended) includes nothing. */
export function planHas(plan: Plan | undefined, feature: Feature, addons?: string[]): boolean {
  if (!plan) return false;
  if (RANK[plan] >= RANK[featurePlan(feature)]) return true;
  return (addons || []).indexOf(feature) >= 0 && ADDONS.some((a) => a.addon === feature && a.plans.indexOf(plan) >= 0);
}

/** Users a plan allows; a licence's own number wins when set. */
export function userLimit(plan: Plan | undefined, licenceUsers?: number): number {
  if (!plan) return 0;
  return licenceUsers && licenceUsers > 0 ? licenceUsers : planInfo(plan).users;
}

/** Support tickets a plan allows per year (0 = unlimited). */
export function ticketLimit(plan: Plan | undefined): number {
  return plan ? planInfo(plan).tickets : 0;
}

// ---------------------------------------------------------------- applying a plan to the configuration

type Mutator = (c: ITimesheetConfig, d: ITimesheetConfig) => void;

/** What each locked feature resets in the running configuration (d = the defaults). */
const LOCKS: Array<{ feature: Feature; apply: Mutator }> = [
  { feature: 'billable', apply: (c) => (c.allowBillableOverride = false) },
  { feature: 'customColumns', apply: (c) => (c.customColumns = []) },
  { feature: 'selfAssign', apply: (c) => (c.selfAssign = false) },
  { feature: 'userCost', apply: (c) => (c.costMode = c.costMode === 'userFirst' || c.costMode === 'userOnly' ? 'projectOnly' : c.costMode) },
  { feature: 'autoApproval', apply: (c) => (c.autoApproval = { ...c.autoApproval, enabled: false }) },
  { feature: 'hourLimits', apply: (c, d) => (c.limits = d.limits) },
  { feature: 'reminders', apply: (c) => (c.reminders = { ...c.reminders, enabled: false }) },
  { feature: 'activityTracking', apply: (c) => (c.requireActivity = false) },
  { feature: 'workDays', apply: (c, d) => (c.workDays = d.workDays) },
  { feature: 'trackingDays', apply: (c) => (c.holidays = []) },
  { feature: 'weekStart', apply: (c, d) => (c.weekStart = d.weekStart) },
  { feature: 'timeOff', apply: (c) => (c.timeOff = { ...c.timeOff, enabled: false }) },
  { feature: 'teamsOutlook', apply: (c) => (c.integrations = { ...c.integrations, teamsChat: false, mailMode: 'flow' }) },
  { feature: 'planner', apply: (c) => (c.integrations = { ...c.integrations, plannerEnabled: false, plannerWriteBack: false }) },
  { feature: 'copilot', apply: (c) => (c.integrations = { ...c.integrations, copilotEnabled: false }) },
  { feature: 'projectEndDates', apply: (c) => (c.projectEndNotice = []) },
  { feature: 'plannedHours', apply: (c) => (c.plannedHoursAlerts = false) },
  { feature: 'hourRules', apply: (c) => (c.rulesEnabled = false) },
  { feature: 'emailTemplates', apply: (c) => (c.templates = defaultTemplates()) },
  { feature: 'powerAutomate', apply: (c) => (c.integrations = { ...c.integrations, webhookUrl: '' }) }
];

/** The configuration the app runs with under `plan`: everything the plan does not include is off. */
export function applyPlan(saved: ITimesheetConfig, plan: Plan | undefined, addons?: string[]): ITimesheetConfig {
  const c: ITimesheetConfig = JSON.parse(JSON.stringify(saved));
  const d = defaultConfig(saved.orgName);
  LOCKS.forEach((l) => {
    if (!planHas(plan, l.feature, addons)) l.apply(c, d);
  });
  return c;
}

/**
 * What to save when an admin saves settings: locked features keep their earlier saved values (the screens show them
 * off; saving must not wipe them).
 */
export function restoreLocked(next: ITimesheetConfig, saved: ITimesheetConfig, plan: Plan | undefined, addons?: string[]): ITimesheetConfig {
  const out: ITimesheetConfig = JSON.parse(JSON.stringify(next));
  const keys: Partial<Record<Feature, Array<keyof ITimesheetConfig>>> = {
    customColumns: ['customColumns'],
    autoApproval: ['autoApproval'],
    hourLimits: ['limits'],
    reminders: ['reminders'],
    workDays: ['workDays'],
    trackingDays: ['holidays'],
    weekStart: ['weekStart'],
    timeOff: ['timeOff'],
    projectEndDates: ['projectEndNotice'],
    hourRules: ['rulesEnabled', 'rules'],
    emailTemplates: ['templates']
  };
  (Object.keys(keys) as Feature[]).forEach((f) => {
    if (planHas(plan, f, addons)) return;
    (keys[f] || []).forEach((k) => {
      (out as unknown as Record<string, unknown>)[k] = JSON.parse(JSON.stringify(saved[k]));
    });
  });
  return out;
}
