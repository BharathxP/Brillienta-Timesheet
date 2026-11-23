// Who may do what: roles, approvals, project management, auto-approval and hourly rates. Pure functions
// (tests/access.test.cjs).

import { CostMode, IProgram, IProject, ITask, ITimeEntry, ITimesheet, ITimesheetConfig, ITsUser, Role } from '../models/types';
import { IIssue } from './rules';

export const ROLE_LABELS: Record<Role, string> = {
  admin: 'Administrator',
  programManager: 'Programme Manager',
  projectManager: 'Project Manager',
  teamManager: 'Team Manager',
  member: 'Member'
};

export const ROLE_HELP: Record<Role, string> = {
  admin: 'Everything: settings, people, rates, every project, every timesheet and report.',
  programManager: 'Runs programmes: creates and manages their projects, approves time on them and sees their reports.',
  projectManager: 'Creates and runs projects (tasks, members, budgets), approves time on them and sees their reports.',
  teamManager: 'Approves the timesheets of the people who report to them and sees the My team report.',
  member: 'Fills in their own timesheet and assigns themselves open tasks.'
};

const lc = (s: string | undefined): string => (s || '').trim().toLowerCase();
export const sameEmail = (a: string | undefined, b: string | undefined): boolean => !!lc(a) && lc(a) === lc(b);
const hasEmail = (list: string[], e: string): boolean => list.some((x) => sameEmail(x, e));

export interface IDirectory {
  users: ITsUser[];
  projects: IProject[];
  programs: IProgram[];
  tasks: ITask[];
}

/** The role in force: an Administrator from the settings or site ownership wins; Programme Manager needs Pro. */
export function effectiveRole(user: ITsUser | undefined, opts: { siteAdmin: boolean; configAdmin: boolean; programManagerAllowed: boolean }): Role {
  if (opts.siteAdmin || opts.configAdmin) return 'admin';
  const r = user ? user.role : 'member';
  if (r === 'programManager' && !opts.programManagerAllowed) return 'projectManager';
  return r;
}

/** Everyone who reports to `email`, directly or indirectly (reporting loops are cut). */
export function reportsOf(users: ITsUser[], email: string, direct: boolean = false): ITsUser[] {
  const out: ITsUser[] = [];
  const seen: Record<string, boolean> = { [lc(email)]: true };
  let level = [lc(email)];
  while (level.length) {
    const next: string[] = [];
    users.forEach((u) => {
      if (u.active && level.indexOf(lc(u.managerEmail)) >= 0 && !seen[lc(u.email)]) {
        seen[lc(u.email)] = true;
        out.push(u);
        next.push(lc(u.email));
      }
    });
    if (direct) break;
    level = next;
  }
  return out;
}

/** Projects of the programmes `email` manages. */
export function programProjects(dir: IDirectory, email: string): IProject[] {
  const ids = dir.programs.filter((p) => hasEmail(p.managerEmails, email)).map((p) => p.id);
  return dir.projects.filter((p) => p.programId && ids.indexOf(p.programId) >= 0);
}

/** Can this person edit the project (details, tasks, members, rates)? */
export function canManageProject(role: Role, email: string, project: IProject, dir: IDirectory): boolean {
  if (role === 'admin') return true;
  if (hasEmail(project.managerEmails, email)) return role === 'projectManager' || role === 'programManager';
  if (role === 'programManager') return programProjects(dir, email).some((p) => p.id === project.id);
  return false;
}

/** Projects whose reports and approvals `email` is responsible for. */
export function managedProjects(role: Role, email: string, dir: IDirectory): IProject[] {
  if (role === 'admin') return dir.projects;
  return dir.projects.filter((p) => canManageProject(role, email, p, dir));
}

export function canCreateProjects(role: Role): boolean {
  return role === 'admin' || role === 'programManager' || role === 'projectManager';
}

/** Projects someone may log time on: active (or within the dates), and they are a member (or it is open to all). */
export function projectsFor(dir: IDirectory, email: string, onDay?: string): IProject[] {
  return dir.projects.filter(
    (p) =>
      p.status === 'Active' &&
      (!p.memberEmails.length || hasEmail(p.memberEmails, email) || hasEmail(p.managerEmails, email)) &&
      (!onDay || ((!p.startDate || p.startDate <= onDay) && (!p.endDate || p.endDate >= addWeek(onDay))))
  );
}

/** A project that ended during the week can still take that week's time. */
function addWeek(day: string): string {
  const d = new Date(Number(day.substring(0, 4)), Number(day.substring(5, 7)) - 1, Number(day.substring(8, 10)) - 6);
  const pad = (n: number): string => (n < 10 ? '0' : '') + n;
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/** Tasks someone may log time on within a project: assigned to them, unassigned, or (if allowed) self-assignable. */
export function tasksFor(dir: IDirectory, projectId: number, email: string): ITask[] {
  return dir.tasks.filter((t) => t.projectId === projectId && t.status !== 'Done' && (!t.assigneeEmails.length || hasEmail(t.assigneeEmails, email)));
}

/** Open tasks someone could take on: self-assignable, on a project they belong to, not theirs yet. */
export function selfAssignable(dir: IDirectory, email: string): ITask[] {
  const mine = projectsFor(dir, email).map((p) => p.id);
  return dir.tasks.filter((t) => t.selfAssign && t.status !== 'Done' && mine.indexOf(t.projectId) >= 0 && !hasEmail(t.assigneeEmails, email));
}

/**
 * Can `email` (with `role`) approve this timesheet? Never their own (except admins, whose own weeks can only be
 * auto-approved or approved by another admin). The person's manager approves when approvalBy is manager / either;
 * project and programme managers of a project on the sheet when approvalBy is project / either; admins always.
 */
export function canApprove(role: Role, email: string, sheet: ITimesheet, entries: ITimeEntry[], dir: IDirectory, cfg: ITimesheetConfig): boolean {
  if (sameEmail(sheet.userEmail, email)) return false;
  if (role === 'admin') return true;
  const owner = dir.users.filter((u) => sameEmail(u.email, sheet.userEmail))[0];
  if (cfg.approvalBy !== 'project' && owner && reportsOf(dir.users, email).some((u) => sameEmail(u.email, owner.email))) return true;
  if (cfg.approvalBy !== 'manager') {
    const projectIds = entries.filter((e) => e.total > 0).map((e) => e.projectId);
    return dir.projects.some((p) => projectIds.indexOf(p.id) >= 0 && canManageProject(role, email, p, dir));
  }
  return false;
}

/** Emails of the people who should be asked to approve a sheet (manager and/or project managers, else admins). */
export function approversFor(sheetOwner: string, entries: ITimeEntry[], dir: IDirectory, cfg: ITimesheetConfig, admins: string[]): string[] {
  const out: string[] = [];
  const add = (e: string): void => {
    if (e && !sameEmail(e, sheetOwner) && !hasEmail(out, e)) out.push(e);
  };
  const owner = dir.users.filter((u) => sameEmail(u.email, sheetOwner))[0];
  if (cfg.approvalBy !== 'project' && owner && owner.managerEmail) add(owner.managerEmail);
  if (cfg.approvalBy !== 'manager' || !out.length) {
    const ids = entries.filter((e) => e.total > 0).map((e) => e.projectId);
    dir.projects.filter((p) => ids.indexOf(p.id) >= 0).forEach((p) => p.managerEmails.forEach(add));
  }
  if (!out.length) admins.forEach(add);
  return out;
}

/** Auto-approve on submit? (role-based auto approval) */
export function shouldAutoApprove(cfg: ITimesheetConfig, role: Role, totalHours: number, issues: IIssue[]): boolean {
  const a = cfg.autoApproval;
  if (!a.enabled || a.roles.indexOf(role) < 0) return false;
  if (a.requireWithinRules && issues.length) return false;
  if (a.maxHours > 0 && totalHours > a.maxHours) return false;
  return true;
}

/**
 * Who belongs in the BrillTS security groups: administrators (the Administrator role, the admins listed in the
 * settings, and whoever applies security); approvers (programme, project and team managers, and anyone with people
 * reporting to them); external users (active guests). Each person is in one group only — the highest. Emails are
 * lower-case and unique.
 */
export function securityMembers(users: ITsUser[], admins: string[]): { admins: string[]; approvers: string[]; externals: string[] } {
  const uniq = (list: string[]): string[] => list.map(lc).filter((e, i, a) => !!e && a.indexOf(e) === i);
  const active = users.filter((u) => u.active);
  const a = uniq(admins.concat(active.filter((u) => u.role === 'admin').map((u) => u.email)));
  const managerEmails = active.filter((u) => u.role !== 'member' && u.role !== 'admin').map((u) => u.email).concat(active.filter((u) => active.some((x) => sameEmail(x.managerEmail, u.email))).map((u) => u.email));
  const ap = uniq(managerEmails).filter((e) => a.indexOf(e) < 0);
  const ex = uniq(active.filter((u) => u.external).map((u) => u.email)).filter((e) => a.indexOf(e) < 0 && ap.indexOf(e) < 0);
  return { admins: a, approvers: ap, externals: ex };
}

/** Hourly cost of `user` on `project`, by the cost mode (user-specific cost vs the project's flat hourly cost). */
export function costRate(mode: CostMode, user: ITsUser | undefined, project: IProject | undefined, fallback: number): number {
  const u = user && user.costRate > 0 ? user.costRate : 0;
  const p = project && project.hourlyCost > 0 ? project.hourlyCost : 0;
  switch (mode) {
    case 'userOnly':
      return u || fallback;
    case 'projectOnly':
      return p || fallback;
    case 'projectFirst':
      return p || u || fallback;
    default:
      return u || p || fallback;
  }
}

/** Bill rate of `user` on `project`: person-specific rate on the project, the person's rate, the project's rate. */
export function billRate(user: ITsUser | undefined, project: IProject | undefined, fallback: number, userRates: boolean): number {
  if (project && user && userRates) {
    const key = Object.keys(project.memberRates || {}).filter((k) => sameEmail(k, user.email))[0];
    if (key && project.memberRates[key] > 0) return project.memberRates[key];
  }
  if (user && userRates && user.billRate > 0) return user.billRate;
  if (project && project.billRate > 0) return project.billRate;
  return fallback;
}

/** Billable default for a new line: the task's, else the activity's and project's. */
export function defaultBillable(project: IProject | undefined, task: ITask | undefined, activityBillable: boolean | undefined): boolean {
  if (project && !project.billable) return false;
  if (task) return task.billable;
  if (activityBillable !== undefined) return activityBillable;
  return project ? project.billable : true;
}
