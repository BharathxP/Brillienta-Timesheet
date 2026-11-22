import * as React from 'react';
import { ICurrentUser, ITimesheetConfig, ITsUser, Role } from '../models/types';
import { SPClient } from '../services/SPClient';
import { TimesheetService, IReference } from '../services/TimesheetService';
import { ISecurityMeta, SettingsStore } from '../services/SettingsStore';
import { GraphService } from '../services/GraphService';
import { Notifier } from '../services/Notifier';
import { Workflow } from '../services/workflow';
import { ReminderRunner } from '../services/ReminderRunner';
import { TimeOffService } from '../services/TimeOffService';
import { TaskManagerService } from '../services/TaskManagerService';
import { ILicenseStatus } from '../services/licensing/license';
import { Feature } from '../services/licensing/plans';
import { IDirectory } from '../services/access';

export interface IRoute {
  path: string;
  parts: string[];
  query: Record<string, string>;
}

export interface IHostInfo {
  inTeams: boolean;
  siteSource: '' | 'property' | 'tenant';
  hostWebUrl: string;
  siteError?: string;
  openTeamsLink?: (url: string) => Promise<void>;
}

export const SHAREPOINT_HOST: IHostInfo = { inTeams: false, siteSource: '', hostWebUrl: '' };

export type ToastKind = 'success' | 'error' | 'info' | 'warning';

export interface IAppContext {
  sp: SPClient;
  svc: TimesheetService;
  settings: SettingsStore;
  graph?: GraphService;
  timeOff: TimeOffService;
  /** Brillienta Task Manager integration. */
  taskManager: TaskManagerService;
  notifier: Notifier;
  workflow: Workflow;
  reminders: ReminderRunner;
  /** Settings in force (the saved settings limited to what the plan includes). */
  config: ITimesheetConfig;
  /** Settings as saved (admin screens edit these). */
  savedConfig: ITimesheetConfig;
  saveConfig: (c: ITimesheetConfig, what?: string) => Promise<void>;
  me: ICurrentUser;
  /** The signed-in person's TS_Users row. */
  myUser?: ITsUser;
  role: Role;
  roleReason: string;
  isAdmin: boolean;
  /** Approves for someone (manager, project or programme manager, admin). */
  isApprover: boolean;
  ref: IReference;
  dir: IDirectory;
  reloadRef: () => Promise<void>;
  license: ILicenseStatus;
  has: (feature: Feature) => boolean;
  refreshLicense: () => Promise<void>;
  /** Active users beyond the plan's user limit (they cannot submit). */
  usersOver: number;
  /** The signed-in person is beyond the user limit. */
  overLimit: boolean;
  admins: string[];
  logo: string;
  saveLogo: (dataUrl: string) => Promise<void>;
  security?: ISecurityMeta;
  setSecurity: (m: ISecurityMeta) => void;
  approvalsCount: number;
  refreshApprovals: () => void;
  notificationsVersion: number;
  bumpNotifications: () => void;
  prefix: string;
  appVersion: string;
  tenantId?: string;
  host: IHostInfo;
  linkBase: string;
  route: IRoute;
  navigate: (path: string) => void;
  toast: (message: string, kind?: ToastKind) => void;
}

export const AppCtx = React.createContext<IAppContext | undefined>(undefined);

export function useApp(): IAppContext {
  const ctx = React.useContext(AppCtx);
  if (!ctx) throw new Error('AppCtx missing');
  return ctx;
}

export function parseRoute(hash: string): IRoute {
  const raw = (hash || '').replace(/^#\/?/, '');
  const [path, qs] = raw.split('?');
  const query: Record<string, string> = {};
  (qs || '').split('&').forEach((pair) => {
    if (!pair) return;
    const [k, v] = pair.split('=');
    query[decodeURIComponent(k)] = decodeURIComponent(v || '');
  });
  const clean = (path || '').replace(/\/$/, '');
  return { path: clean, parts: clean ? clean.split('/') : [], query };
}

export function buildQuery(q: Record<string, string | number | undefined>): string {
  const parts = Object.keys(q)
    .filter((k) => q[k] !== undefined && q[k] !== '')
    .map((k) => `${encodeURIComponent(k)}=${encodeURIComponent(String(q[k]))}`);
  return parts.length ? '?' + parts.join('&') : '';
}

export function errorText(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}
