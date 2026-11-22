import { SPClient, SPError, odataString } from './SPClient';
import { IFieldDef, IListDef, LISTS, SecurityProfile, expectedTypeAsString, fieldSchemaXml, listTitle } from '../schema/schema';

export type LogLevel = 'info' | 'success' | 'warning' | 'error';
export type ProgressFn = (message: string, level: LogLevel) => void;

export interface IListReport {
  key: string;
  title: string;
  exists: boolean;
  created: boolean;
  missingFields: string[];
  errors: string[];
}

export interface ISchemaReport {
  ok: boolean;
  lists: IListReport[];
  checkedAt: string;
}

export interface ISecurityPlan {
  tenantId?: string;
  /** Administrators (emails): edit every list. */
  admins: string[];
  /** Approvers and managers (emails): see and edit every timesheet; edit projects, tasks and people. */
  managers: string[];
  /** Microsoft 365 guests who track time (they are not in "Everyone except external users"). */
  externals: string[];
}

interface ISPField {
  InternalName: string;
  TypeAsString: string;
  Indexed: boolean;
  Choices?: string[];
}

interface IRoleAssignment {
  PrincipalId: number;
  Member: { Id: number; Title: string; PrincipalType: number };
  RoleDefinitionBindings: Array<{ Id: number; RoleTypeKind: number }>;
}

// SP.RoleType
const READER = 2;
const CONTRIBUTOR = 3;
const ADMINISTRATOR = 5;
const EDITOR = 6;
// Add field using its Name as internal name (8) and show it in the default view (16).
const FIELD_OPTIONS = 8 | 16;

/** Item-level security of a profile: 2 = people only see / edit the rows they created (approvers see all). */
function itemSecurity(p: SecurityProfile): 1 | 2 {
  return p === 'timesheet' || p === 'adminOnly' ? 2 : 1;
}

/**
 * Creates, validates and repairs every list and column in schema/schema.ts, and applies list permissions.
 * Needs a person with Manage Lists (and Manage Permissions for security).
 */
export class Provisioner {
  private readonly _sp: SPClient;
  private readonly _prefix: string;

  public constructor(sp: SPClient, prefix: string) {
    this._sp = sp;
    this._prefix = prefix;
  }

  public async ensureSchema(log: ProgressFn, repair: boolean = true): Promise<ISchemaReport> {
    const lists: IListReport[] = [];
    for (const def of LISTS) lists.push(await this._ensureList(def, log, repair));
    return { ok: lists.every((l) => l.exists && !l.errors.length && (repair || !l.missingFields.length)), lists, checkedAt: new Date().toISOString() };
  }

  private async _ensureList(def: IListDef, log: ProgressFn, repair: boolean): Promise<IListReport> {
    const title = listTitle(this._prefix, def.key);
    const report: IListReport = { key: def.key, title, exists: true, created: false, missingFields: [], errors: [] };
    const listPath = this._sp.list(title);
    try {
      let info: { EnableVersioning: boolean; ItemCount: number; ReadSecurity: number; WriteSecurity: number; Hidden?: boolean } | undefined;
      try {
        info = await this._sp.get(`${listPath}?$select=Id,EnableVersioning,ItemCount,ReadSecurity,WriteSecurity,Hidden`);
      } catch (e) {
        if (!(e instanceof SPError && e.status === 404)) throw e;
      }
      if (!info) {
        report.exists = false;
        if (!repair) {
          log(`List ${title} is missing`, 'warning');
          report.missingFields = def.fields.map((f) => f.name);
          return report;
        }
        log(`Creating list ${title}`, 'info');
        await this._sp.post('web/lists', { Title: title, Description: def.description, BaseTemplate: 100, AllowContentTypes: false, ContentTypesEnabled: false });
        report.created = true;
        report.exists = true;
        info = { EnableVersioning: false, ItemCount: 0, ReadSecurity: 1, WriteSecurity: 1, Hidden: false };
        if (def.titleDisplayName) await this._sp.merge(`${listPath}/fields/getbyinternalnameortitle('Title')`, { Title: def.titleDisplayName });
      }
      if (repair && (info.EnableVersioning !== !!def.versioning || info.Hidden !== true)) {
        // Lists are hidden from Site contents: people use the app (and Power BI / Power Automate use the lists).
        await this._sp.merge(listPath, { EnableVersioning: !!def.versioning, EnableAttachments: false, Hidden: true });
      }
      const existing = await this._sp.getAll<ISPField>(`${listPath}/fields?$select=InternalName,TypeAsString,Indexed,Choices&$top=500`);
      const byName: Record<string, ISPField> = {};
      existing.forEach((f) => (byName[f.InternalName] = f));
      for (const f of def.fields) {
        const current = byName[f.name];
        if (!current) {
          report.missingFields.push(f.name);
          if (repair) await this._createField(listPath, title, f, report, log);
          continue;
        }
        if (current.TypeAsString !== expectedTypeAsString(f)) {
          report.errors.push(`${f.name} has type ${current.TypeAsString}, expected ${expectedTypeAsString(f)}`);
          log(`${title}.${f.name} has type ${current.TypeAsString}, expected ${expectedTypeAsString(f)} — not changed`, 'warning');
          continue;
        }
        if (repair && f.indexed && !current.Indexed) {
          try {
            await this._sp.merge(`${listPath}/fields/getbyinternalnameortitle(${odataString(f.name)})`, { Indexed: true });
          } catch (e) {
            report.errors.push(`Index ${f.name}: ${(e as Error).message}`);
          }
        }
        if (repair && f.choices && f.type === 'Choice') {
          const have = current.Choices || [];
          const missing = f.choices.filter((c) => have.indexOf(c) < 0);
          if (missing.length) {
            await this._sp.postVerbose(`${listPath}/fields/getbyinternalnameortitle(${odataString(f.name)})`, { __metadata: { type: 'SP.FieldChoice' }, Choices: { results: have.concat(missing) } }, true);
          }
        }
      }
      log(`${title} ${report.created ? 'created' : 'checked'}`, 'success');
    } catch (e) {
      report.errors.push((e as Error).message);
      log(`${title}: ${(e as Error).message}`, 'error');
    }
    return report;
  }

  private async _createField(listPath: string, title: string, f: IFieldDef, report: IListReport, log: ProgressFn): Promise<void> {
    try {
      await this._sp.post(`${listPath}/fields/createfieldasxml`, { parameters: { SchemaXml: fieldSchemaXml(f), Options: FIELD_OPTIONS } });
      log(`${title}: added column ${f.title}`, 'info');
    } catch (e) {
      report.errors.push(`Create ${f.name}: ${(e as Error).message}`);
      log(`${title}: could not add ${f.name}: ${(e as Error).message}`, 'error');
    }
  }

  /** Delete every timesheet list (Admin › Schema & security › Uninstall). */
  public async removeAll(log: ProgressFn): Promise<void> {
    for (const def of LISTS) {
      const title = listTitle(this._prefix, def.key);
      try {
        await this._sp.del(this._sp.list(title));
        log(`Deleted ${title}`, 'success');
      } catch (e) {
        log(`${title}: ${(e as Error).message}`, 'warning');
      }
    }
  }

  // ------------------------------------------------------------------ security

  private async _roleIds(): Promise<Record<number, number>> {
    const roles = (await this._sp.get<{ value: Array<{ Id: number; RoleTypeKind: number }> }>('web/roledefinitions?$select=Id,RoleTypeKind')).value;
    const out: Record<number, number> = {};
    roles.forEach((r) => {
      if (r.RoleTypeKind && !out[r.RoleTypeKind]) out[r.RoleTypeKind] = r.Id;
    });
    return out;
  }

  private async _grant(listPath: string, principalId: number, roleId: number): Promise<void> {
    try {
      await this._sp.post(`${listPath}/roleassignments/getbyprincipalid(${principalId})/deleteobject()`);
    } catch {
      // had no assignment
    }
    await this._sp.post(`${listPath}/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleId})`);
  }

  /**
   * Apply the permission model:
   * - every list gets its own permissions (copied from the site, so owners keep Full Control);
   * - site groups that could edit are reduced to Read on configuration lists and to Contribute on timesheet lists;
   * - "Everyone except external users" and each external user: Read on configuration, projects and people;
   *   Contribute on timesheets (own rows only), notifications, tickets and saved reports;
   * - managers / approvers: Edit on projects, tasks, people and every timesheet (they see all rows);
   * - administrators: Edit on every list.
   */
  public async applySecurity(plan: ISecurityPlan, log: ProgressFn): Promise<string[]> {
    const problems: string[] = [];
    const ids = await this._roleIds();
    const owners = await this._sp.get<{ Id: number }>('web/associatedownergroup?$select=Id').catch(() => ({ Id: -1 }));
    const ensure = async (login: string): Promise<number | undefined> => {
      try {
        return await this._sp.ensureUser(login);
      } catch (e) {
        problems.push(`${login}: ${(e as Error).message}`);
        return undefined;
      }
    };
    const everyone = plan.tenantId ? await ensure(`c:0-.f|rolemanager|spo-grid-all-users/${plan.tenantId}`) : undefined;
    const toIds = async (list: string[]): Promise<number[]> => {
      const out: number[] = [];
      for (const e of list.filter((x, i, a) => !!x && a.map((y) => y.toLowerCase()).indexOf(x.toLowerCase()) === i)) {
        const id = await ensure(e.indexOf('|') >= 0 ? e : `i:0#.f|membership|${e}`);
        if (id) out.push(id);
      }
      return out;
    };
    const admins = await toIds(plan.admins);
    const managers = (await toIds(plan.managers)).filter((m) => admins.indexOf(m) < 0);
    const externals = (await toIds(plan.externals)).filter((m) => admins.indexOf(m) < 0 && managers.indexOf(m) < 0);

    for (const def of LISTS) {
      const title = listTitle(this._prefix, def.key);
      const listPath = this._sp.list(title);
      try {
        const sec = itemSecurity(def.security);
        await this._sp.merge(listPath, { ReadSecurity: sec, WriteSecurity: sec });
        const state = await this._sp.get<{ HasUniqueRoleAssignments: boolean }>(`${listPath}?$select=HasUniqueRoleAssignments`);
        if (!state.HasUniqueRoleAssignments) await this._sp.post(`${listPath}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`);
        const userRole = def.security === 'config' || def.security === 'managed' ? ids[READER] : ids[CONTRIBUTOR];
        // Site groups (members…) that could edit get the users' role: Edit would let them see everyone's timesheets.
        const assignments = (
          await this._sp.get<{ value: IRoleAssignment[] }>(`${listPath}/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Id,Member/PrincipalType,RoleDefinitionBindings/Id,RoleDefinitionBindings/RoleTypeKind`)
        ).value;
        for (const a of assignments) {
          if (a.PrincipalId === owners.Id) continue;
          const kinds = a.RoleDefinitionBindings.map((r) => r.RoleTypeKind);
          if (kinds.indexOf(ADMINISTRATOR) >= 0) continue;
          if (kinds.some((k) => k === EDITOR || k === CONTRIBUTOR || k === 4 /* web designer */)) await this._grant(listPath, a.PrincipalId, userRole);
        }
        if (everyone) await this._grant(listPath, everyone, def.security === 'adminOnly' ? ids[CONTRIBUTOR] : userRole);
        for (const x of externals) await this._grant(listPath, x, userRole);
        const managerRole = def.security === 'config' || def.security === 'adminOnly' ? userRole : ids[EDITOR];
        for (const m of managers) await this._grant(listPath, m, managerRole);
        for (const a of admins) await this._grant(listPath, a, ids[EDITOR]);
        log(`${title}: permissions applied`, 'success');
      } catch (e) {
        problems.push(`${title}: ${(e as Error).message}`);
        log(`${title}: ${(e as Error).message}`, 'error');
      }
    }
    return problems;
  }
}
