// Time off Manager 365 integration: reads approved leave from its SharePoint list (on this site or another site of
// the tenant) with the column names an admin mapped in Admin › Integrations. The timesheet shows leave days, counts
// them towards daily and weekly minimums and can put them on the sheet on the time-off activity.

import { DateKey, ITimeOff, ITimeOffConfig } from '../models/types';
import { SPClient, odataString } from './SPClient';
import { dateOnly } from './mappers';

export class TimeOffService {
  private readonly _sp: SPClient;

  public constructor(sp: SPClient) {
    this._sp = sp;
  }

  private _client(cfg: ITimeOffConfig): SPClient {
    return cfg.siteUrl ? this._sp.forWeb(cfg.siteUrl) : this._sp;
  }

  /** Approved leave overlapping [from, to] (for one person, or everyone when email is empty). */
  public async load(cfg: ITimeOffConfig, from: DateKey, to: DateKey, email?: string): Promise<ITimeOff[]> {
    if (!cfg.enabled || !cfg.listTitle) return [];
    const sp = this._client(cfg);
    const f = [
      `${cfg.startField} le datetime'${to}T23:59:59Z'`,
      `${cfg.endField} ge datetime'${from}T00:00:00Z'`,
      cfg.statusField ? `${cfg.statusField} eq ${odataString(cfg.approvedValue)}` : '',
      email && cfg.emailField ? `${cfg.emailField} eq ${odataString(email)}` : ''
    ]
      .filter((x) => !!x)
      .join(' and ');
    const fields = [cfg.emailField, cfg.startField, cfg.endField, cfg.typeField].filter((x) => !!x);
    const rows = await sp.getAll<Record<string, unknown>>(`${sp.list(cfg.listTitle)}/items?$select=${fields.join(',')}&$filter=${encodeURIComponent(f)}&$top=2000`, 10000);
    return rows
      .map((r) => ({
        email: String(r[cfg.emailField] || ''),
        start: dateOnly(r[cfg.startField]),
        end: dateOnly(r[cfg.endField]),
        type: cfg.typeField ? String(r[cfg.typeField] || 'Time off') : 'Time off'
      }))
      .filter((l) => !!l.email && !!l.start && !!l.end);
  }

  /** Check the mapping: the list exists and has the columns. Returns problems ('' = fine). */
  public async test(cfg: ITimeOffConfig): Promise<string[]> {
    const sp = this._client(cfg);
    try {
      const fields = await sp.getAll<{ InternalName: string }>(`${sp.list(cfg.listTitle)}/fields?$select=InternalName&$top=500`);
      const names = fields.map((f) => f.InternalName);
      return [cfg.emailField, cfg.startField, cfg.endField, cfg.statusField, cfg.typeField].filter((f) => !!f && names.indexOf(f) < 0).map((f) => `The list has no column with the internal name "${f}".`);
    } catch (e) {
      return [`The list "${cfg.listTitle}" could not be read: ${(e as Error).message}`];
    }
  }
}
