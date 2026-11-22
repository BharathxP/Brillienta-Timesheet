import { SPClient, SPError, odataString } from './SPClient';
import { listTitle } from '../schema/schema';

export const SETTING_KEYS = {
  config: 'Config',
  schemaVersion: 'SchemaVersion',
  appVersion: 'AppVersion',
  installedAt: 'InstalledAt',
  installedBy: 'InstalledBy',
  security: 'Security',
  licenseKey: 'LicenseKey',
  brandLogo: 'BrandLogo',
  pageUrl: 'PageUrl',
  /** Last reminder / project notice run: "<ISO time>|<name>" (one browser runs them at a time). */
  reminderRun: 'ReminderRun',
  /** Weeks already reminded: JSON { "<email>|<week>": "<ISO>" } (pruned to the lookback). */
  reminderLog: 'ReminderLog'
};

export interface ISecurityMeta {
  appliedAt?: string;
  by?: string;
}

interface ISettingRow {
  Id: number;
  Title: string;
  SettingValue: string | null;
}

/** Key/value rows of the <prefix>Settings list. */
export class SettingsStore {
  private readonly _sp: SPClient;
  private readonly _title: string;
  private _rows: Record<string, ISettingRow> = {};

  public constructor(sp: SPClient, prefix: string) {
    this._sp = sp;
    this._title = listTitle(prefix, 'settings');
  }

  /** Loads all rows. Resolves false when the Settings list does not exist (not installed yet). */
  public async load(): Promise<boolean> {
    try {
      const rows = await this._sp.getAll<ISettingRow>(`${this._sp.list(this._title)}/items?$select=Id,Title,SettingValue&$top=500`);
      const map: Record<string, ISettingRow> = {};
      rows.forEach((r) => (map[r.Title] = r));
      this._rows = map;
      return true;
    } catch (e) {
      if (e instanceof SPError && e.status === 404) return false;
      throw e;
    }
  }

  public get(key: string): string | undefined {
    const row = this._rows[key];
    return row && row.SettingValue !== null ? row.SettingValue : undefined;
  }

  public getJson<T>(key: string): T | undefined {
    const raw = this.get(key);
    if (!raw) return undefined;
    try {
      return JSON.parse(raw) as T;
    } catch {
      return undefined;
    }
  }

  public async set(key: string, value: string): Promise<void> {
    const existing = this._rows[key];
    if (existing) {
      await this._sp.merge(`${this._sp.list(this._title)}/items(${existing.Id})`, { SettingValue: value });
      existing.SettingValue = value;
    } else {
      const created = await this._sp.post<ISettingRow>(`${this._sp.list(this._title)}/items`, { Title: key, SettingValue: value });
      this._rows[key] = { Id: created.Id, Title: key, SettingValue: value };
    }
  }

  public async setJson(key: string, value: unknown): Promise<void> {
    await this.set(key, JSON.stringify(value));
  }

  /** Re-read one key straight from SharePoint. */
  public async fresh(key: string): Promise<string | undefined> {
    const rows = await this._sp.get<{ value: ISettingRow[] }>(`${this._sp.list(this._title)}/items?$select=Id,Title,SettingValue&$filter=Title eq ${odataString(key)}&$top=1`);
    const row = rows.value[0];
    if (row) this._rows[key] = row;
    return row && row.SettingValue !== null ? row.SettingValue : undefined;
  }
}
