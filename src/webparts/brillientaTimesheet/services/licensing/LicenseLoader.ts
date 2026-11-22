// Reads what the licence status depends on: the stored key, the SharePoint server's time and the trial start (the
// earliest of the stored install date and the creation dates of this site's timesheet lists).

import { SPClient } from '../SPClient';
import { SETTING_KEYS, SettingsStore } from '../SettingsStore';
import { ILicenseStatus, checkLicenseKey, earliest, licenseStatus } from './license';
import { LICENSE_PUBLIC_KEY } from './publicKey';

export async function trialStart(sp: SPClient, settings: SettingsStore, prefix: string, now: Date): Promise<Date> {
  const dates: Array<string | undefined> = [settings.get(SETTING_KEYS.installedAt)];
  try {
    const res = await sp.get<{ value: Array<{ Title: string; Created: string }> }>(`web/lists?$select=Title,Created&$filter=startswith(Title,'${prefix.replace(/'/g, "''")}')`);
    res.value.filter((l) => l.Title.indexOf(prefix) === 0).forEach((l) => dates.push(l.Created));
  } catch {
    // keep the stored date
  }
  return earliest(dates, now);
}

export async function loadLicenseStatus(sp: SPClient, settings: SettingsStore, prefix: string, tenantId: string | undefined): Promise<ILicenseStatus> {
  const now = await sp.serverTime();
  const start = await trialStart(sp, settings, prefix, now);
  const key = settings.get(SETTING_KEYS.licenseKey) || '';
  if (!key) return licenseStatus(undefined, start, now);
  try {
    return licenseStatus(await checkLicenseKey(key, tenantId, LICENSE_PUBLIC_KEY), start, now);
  } catch (e) {
    return licenseStatus(undefined, start, now, (e as Error).message);
  }
}
