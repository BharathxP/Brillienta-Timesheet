// Licence keys and the 14-day trial (same scheme as Brillienta Employee Directory).
//
// A key is signed by Brillienta with a private key that never leaves the vendor (tools/license/license.cjs); the app
// holds only the public key, so a key cannot be made up or edited. Each key names one Microsoft 365 tenant, and
// "product":"timesheet" — an Employee Directory key does not unlock Timesheet, nor the other way round.
//
// The trial runs 14 days from installation: the earliest of the stored install date and the creation date of the
// Settings list (set by SharePoint, not editable). "Now" comes from the SharePoint server, not the PC clock.
//
// Format: BT1.<base64url(payload JSON)>.<base64url(ECDSA P-256 / SHA-256 signature, IEEE P1363)>

import { ADDONS, Addon, Plan, planInfo, userLimit } from './plans';

export const TRIAL_DAYS = 14;
export const GRACE_DAYS = 14;
export const KEY_PREFIX = 'BT1';
const DAY = 86400000;
const PLANS: Plan[] = ['standard', 'pro', 'enterprise'];

export interface ILicensePayload {
  v: number;
  product: 'timesheet';
  id: string;
  plan: Plan;
  tenant: string;
  customer: string;
  /** Users included (0 = the plan's standard number). */
  users: number;
  issued: string;
  expires?: string;
  addons?: string[];
  /** Enterprise: the customer's dedicated account manager. */
  accountManager?: { name: string; email: string; phone?: string };
}

export type LicenseState = 'trial' | 'licensed' | 'grace' | 'trialEnded' | 'expired';

export interface ILicenseStatus {
  state: LicenseState;
  plan?: Plan;
  users: number;
  trialStart: string;
  trialEnds: string;
  daysLeft?: number;
  license?: ILicensePayload;
  keyProblem?: string;
  addons: Addon[];
}

function fromB64url(s: string): Uint8Array {
  const b64 = s.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((s.length + 3) % 4);
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

const knownAddons = (list: unknown): Addon[] => (Array.isArray(list) ? (list.filter((a) => ADDONS.some((x) => x.addon === a)) as Addon[]) : []);

export interface IParsedKey {
  payload: ILicensePayload;
  signed: string;
  signature: Uint8Array;
}

/** Split and decode a key (no signature check). Throws a readable error. */
export function parseLicenseKey(key: string): IParsedKey {
  const clean = (key || '').replace(/\s+/g, '');
  const parts = clean.split('.');
  if (parts.length === 3 && parts[0] === 'ED1') throw new Error('This is an Employee Directory licence key. Brillienta Timesheet keys start with BT1.');
  if (parts.length !== 3 || parts[0] !== KEY_PREFIX) throw new Error('This is not a Brillienta Timesheet licence key.');
  let payload: ILicensePayload;
  try {
    payload = JSON.parse(new TextDecoder().decode(fromB64url(parts[1])));
  } catch {
    throw new Error('The licence key is damaged. Copy it again, in full.');
  }
  if (!payload || payload.v !== 1 || payload.product !== 'timesheet' || PLANS.indexOf(payload.plan) < 0 || !payload.tenant || !payload.id) {
    throw new Error('The licence key is not valid for this version.');
  }
  return { payload, signed: parts[0] + '.' + parts[1], signature: fromB64url(parts[2]) };
}

export async function verifySignature(publicKeySpkiB64: string, signed: string, signature: Uint8Array, subtle: SubtleCrypto = crypto.subtle): Promise<boolean> {
  const spki = Uint8Array.from(atob(publicKeySpkiB64), (c) => c.charCodeAt(0));
  const key = await subtle.importKey('spki', spki, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['verify']);
  const data = new Uint8Array(signed.length);
  for (let i = 0; i < signed.length; i++) data[i] = signed.charCodeAt(i);
  return subtle.verify({ name: 'ECDSA', hash: 'SHA-256' }, key, signature, data);
}

/** Verify a key's signature and tenant; returns the payload or throws the reason. */
export async function checkLicenseKey(key: string, tenantId: string | undefined, publicKey: string, subtle?: SubtleCrypto): Promise<ILicensePayload> {
  const parsed = parseLicenseKey(key);
  let ok = false;
  try {
    ok = await verifySignature(publicKey, parsed.signed, parsed.signature, subtle);
  } catch {
    ok = false;
  }
  if (!ok) throw new Error('The licence key is not genuine (signature check failed).');
  if (!tenantId) throw new Error('This site did not report its Microsoft 365 tenant id, so the licence cannot be checked.');
  if (parsed.payload.tenant.toLowerCase() !== tenantId.toLowerCase()) throw new Error('This licence key was issued for another Microsoft 365 organisation.');
  return parsed.payload;
}

function endOf(expires?: string): number | undefined {
  if (!expires) return undefined;
  const t = Date.parse(expires + 'T23:59:59Z');
  return isNaN(t) ? undefined : t;
}

/** Earliest plausible date among the candidates (ISO strings). */
export function earliest(dates: Array<string | undefined>, fallback: Date): Date {
  const times = dates.map((d) => (d ? Date.parse(d) : NaN)).filter((t) => !isNaN(t) && t > Date.UTC(2020, 0, 1));
  return times.length ? new Date(Math.min(...times, fallback.getTime())) : fallback;
}

/** Licence status from a verified payload (or none) and the trial start. A valid licence wins over the trial. */
export function licenseStatus(payload: ILicensePayload | undefined, trialStart: Date, now: Date, keyProblem?: string): ILicenseStatus {
  const trialEnd = trialStart.getTime() + TRIAL_DAYS * DAY;
  const base = { trialStart: trialStart.toISOString(), trialEnds: new Date(trialEnd).toISOString(), keyProblem };
  const daysUntil = (t: number): number => Math.max(0, Math.ceil((t - now.getTime()) / DAY));
  if (payload) {
    const end = endOf(payload.expires);
    const users = userLimit(payload.plan, payload.users);
    const addons = knownAddons(payload.addons);
    if (end === undefined || now.getTime() <= end) {
      return { ...base, state: 'licensed', plan: payload.plan, users, license: payload, addons, daysLeft: end === undefined ? undefined : daysUntil(end) };
    }
    if (now.getTime() <= end + GRACE_DAYS * DAY) return { ...base, state: 'grace', plan: payload.plan, users, license: payload, addons, daysLeft: daysUntil(end + GRACE_DAYS * DAY) };
    return { ...base, state: 'expired', users: 0, license: payload, addons: [], daysLeft: 0 };
  }
  if (now.getTime() < trialEnd) return { ...base, state: 'trial', plan: 'enterprise', users: planInfo('enterprise').users, addons: ADDONS.map((a) => a.addon), daysLeft: daysUntil(trialEnd) };
  return { ...base, state: 'trialEnded', users: 0, addons: [], daysLeft: 0 };
}

export function describeStatus(s: ILicenseStatus): string {
  const plan = s.plan ? planInfo(s.plan).name : '';
  switch (s.state) {
    case 'trial':
      return `Trial (every Enterprise feature): ${s.daysLeft} day${s.daysLeft === 1 ? '' : 's'} left`;
    case 'licensed':
      return s.daysLeft === undefined ? `${plan} plan` : `${plan} plan until ${s.license && s.license.expires}`;
    case 'grace':
      return `${plan} licence ended on ${s.license && s.license.expires}; the app stops in ${s.daysLeft} day${s.daysLeft === 1 ? '' : 's'}`;
    case 'trialEnded':
      return 'The 14-day trial has ended';
    default:
      return 'The licence has ended';
  }
}
