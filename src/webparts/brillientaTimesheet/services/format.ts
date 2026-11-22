// Number, money and CSV formatting. Pure functions.

export function hours(n: number, digits: number = 2): string {
  const v = Math.round((n || 0) * 100) / 100;
  return v.toFixed(digits).replace(/\.?0+$/, '') || '0';
}

/** "7:30" style display of decimal hours. */
export function hhmm(n: number): string {
  const total = Math.round((n || 0) * 60);
  return `${Math.floor(total / 60)}:${('0' + (total % 60)).slice(-2)}`;
}

const SYMBOLS: Record<string, string> = { INR: '₹', USD: '$', EUR: '€', GBP: '£', AUD: 'A$', CAD: 'C$', SGD: 'S$', AED: 'AED ', JPY: '¥' };

export const CURRENCIES = Object.keys(SYMBOLS);

export function money(n: number, currency: string): string {
  const sym = SYMBOLS[currency] !== undefined ? SYMBOLS[currency] : currency + ' ';
  const v = Math.round((n || 0) * 100) / 100;
  const s = Math.abs(v).toLocaleString(currency === 'INR' ? 'en-IN' : 'en-GB', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  return (v < 0 ? '−' : '') + sym + s;
}

export function percent(n: number): string {
  return `${Math.round(n * 10) / 10}%`;
}

export function initials(name: string): string {
  const parts = (name || '?').replace(/\(.*?\)/g, '').trim().split(/\s+/);
  return ((parts[0] || '?').charAt(0) + (parts.length > 1 ? parts[parts.length - 1].charAt(0) : '')).toUpperCase();
}

export function formatDateTime(iso: string | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleString(undefined, { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
}

export function relativeTime(iso: string | undefined): string {
  if (!iso) return '';
  const diff = Date.now() - new Date(iso).getTime();
  const m = Math.round(diff / 60000);
  if (m < 1) return 'just now';
  if (m < 60) return `${m} min ago`;
  const h = Math.round(m / 60);
  if (h < 24) return `${h} h ago`;
  const d = Math.round(h / 24);
  if (d < 30) return `${d} day${d === 1 ? '' : 's'} ago`;
  return formatDateTime(iso);
}

/** CSV with formula-injection protection (cells starting with = + - @ are prefixed with '). */
export function toCsv(rows: Array<Array<string | number | boolean | undefined>>): string {
  const cell = (v: string | number | boolean | undefined): string => {
    let s = v === undefined || v === null ? '' : String(v);
    if (typeof v === 'string' && /^[=+\-@\t\r]/.test(s)) s = "'" + s;
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return '﻿' + rows.map((r) => r.map(cell).join(',')).join('\r\n');
}

/** Save bytes or text as a file in the browser. */
export function download(name: string, content: string | Uint8Array, type: string): void {
  const blob = new Blob([content as BlobPart], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.setAttribute('data-interception', 'off');
  document.body.appendChild(a);
  a.click();
  window.setTimeout(() => {
    URL.revokeObjectURL(url);
    a.remove();
  }, 1000);
}

export function slug(s: string): string {
  return (s || 'export').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '').toLowerCase() || 'export';
}
