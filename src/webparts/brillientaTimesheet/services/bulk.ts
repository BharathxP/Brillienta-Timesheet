// Bulk project / task / activity creation: parse rows pasted from Excel (tab separated) or a CSV file into
// validated records. Pure functions (tests/bulk.test.cjs).

import { isKey } from './dates';

export type BulkKind = 'projects' | 'tasks' | 'activities';

export interface IBulkColumn {
  key: string;
  label: string;
  required?: boolean;
  type?: 'text' | 'number' | 'boolean' | 'date' | 'emails';
  hint?: string;
}

export const BULK_COLUMNS: Record<BulkKind, IBulkColumn[]> = {
  projects: [
    { key: 'title', label: 'Project', required: true },
    { key: 'code', label: 'Code' },
    { key: 'client', label: 'Client' },
    { key: 'managers', label: 'Managers', type: 'emails', hint: 'emails separated by ;' },
    { key: 'members', label: 'Members', type: 'emails', hint: 'emails separated by ; (empty = everyone)' },
    { key: 'billable', label: 'Billable', type: 'boolean', hint: 'yes / no' },
    { key: 'start', label: 'Start date', type: 'date', hint: 'yyyy-mm-dd' },
    { key: 'end', label: 'End date', type: 'date', hint: 'yyyy-mm-dd' },
    { key: 'planned', label: 'Planned hours', type: 'number' },
    { key: 'cost', label: 'Hourly cost', type: 'number' },
    { key: 'rate', label: 'Bill rate', type: 'number' },
    { key: 'program', label: 'Programme' }
  ],
  tasks: [
    { key: 'project', label: 'Project', required: true, hint: 'project name or code' },
    { key: 'title', label: 'Task', required: true },
    { key: 'assignees', label: 'Assignees', type: 'emails', hint: 'emails separated by ;' },
    { key: 'billable', label: 'Billable', type: 'boolean' },
    { key: 'planned', label: 'Planned hours', type: 'number' },
    { key: 'due', label: 'Due date', type: 'date' },
    { key: 'selfAssign', label: 'Self-assign', type: 'boolean' }
  ],
  activities: [
    { key: 'title', label: 'Activity', required: true },
    { key: 'billable', label: 'Billable', type: 'boolean' }
  ]
};

export interface IBulkRow {
  line: number;
  values: Record<string, string | number | boolean | string[]>;
  errors: string[];
}

/** Split pasted text into cells: tabs (Excel) when present, else CSV with quotes. */
export function splitRows(text: string): string[][] {
  const lines = (text || '').replace(/\r\n?/g, '\n').split('\n').filter((l) => l.trim());
  const tabbed = lines.some((l) => l.indexOf('\t') >= 0);
  return lines.map((l) => (tabbed ? l.split('\t') : splitCsvLine(l)).map((c) => c.trim()));
}

function splitCsvLine(line: string): string[] {
  const out: string[] = [];
  let cur = '';
  let q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (q) {
      if (ch === '"' && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else if (ch === '"') q = false;
      else cur += ch;
    } else if (ch === '"') q = true;
    else if (ch === ',' || ch === ';') {
      // ';' separates cells only when the line has no commas (European Excel CSV)
      if (ch === ';' && line.indexOf(',') >= 0) cur += ch;
      else {
        out.push(cur);
        cur = '';
      }
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

const YES = ['yes', 'y', 'true', '1', 'billable', 'x'];
const NO = ['no', 'n', 'false', '0', 'non-billable', 'nonbillable', ''];

/** Parse rows for a kind. A first row that matches the column labels is treated as a header (and may reorder). */
export function parseBulk(kind: BulkKind, text: string): IBulkRow[] {
  const cols = BULK_COLUMNS[kind];
  let rows = splitRows(text);
  let order = cols.map((c) => c.key);
  if (rows.length) {
    const head = rows[0].map((h) => h.toLowerCase().replace(/\*$/, '').trim());
    const matched = head.map((h) => {
      const c = cols.filter((x) => x.label.toLowerCase() === h || x.key.toLowerCase() === h)[0];
      return c ? c.key : '';
    });
    if (matched.filter((m) => !!m).length >= Math.min(2, cols.length)) {
      order = matched;
      rows = rows.slice(1);
    }
  }
  return rows.map((cells, i) => {
    const values: IBulkRow['values'] = {};
    const errors: string[] = [];
    order.forEach((key, j) => {
      const col = cols.filter((c) => c.key === key)[0];
      if (!col) return;
      const raw = (cells[j] || '').trim();
      switch (col.type) {
        case 'number': {
          // Thousands separators and currency symbols are fine ("₹1,500"); anything else is not a number.
          const n = raw ? Number(raw.replace(/[\s,₹$€£]/g, '')) : 0;
          if (raw && isNaN(n)) errors.push(`${col.label} is not a number`);
          values[key] = isNaN(n) ? 0 : n;
          break;
        }
        case 'boolean': {
          const v = raw.toLowerCase();
          if (YES.indexOf(v) >= 0) values[key] = true;
          else if (NO.indexOf(v) >= 0) values[key] = raw ? false : '';
          else errors.push(`${col.label} should be yes or no`);
          break;
        }
        case 'date':
          if (raw && !isKey(raw)) errors.push(`${col.label} should be yyyy-mm-dd`);
          values[key] = isKey(raw) ? raw : '';
          break;
        case 'emails': {
          const list = raw
            .split(/[;,\s]+/)
            .map((s) => s.trim())
            .filter((s) => !!s);
          const bad = list.filter((s) => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s));
          if (bad.length) errors.push(`Not an email: ${bad.join(', ')}`);
          values[key] = list.filter((s) => bad.indexOf(s) < 0);
          break;
        }
        default:
          values[key] = raw.substring(0, 255);
      }
    });
    cols.filter((c) => c.required).forEach((c) => {
      if (!values[c.key]) errors.push(`${c.label} is required`);
    });
    return { line: i + 1, values, errors };
  });
}

/** Header line for a template file the user can fill in Excel. */
export function bulkTemplate(kind: BulkKind): string {
  return BULK_COLUMNS[kind].map((c) => c.label + (c.required ? '*' : '')).join('\t');
}
