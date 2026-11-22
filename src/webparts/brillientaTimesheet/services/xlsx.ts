// Excel export (.xlsx, Office Open XML) without a library: one workbook, one or more sheets. Pure (unit tested).
//
// Text goes in as inline strings, never formulas, so a value such as "=HYPERLINK(...)" typed into a profile stays
// text when the file is opened. Numbers stay numbers; ISO dates ("2026-03-14", "2026-03-14T09:30:00Z") become real
// Excel dates. The header row is bold, frozen and filterable; column widths follow the content.

import { zipFiles } from './zip';

export type XlsxValue = string | number | boolean | null | undefined;

export interface IXlsxSheet {
  name: string;
  columns: string[];
  rows: XlsxValue[][];
}

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?$/;

/** Text safe inside XML: escaped, and without the control characters XML 1.0 does not allow. */
function xmlText(v: string): string {
  return (
    v
      // eslint-disable-next-line no-control-regex
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
  );
}

/** Column letters for a 0-based index (0 → A, 26 → AA). */
export function columnName(i: number): string {
  let n = i + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Excel's serial date (days since 1899-12-30, the 1900 date system) for an ISO date, or undefined. */
export function excelDate(v: string): { serial: number; time: boolean } | undefined {
  const m = ISO_DATE.exec(v);
  if (!m) return undefined;
  const hasTime = m[4] !== undefined;
  let ms: number;
  if (hasTime && m[7]) {
    const t = Date.parse(v);
    if (isNaN(t)) return undefined;
    // An instant with a zone: shown in the reader's local time, like the app does.
    const d = new Date(t);
    ms = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds());
  } else {
    ms = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), hasTime ? Number(m[4]) : 0, hasTime ? Number(m[5]) : 0, m[6] ? Number(m[6]) : 0);
  }
  const date = new Date(ms);
  if (isNaN(ms) || date.getUTCMonth() !== Number(m[2]) - 1) return undefined; // 2026-02-30 and the like
  const serial = ms / 86400000 + 25569;
  if (serial < 1) return undefined;
  return { serial: Math.round(serial * 86400) / 86400, time: hasTime && (date.getUTCHours() !== 0 || date.getUTCMinutes() !== 0 || date.getUTCSeconds() !== 0) };
}

/** Sheet names: at most 31 characters, none of : \ / ? * [ ], unique (case-insensitive), never empty. */
export function sheetNames(names: string[]): string[] {
  const used: Record<string, boolean> = {};
  return names.map((raw, i) => {
    const base = (raw || '').replace(/[:\\/?*[\]]/g, ' ').replace(/^'+|'+$/g, '').replace(/\s+/g, ' ').trim().substring(0, 31) || `Sheet ${i + 1}`;
    let name = base;
    let n = 2;
    while (used[name.toLowerCase()]) {
      const suffix = ` (${n++})`;
      name = base.substring(0, 31 - suffix.length) + suffix;
    }
    used[name.toLowerCase()] = true;
    return name;
  });
}

const STYLE = { header: 1, date: 2, dateTime: 3 };

function cell(ref: string, v: XlsxValue): string {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return isFinite(v) ? `<c r="${ref}"><v>${v}</v></c>` : '';
  if (typeof v === 'boolean') return `<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>`;
  const d = excelDate(v);
  if (d) return `<c r="${ref}" s="${d.time ? STYLE.dateTime : STYLE.date}"><v>${d.serial}</v></c>`;
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlText(v)}</t></is></c>`;
}

function sheetXml(s: IXlsxSheet): string {
  const cols = s.columns.length;
  const widths = s.columns.map((c, i) => {
    let w = c.length + 2;
    for (let r = 0; r < Math.min(s.rows.length, 500); r++) {
      const v = s.rows[r][i];
      const len = v === null || v === undefined ? 0 : typeof v === 'string' ? (excelDate(v) ? 16 : Math.min(v.length, 80)) : String(v).length;
      if (len + 2 > w) w = len + 2;
    }
    return Math.min(Math.max(w, 8), 60);
  });
  const header = `<row r="1">${s.columns.map((c, i) => `<c r="${columnName(i)}1" t="inlineStr" s="${STYLE.header}"><is><t xml:space="preserve">${xmlText(c)}</t></is></c>`).join('')}</row>`;
  const body = s.rows.map((row, r) => `<row r="${r + 2}">${s.columns.map((_, i) => cell(`${columnName(i)}${r + 2}`, row[i])).join('')}</row>`).join('');
  const last = `${columnName(Math.max(cols, 1) - 1)}${s.rows.length + 1}`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<dimension ref="A1:${last}"/>` +
    '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
    `<cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>` +
    `<sheetData>${header}${body}</sheetData>` +
    (cols ? `<autoFilter ref="A1:${last}"/>` : '') +
    '</worksheet>'
  );
}

/** The workbook's files (path → XML). */
export function xlsxFiles(sheets: IXlsxSheet[]): Array<{ path: string; content: string }> {
  const list = sheets.length ? sheets : [{ name: 'Sheet 1', columns: [], rows: [] }];
  const names = sheetNames(list.map((s) => s.name));
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  return [
    {
      path: '[Content_Types].xml',
      content:
        head +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        list.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
        '</Types>'
    },
    {
      path: '_rels/.rels',
      content:
        head +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'
    },
    {
      path: 'xl/workbook.xml',
      content:
        head +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        `<sheets>${names.map((n, i) => `<sheet name="${xmlText(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>` +
        `<definedNames>${list
          .map((s, i) => (s.columns.length ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${xmlText(names[i].replace(/'/g, "''"))}'!$A$1:$${columnName(s.columns.length - 1)}$${s.rows.length + 1}</definedName>` : ''))
          .join('')}</definedNames>` +
        '</workbook>'
    },
    {
      path: 'xl/_rels/workbook.xml.rels',
      content:
        head +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        list.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
        `<Relationship Id="rId${list.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
        '</Relationships>'
    },
    {
      path: 'xl/styles.xml',
      content:
        head +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm"/></numFmts>' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="4">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
        '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
        '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
        '</cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        '</styleSheet>'
    },
    ...list.map((s, i) => ({ path: `xl/worksheets/sheet${i + 1}.xml`, content: sheetXml(s) }))
  ];
}

/** The workbook as .xlsx bytes. */
export function xlsxWorkbook(sheets: IXlsxSheet[]): Uint8Array {
  return zipFiles(xlsxFiles(sheets));
}
