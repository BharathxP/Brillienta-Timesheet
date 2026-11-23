// In-memory stand-in for the SharePoint REST API, for the local preview only (npm run preview).
// Implements just what the directory calls: current user, list items (read / create / merge / delete with simple
// $filter, $orderby, $top), list change dates and $batch.

const lists = {};
let nextId = 1;
const nowIso = () => new Date().toISOString();

// Date-only columns: like SharePoint, keep only the calendar day and return it as midnight in the site's time zone
// (this stand-in's site runs on UTC), not the instant that was written.
const DATE_ONLY = ['StartDate', 'EndDate', 'DueDate'];
function stored(data) {
  const out = { ...data };
  DATE_ONLY.forEach((k) => {
    if (typeof out[k] === 'string' && out[k]) {
      const d = new Date(out[k]);
      if (!isNaN(d.getTime())) out[k] = d.toISOString().substring(0, 10) + 'T00:00:00Z';
    }
  });
  return out;
}

function list(title) {
  if (!lists[title]) lists[title] = { items: [], modified: nowIso() };
  return lists[title];
}

/** The in-memory lists, for browser tests. */
export function fakeLists() {
  return lists;
}

export function seed(title, rows, author) {
  const l = list(title);
  rows.forEach((r) => l.items.push({ Id: nextId++, Created: r.Created || nowIso(), Modified: nowIso(), Author: author, AuthorId: author.Id, ...stored(r) }));
}

function response(status, body, binary) {
  const text = body === undefined ? '' : typeof body === 'string' ? body : JSON.stringify(body);
  return { ok: status >= 200 && status < 300, status, statusText: String(status), text: async () => text, json: async () => JSON.parse(text), arrayBuffer: async () => binary || new ArrayBuffer(0) };
}

const toB64 = (body) => {
  const bytes = new Uint8Array(body instanceof ArrayBuffer ? body : body && body.buffer ? body.buffer : new ArrayBuffer(0));
  let bin = '';
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
};
const fromB64 = (b64) => {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out.buffer;
};

// A small OData $filter: eq/ne/lt/le/gt/ge on strings, numbers, true/false, null and datetime'…', joined by
// and/or with parentheses (what the directory and Calendar View send).
function parseFilter(expr) {
  if (!expr) return () => true;
  const src = decodeURIComponent(expr);
  const tokens = [];
  const re = /\s*(\(|\)|'(?:[^']|'')*'|datetime'[^']*'|[\w/.:-]+)/gy;
  let m;
  while ((m = re.exec(src))) tokens.push(m[1]);
  let i = 0;
  const literal = (t) => {
    if (/^datetime'/.test(t)) return { date: Date.parse(t.slice(9, -1)) };
    if (/^'/.test(t)) return t.slice(1, -1).replace(/''/g, "'");
    if (t === 'true' || t === 'false') return t === 'true';
    if (t === 'null') return null;
    return /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : t;
  };
  const compare = (field, op, value) => (row) => {
    const v = field.split('/').reduce((o, k) => (o == null ? o : o[k]), row);
    if (value && typeof value === 'object' && 'date' in value) {
      const x = Date.parse(v);
      if (isNaN(x)) return op === 'ne';
      return { eq: x === value.date, ne: x !== value.date, lt: x < value.date, le: x <= value.date, gt: x > value.date, ge: x >= value.date }[op];
    }
    if (value === null) return op === 'eq' ? v == null || v === '' : !(v == null || v === '');
    if (op === 'eq' || op === 'ne') {
      const same = String(v) === String(value) || (typeof value === 'string' && String(v).toLowerCase() === value.toLowerCase()) || (typeof value === 'boolean' && !!v === value);
      return op === 'eq' ? same : !same;
    }
    const a = typeof value === 'number' ? Number(v) : String(v);
    return { lt: a < value, le: a <= value, gt: a > value, ge: a >= value }[op];
  };
  const primary = () => {
    if (tokens[i] === '(') {
      i++;
      const e = orExpr();
      i++; // ')'
      return e;
    }
    const field = tokens[i++];
    const op = (tokens[i++] || '').toLowerCase();
    const value = literal(tokens[i++] || '');
    return ['eq', 'ne', 'lt', 'le', 'gt', 'ge'].indexOf(op) >= 0 ? compare(field, op, value) : () => true;
  };
  const andExpr = () => {
    const parts = [primary()];
    while ((tokens[i] || '').toLowerCase() === 'and') {
      i++;
      parts.push(primary());
    }
    return (row) => parts.every((t) => t(row));
  };
  const orExpr = () => {
    const parts = [andExpr()];
    while ((tokens[i] || '').toLowerCase() === 'or') {
      i++;
      parts.push(andExpr());
    }
    return (row) => parts.some((t) => t(row));
  };
  return orExpr();
}

function query(url) {
  const q = {};
  const i = url.indexOf('?');
  if (i < 0) return q;
  url
    .substring(i + 1)
    .split('&')
    .forEach((pair) => {
      const j = pair.indexOf('=');
      q[decodeURIComponent(pair.substring(0, j))] = pair.substring(j + 1);
    });
  return q;
}

const titleOf = (url) => {
  const m = /getbytitle\('([^']+)'\)/.exec(decodeURIComponent(url));
  return m ? m[1] : undefined;
};
const itemIdOf = (url) => {
  const m = /\/items\((\d+)\)/.exec(url);
  return m ? Number(m[1]) : undefined;
};

export function createFakeHttp(user) {
  const handle = (method, url, body, headers) => {
    const path = url.replace(/^https?:\/\/[^/]+/, '');
    const verb = (headers && (headers['X-HTTP-Method'] || headers['x-http-method'])) || method;
    if (/\/_api\/\$batch/.test(path)) return batch(body);
    if (/web\/currentuser\/groups/.test(path)) return response(200, { value: [] });
    if (/web\/currentuser/.test(path)) return response(200, user);
    if (/clientPeoplePickerSearchUser/.test(path)) return response(200, { value: '[]' });
    if (/web\/ensureuser/.test(path)) {
      // A stable id per person (window.__edPrincipals maps it back), so item permissions can be checked.
      const login = (body ? JSON.parse(body).logonName : '') || '';
      window.__edPrincipals = window.__edPrincipals || {};
      const known = Object.keys(window.__edPrincipals).filter((k) => window.__edPrincipals[k] === login)[0];
      const pid = known ? Number(known) : 1000 + Object.keys(window.__edPrincipals).length;
      window.__edPrincipals[pid] = login;
      return response(200, { Id: pid });
    }
    if (/web\/RegionalSettings/i.test(path)) return response(200, { FirstDayOfWeek: 1 });
    if (/web\/roledefinitions/.test(path) && method === 'GET') {
      return response(200, { value: [{ Id: 1073741829, Name: 'Full Control', RoleTypeKind: 5 }, { Id: 1073741830, Name: 'Edit', RoleTypeKind: 6 }, { Id: 1073741827, Name: 'Contribute', RoleTypeKind: 3 }, { Id: 1073741826, Name: 'Read', RoleTypeKind: 2 }] });
    }
    if (/web\/lists\?/.test(path)) {
      return response(200, { value: Object.keys(lists).map((t) => ({ Title: t, LastItemModifiedDate: lists[t].modified, LastItemDeletedDate: lists[t].modified })) });
    }
    // Files by server-relative address (attachments read back, e.g. when an approved suggestion's files are copied).
    const fileAt = /GetFileByServerRelativePath\(decodedurl=@a\)\/\$value\?@a=([^&]+)/.exec(path);
    if (fileAt) {
      const u = decodeURIComponent(fileAt[1]).replace(/^'|'$/g, '').replace(/''/g, "'");
      const b64 = (window.__edPreviewFiles || {})['att:' + u];
      return b64 ? response(200, '', fromB64(b64)) : response(404, { 'odata.error': { message: { value: 'File not found' } } });
    }
    const title = titleOf(path);
    if (!title) return response(404, { 'odata.error': { message: { value: 'Not found in preview: ' + path } } });
    const l = list(title);
    const id = itemIdOf(path);
    // Libraries and list folders: a root folder address, and uploaded files kept in memory (window.__edPreviewFiles,
    // name → base64) so a browser test can serve them back.
    if (/\/RootFolder\?\$select=ServerRelativeUrl/.test(path)) return response(200, { ServerRelativeUrl: '/sites/timesheet/' + title });
    const add = /\/RootFolder\/Files\/add\(url='([^']+)'/.exec(decodeURIComponent(path));
    if (add && method === 'POST') {
      const bytes = new Uint8Array(body instanceof ArrayBuffer ? body : body && body.buffer ? body.buffer : new ArrayBuffer(0));
      let bin = '';
      for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
      window.__edPreviewFiles = window.__edPreviewFiles || {};
      window.__edPreviewFiles[title + '/' + add[1]] = btoa(bin);
      return response(200, { Name: add[1], ServerRelativeUrl: '/sites/timesheet/' + title + '/' + add[1] });
    }
    // List-item attachments (window.__edPreviewFiles['att:' + server-relative URL] holds the bytes).
    if (/\/AttachmentFiles/.test(path) && id !== undefined) {
      const row = l.items.filter((x) => x.Id === id)[0];
      if (!row) return response(404, { 'odata.error': { message: { value: 'Item not found' } } });
      row.__files = row.__files || [];
      window.__edPreviewFiles = window.__edPreviewFiles || {};
      const base = `/sites/timesheet/Lists/${title}/Attachments/${id}/`;
      const named = /(?:add\(FileName|getByFileName\()=?'((?:[^']|'')*)'/.exec(decodeURIComponent(path));
      const name = named ? named[1].replace(/''/g, "'") : '';
      if (/AttachmentFiles\/add/.test(path)) {
        if (row.__files.indexOf(name) >= 0) return response(400, { 'odata.error': { message: { value: 'A file with this name already exists' } } });
        row.__files.push(name);
        window.__edPreviewFiles['att:' + base + name] = toB64(body);
        return response(200, { FileName: name, ServerRelativeUrl: base + name });
      }
      if (name && (verb === 'DELETE' || method === 'DELETE')) {
        row.__files = row.__files.filter((n) => n !== name);
        delete window.__edPreviewFiles['att:' + base + name];
        return response(200);
      }
      return response(200, { value: row.__files.map((n) => ({ FileName: n, ServerRelativeUrl: base + n })) });
    }
    // Item and list permissions: remembered on the item (__readers, __unique) so a browser test can check them.
    if (/breakroleinheritance|resetroleinheritance|roleassignments/.test(path)) {
      const row = id !== undefined ? l.items.filter((x) => x.Id === id)[0] : undefined;
      if (row && /breakroleinheritance/.test(path)) Object.assign(row, { HasUniqueRoleAssignments: true, __readers: [] });
      if (row && /resetroleinheritance/.test(path)) Object.assign(row, { HasUniqueRoleAssignments: false, __readers: undefined });
      const add = /addroleassignment\(principalid=(\d+)/.exec(path);
      if (row && add) row.__readers.push(Number(add[1]));
      return response(200, method === 'GET' ? { value: [] } : {});
    }
    if (/\/items/.test(path)) {
      if (method === 'GET' && id !== undefined) {
        const row = l.items.filter((x) => x.Id === id)[0];
        return row ? response(200, row) : response(404, { 'odata.error': { message: { value: 'Item not found' } } });
      }
      if (method === 'GET') {
        const q = query(path);
        let rows = l.items.filter(parseFilter(q.$filter));
        if (q.$orderby) {
          const [field, dir] = decodeURIComponent(q.$orderby).split(',')[0].split(' ');
          rows = rows.slice().sort((a, b) => (String(a[field]) < String(b[field]) ? -1 : String(a[field]) > String(b[field]) ? 1 : 0) * (dir === 'desc' ? -1 : 1));
        }
        if (q.$top) rows = rows.slice(0, Number(q.$top));
        return response(200, { value: rows });
      }
      const data = stored(body ? JSON.parse(body) : {});
      l.modified = nowIso();
      if (verb === 'DELETE') {
        l.items = l.items.filter((x) => x.Id !== id);
        return response(200);
      }
      if (verb === 'MERGE') {
        const row = l.items.filter((x) => x.Id === id)[0];
        if (!row) return response(404, { 'odata.error': { message: { value: 'Item not found' } } });
        Object.assign(row, data, { Modified: nowIso() });
        if (data.ReviewerId) row.Reviewer = { Id: data.ReviewerId, Title: user.Title, EMail: user.Email };
        return response(204);
      }
      const row = { Id: nextId++, Created: nowIso(), Modified: nowIso(), Author: { Id: user.Id, Title: user.Title, EMail: user.Email }, AuthorId: user.Id, ...data };
      l.items.push(row);
      return response(201, row);
    }
    return response(200, { Id: '0f8fad5b-d9cb-469f-a165-70867728950e', EnableAttachments: false, EnableVersioning: false, ItemCount: l.items.length });
  };

  const batch = (body) => {
    const out = [];
    const parts = String(body).split(/\r\n--changeset_[^\r\n]*--/);
    parts.forEach((p) => {
      const m = /POST (\S+) HTTP\/1\.1\r\n([\s\S]*?)\r\n\r\n([\s\S]*)$/.exec(p);
      if (!m) return;
      const headers = {};
      m[2].split('\r\n').forEach((h) => {
        const i = h.indexOf(':');
        if (i > 0) headers[h.substring(0, i).trim()] = h.substring(i + 1).trim();
      });
      const json = m[3].trim();
      const res = handle('POST', m[1], json || undefined, headers);
      out.push(res);
    });
    return Promise.all(out.map((r) => r.text().then((t) => ({ status: r.status, t })))).then((rs) => {
      const lines = [];
      rs.forEach((r) => lines.push('--batchresponse_x', 'Content-Type: application/http', 'Content-Transfer-Encoding: binary', '', `HTTP/1.1 ${r.status} X`, 'CONTENT-TYPE: application/json', '', r.t));
      lines.push('--batchresponse_x--', '');
      return response(200, lines.join('\r\n'));
    });
  };

  const run = (method) => (url, _config, options) => {
    const o = options || {};
    return new Promise((resolve) => setTimeout(() => resolve(handle(method, url, o.body, o.headers || {})), 15));
  };
  return { get: run('GET'), post: run('POST') };
}

/** ?flow=1: stands in for the Power Automate flow — sends queued emails a few seconds later ("fail" in the address fails). */
export function simulateFlow() {
  setInterval(() => {
    const q = lists.BrillTS_EmailQueue;
    if (!q) return;
    q.items.forEach((row) => {
      if (row.MailStatus !== 'Queued' || Date.now() - Date.parse(row.Created) < 3000) return;
      const fail = /fail/i.test(row.MailTo || '');
      Object.assign(row, { MailStatus: fail ? 'Failed' : 'Sent', MailError: fail ? 'Mailbox unavailable' : '', Modified: nowIso() });
      q.modified = nowIso();
    });
  }, 1000);
}
