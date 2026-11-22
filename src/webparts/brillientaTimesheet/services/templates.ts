// Template rendering and the branded email layout. Pure functions (unit tested).

export function escapeHtml(v: string): string {
  return v.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/**
 * Replace {{name}} (HTML-escaped when `html`) and {{{name}}} (inserted as is — only for HTML the app builds).
 * Unknown placeholders become empty.
 */
export function renderTemplate(text: string, vars: Record<string, string | number | undefined>, html: boolean = true): string {
  return (text || '')
    .replace(/\{\{\{\s*(\w+)\s*\}\}\}/g, (_, k: string) => String(vars[k] === undefined ? '' : vars[k]))
    .replace(/\{\{\s*(\w+)\s*\}\}/g, (_, k: string) => {
      const v = String(vars[k] === undefined ? '' : vars[k]);
      return html ? escapeHtml(v) : v;
    });
}

/** Plain text of an HTML fragment (for in-app notifications). */
export function toPlainText(html: string, max: number = 400): string {
  const text = html
    .replace(/<\s*(br|\/p|\/li|\/blockquote|\/h\d)\s*\/?>/gi, '\n')
    .replace(/<li[^>]*>/gi, '• ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n+/g, '\n')
    .trim();
  return text.length > max ? text.substring(0, max - 1) + '…' : text;
}

function safeColor(c: string, fallback: string): string {
  return /^#[0-9a-f]{3,8}$/i.test(c || '') ? c : fallback;
}

/** Table-based email (works in Outlook) in the Brillienta navy-and-lime style. */
/**
 * `headerHtml` and `footerHtml` are inserted as they are: pass them through sanitizeHtml (services/html.ts) and
 * renderTemplate first, as Notifier.compose does.
 */
export function emailLayout(o: { orgName: string; brandColor: string; title: string; bodyHtml: string; cta?: string; link?: string; headerHtml?: string; footerHtml?: string }): string {
  const navy = safeColor(o.brandColor, '#17232e');
  const font = "'Segoe UI', -apple-system, BlinkMacSystemFont, Arial, sans-serif";
  const button =
    o.cta && o.link
      ? `<tr><td style="padding:8px 32px 28px"><a href="${escapeHtml(o.link)}" style="display:inline-block;background:${navy};color:#ffffff;text-decoration:none;font-weight:600;font-size:14px;padding:12px 22px;border-radius:7px;font-family:${font}">${escapeHtml(o.cta)} &nbsp;↗</a></td></tr>`
      : '';
  return (
    `<!doctype html><html><body style="margin:0;padding:0;background:#f4f7f6">` +
    `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f7f6;padding:24px 12px"><tr><td align="center">` +
    `<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border:1px solid #dfe5e6;border-radius:12px;overflow:hidden">` +
    `<tr><td style="background:${navy};padding:18px 32px;font-family:${font}">` +
    `<span style="display:inline-block;width:26px;height:28px;line-height:28px;text-align:center;background:#caf578;color:${navy};font-weight:800;font-size:18px;border-radius:6px 6px 6px 2px">${escapeHtml((o.orgName || 'e').charAt(0).toLowerCase())}</span>` +
    `<span style="color:#ffffff;font-weight:700;font-size:16px;margin-left:10px;vertical-align:middle">${o.headerHtml === undefined ? escapeHtml(o.orgName) + ' timesheet' : o.headerHtml}</span></td></tr>` +
    `<tr><td style="padding:28px 32px 4px;font-family:${font}"><h1 style="margin:0 0 14px;font-size:22px;line-height:1.3;color:#17232e">${escapeHtml(o.title)}</h1>` +
    `<div style="font-size:15px;line-height:1.65;color:#3d4a53">${o.bodyHtml}</div></td></tr>` +
    button +
    `<tr><td style="padding:16px 32px 22px;border-top:1px solid #eef1f1;font-family:${font};font-size:12px;color:#8a959c">${o.footerHtml || ''}</td></tr>` +
    `</table></td></tr></table></body></html>`
  ).replace(/<blockquote>/g, '<blockquote style="margin:12px 0;padding:10px 16px;border-left:3px solid #afce8b;background:#f5f7f4;color:#3d4a53">');
}
