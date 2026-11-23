// Allow-list HTML cleaner for admin-written email content (template messages, header, footer). Pure (unit tested).
//
// The rich text editor loads stored HTML into the page, and emails carry it to every recipient, so anything that can
// run code or load from elsewhere is removed: scripts, styles, frames, forms, event handlers, javascript: links and CSS
// with url()/expression(). Formatting that email clients understand is kept. {{placeholders}} pass through untouched.

const ALLOWED_TAGS: Record<string, string[]> = {
  p: [],
  br: [],
  div: [],
  span: [],
  strong: [],
  b: [],
  em: [],
  i: [],
  u: [],
  s: [],
  strike: [],
  del: [],
  sub: [],
  sup: [],
  small: [],
  code: [],
  pre: [],
  hr: [],
  h1: [],
  h2: [],
  h3: [],
  h4: [],
  blockquote: [],
  ul: [],
  ol: [],
  li: [],
  a: ['href', 'title', 'target'],
  img: ['src', 'alt', 'width', 'height'],
  table: ['width', 'cellpadding', 'cellspacing', 'border'],
  thead: [],
  tbody: [],
  tr: [],
  td: ['colspan', 'rowspan', 'align', 'valign', 'width'],
  th: ['colspan', 'rowspan', 'align', 'valign', 'width']
};
const COMMON_ATTRS = ['style', 'align'];
const VOID = ['br', 'hr', 'img'];
/** Removed together with everything inside them. */
const DROP_WITH_CONTENT = ['script', 'style', 'iframe', 'object', 'embed', 'noscript', 'template', 'svg', 'math', 'head', 'title', 'textarea', 'select', 'button', 'form', 'frameset', 'frame', 'applet', 'xml'];
const CSS_PROPS = /^(color|background|background-color|font|font-weight|font-style|font-size|font-family|text-decoration|text-align|text-transform|line-height|letter-spacing|vertical-align|white-space|width|max-width|min-width|height|display|border|border-(top|right|bottom|left)(-(width|style|color))?|border-(width|style|color|radius|collapse)|margin|margin-(top|right|bottom|left)|padding|padding-(top|right|bottom|left)|list-style(-type)?)$/i;

function decodeEntities(v: string): string {
  return v
    .replace(/&#x([0-9a-f]+);?/gi, (_, h: string) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&#(\d+);?/g, (_, d: string) => String.fromCharCode(Number(d)))
    .replace(/&colon;/gi, ':')
    .replace(/&tab;/gi, '\t')
    .replace(/&newline;/gi, '\n')
    .replace(/&quot;/gi, '"')
    .replace(/&apos;/gi, "'")
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&amp;/gi, '&');
}

function attrEscape(v: string): string {
  return v.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** A link or image address that is safe in an email: http(s), mailto, tel, an anchor, or a {{placeholder}}. */
export function safeUrl(raw: string, images: boolean = false): string {
  const v = decodeEntities(raw).trim();
  // Placeholders: only the app's own links (values people write, such as a note, never become addresses).
  if (/\{\{/.test(v)) return /^\{\{\s*(link|directoryLink)\s*\}\}$/.test(v) ? v : '';
  // Browsers ignore control characters and whitespace inside a scheme ("java\tscript:"), so compare without them.
  // eslint-disable-next-line no-control-regex
  const bare = v.replace(/[\u0000- \u007f-\u009f]+/g, '').toLowerCase();
  if (images) return /^https:\/\//.test(bare) ? v : '';
  if (/^(https?:|mailto:|tel:)/.test(bare) || /^#/.test(bare)) return v;
  return /^[a-z][a-z0-9+.-]*:/.test(bare) ? '' : v.indexOf('//') === 0 ? '' : v;
}

/** Inline CSS limited to formatting properties, without url(), expression() or imports. */
export function safeStyle(raw: string): string {
  const v = decodeEntities(raw);
  if (/url\s*\(|expression\s*\(|javascript:|behavior\s*:|-moz-binding|@import|[<>\\]/i.test(v)) return '';
  return v
    .split(';')
    .map((d) => d.trim())
    .filter((d) => {
      const i = d.indexOf(':');
      if (i <= 0) return false;
      const prop = d.substring(0, i).trim();
      const value = d.substring(i + 1).trim();
      return CSS_PROPS.test(prop) && /^[#\w\s.,%()'"!/+-]*$/.test(value) && !!value;
    })
    .join('; ');
}

function cleanAttrs(tag: string, raw: string): string {
  const allowed = (ALLOWED_TAGS[tag] || []).concat(COMMON_ATTRS);
  const out: string[] = [];
  const re = /([^\s=/"'>]+)(?:\s*=\s*("[^"]*"|'[^']*'|[^\s"'=<>`]+))?/g;
  let m: RegExpExecArray | null;
  const seen: Record<string, boolean> = {};
  while ((m = re.exec(raw))) {
    const name = m[1].toLowerCase();
    if (allowed.indexOf(name) < 0 || seen[name]) continue;
    let value = m[2] === undefined ? '' : m[2].replace(/^["']|["']$/g, '');
    if (name === 'href') value = safeUrl(value);
    else if (name === 'src') value = safeUrl(value, true);
    else if (name === 'style') value = safeStyle(value);
    else if (name === 'target') value = value === '_blank' ? '_blank' : '';
    else value = decodeEntities(value).replace(/[^\w\s.,%#:-]/g, '');
    if (!value) continue;
    seen[name] = true;
    out.push(`${name}="${attrEscape(value)}"`);
  }
  if (tag === 'a' && seen.target) out.push('rel="noopener noreferrer"');
  return out.length ? ' ' + out.join(' ') : '';
}

/** Clean HTML: allowed tags and attributes only, every open tag closed. */
export function sanitizeHtml(html: string): string {
  const src = String(html || '');
  const out: string[] = [];
  const open: string[] = [];
  let skip: string | undefined;
  let skipDepth = 0;
  const re = /<!--[\s\S]*?(-->|$)|<!\[CDATA\[[\s\S]*?(\]\]>|$)|<![^>]*>|<\?[^>]*>|<(\/?)([a-zA-Z][a-zA-Z0-9-]*)((?:[^>"']|"[^"]*"|'[^']*')*)>|<|[^<]+/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(src))) {
    const token = m[0];
    const tag = m[4] ? m[4].toLowerCase() : '';
    if (skip) {
      if (tag === skip) skipDepth += m[3] ? -1 : 1;
      if (skipDepth === 0) skip = undefined;
      continue;
    }
    if (token[0] !== '<') {
      out.push(token.replace(/>/g, '&gt;'));
      continue;
    }
    if (!tag) {
      if (token === '<') out.push('&lt;');
      continue; // comments, doctype, processing instructions
    }
    if (DROP_WITH_CONTENT.indexOf(tag) >= 0) {
      if (!m[3] && !/\/\s*$/.test(m[5] || '')) {
        skip = tag;
        skipDepth = 1;
      }
      continue;
    }
    if (!ALLOWED_TAGS[tag]) continue; // unknown tag: dropped, its text kept
    if (m[3]) {
      const at = open.lastIndexOf(tag);
      if (at < 0) continue;
      while (open.length > at) out.push(`</${open.pop()}>`);
      continue;
    }
    out.push(`<${tag}${cleanAttrs(tag, m[5] || '')}>`);
    if (VOID.indexOf(tag) < 0) open.push(tag);
  }
  while (open.length) out.push(`</${open.pop()}>`);
  return out.join('');
}

/**
 * The header sits on one line in the coloured band: paragraphs, headings and blocks become line breaks there, so
 * the band does not stretch with their margins.
 */
export function headerInline(html: string): string {
  return html
    .replace(/<(p|div|h[1-4]|blockquote|ul|ol|li|pre)(\s[^>]*)?>/gi, '')
    .replace(/<\/(p|div|h[1-4]|blockquote|ul|ol|li|pre)>/gi, '<br>')
    .replace(/(<br>\s*)+$/i, '')
    .replace(/^(\s*<br>)+/i, '');
}
