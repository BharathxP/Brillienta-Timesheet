// Where the timesheet app runs (a SharePoint page or Microsoft Teams) and which site holds its lists.
//
// On a SharePoint page the lists live on the page's own site. In Teams the host site is not the timesheet app's site:
// a channel tab runs on the team's site and a personal app on the tenant root site. There the timesheet app site comes
// from the tab's "Timesheet site" setting, else the tenant property below, else the host site itself.

/** Tenant property (app catalog storage entity) holding the timesheet app site URL for Teams. */
export const TENANT_SITE_KEY = 'BrillientaTimesheetSiteUrl';

/** Pages that host a web part in Teams: links to them only open inside Teams, so they are never shared. */
const TEAMS_HOST_PAGE = /\/_layouts\/15\/(teamshostedapp|teamslogon)\.aspx/i;
/** The local workbench: a page nobody else can open. */
const WORKBENCH_PAGE = /\/_layouts\/15\/workbench\.aspx/i;
/** Parts of a URL below the site itself (pages, lists, libraries, layouts). */
const BELOW_SITE = /\/(SitePages|Pages|Lists|Shared%20Documents|Shared Documents|SiteAssets|_layouts)(\/.*)?$/i;

/**
 * The site URL an admin typed (absolute or server-relative, possibly a page or list inside it), as an absolute URL
 * without a trailing slash, query or hash. '' when it is not a SharePoint address on `origin`: the web part can only
 * call its own tenant.
 */
export function normalizeSiteUrl(input: string, origin: string): string {
  const raw = (input || '').trim();
  if (!raw) return '';
  let url: URL;
  try {
    url = new URL(raw, origin);
  } catch {
    return '';
  }
  if (url.protocol !== 'https:' && url.protocol !== 'http:') return '';
  if (url.origin.toLowerCase() !== origin.replace(/\/$/, '').toLowerCase()) return '';
  const path = url.pathname.replace(BELOW_SITE, '').replace(/\/+$/, '');
  return url.origin + path;
}

/** Same site? (case and trailing slash do not matter) */
export function sameSite(a: string, b: string): boolean {
  const n = (u: string): string => (u || '').replace(/\/+$/, '').toLowerCase();
  return !!a && n(a) === n(b);
}

/** True for the page Teams loads a web part into. */
export function isTeamsHostPage(url: string): boolean {
  return TEAMS_HOST_PAGE.test(url || '');
}

/** A page address worth recording as "the timesheet app page": a real SharePoint page, not Teams' host page or the workbench. */
export function isShareablePage(url: string): boolean {
  const u = (url || '').split('#')[0];
  return /^https?:\/\//i.test(u) && !isTeamsHostPage(u) && !WORKBENCH_PAGE.test(u);
}

/**
 * Base for links that leave this browser (email links, QR codes, "copy link", vCards): the current page when it is a
 * real SharePoint page, otherwise the recorded timesheet page, otherwise the current page anyway.
 */
export function shareableBase(currentHref: string, storedPageUrl: string): string {
  const current = (currentHref || '').split('#')[0];
  if (isShareablePage(current)) return current;
  const stored = (storedPageUrl || '').split('#')[0];
  return isShareablePage(stored) ? stored : current;
}

/** Teams deep links (chat, call, channel, meeting): inside Teams they must open through the Teams SDK. */
export function isTeamsDeepLink(href: string): boolean {
  return /^https:\/\/teams\.(microsoft|cloud\.microsoft)(\.com)?\/l\//i.test(href || '');
}

// ---- Finding the timesheet from Teams without any setup ----
//
// In Teams the web part runs on the team's site or the tenant root site. When neither the tab's setting, that site, nor
// the tenant property points at a timesheet, it looks for one with SharePoint search: the Settings list is kept in
// search, and search only returns sites the person can read. No extra permission is needed.

/** Setting (in a timesheet's own Settings list) marking it as the one Teams opens when search finds several. */
export const TEAMS_DEFAULT_SETTING = 'TeamsDefault';
/** This browser's choice among several timesheets (localStorage). */
export const TEAMS_CHOICE_KEY = 'brillts.teamsTimesheet';

/** SharePoint search query (KQL) for Settings lists titled `settingsTitle` (e.g. "BrillTS_Settings"). */
export function settingsSearchQuery(settingsTitle: string): string {
  return `${settingsTitle.replace(/[^A-Za-z0-9_]/g, '')} contentclass:STS_List*`;
}

/** One result row of SharePoint's search REST API (nometadata), as cells. */
export interface ISearchRow {
  Cells: Array<{ Key: string; Value: string | null }>;
}

/**
 * Site URLs of the timesheets in search results: rows that are the `settingsTitle` list itself (its title, or a path
 * inside /Lists/<title>), on this tenant, each site once, in result order.
 */
export function sitesFromSearch(rows: ISearchRow[], settingsTitle: string, origin: string): string[] {
  const title = settingsTitle.toLowerCase();
  const marker = `/lists/${title}`;
  const out: string[] = [];
  (rows || []).forEach((row) => {
    const cell: Record<string, string> = {};
    (row.Cells || []).forEach((c) => (cell[c.Key] = c.Value || ''));
    let path = cell.Path || '';
    try {
      path = decodeURIComponent(path);
    } catch {
      // keep it encoded
    }
    const at = path.toLowerCase().indexOf(marker);
    const next = at >= 0 ? path.charAt(at + marker.length) : 'x';
    const inList = at >= 0 && (next === '' || next === '/' || next === '?');
    if ((cell.Title || '').toLowerCase() !== title && !inList) return;
    const site = normalizeSiteUrl(cell.SPWebUrl || (inList ? path.slice(0, at) : ''), origin);
    if (site && !out.some((s) => sameSite(s, site))) out.push(site);
  });
  return out;
}

/** Which found timesheet Teams opens: the only one, else this browser's choice, else the only marked one; else the person chooses. */
export function pickSite(found: Array<{ url: string; marked: boolean }>, remembered: string): { url: string } | { choose: string[] } | undefined {
  if (!found.length) return undefined;
  if (found.length === 1) return { url: found[0].url };
  const mine = found.filter((f) => sameSite(f.url, remembered))[0];
  if (mine) return { url: mine.url };
  const marked = found.filter((f) => f.marked);
  if (marked.length === 1) return { url: marked[0].url };
  return { choose: (marked.length ? marked : found).map((f) => f.url) };
}
