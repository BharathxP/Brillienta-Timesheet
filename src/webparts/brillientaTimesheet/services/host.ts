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
