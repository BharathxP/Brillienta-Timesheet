import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPPermission } from '@microsoft/sp-page-context';
import { SPHttpClient } from '@microsoft/sp-http';
import { initializeIcons } from '@fluentui/react';

import * as strings from 'BrillientaTimesheetWebPartStrings';
import App, { IAppProps } from './components/App';
import { LIST_PREFIX, listTitle } from './schema/schema';
import { ISearchRow, TEAMS_CHOICE_KEY, TEAMS_DEFAULT_SETTING, TENANT_SITE_KEY, normalizeSiteUrl, pickSite, sameSite, settingsSearchQuery, sitesFromSearch } from './services/host';

export interface IBrillientaTimesheetWebPartProps {
  /** Site that holds the timesheet lists when it is not the site the web part runs on (Teams tabs). */
  siteUrl?: string;
}

interface ITimesheetWeb {
  url: string;
  title: string;
  manageLists: boolean;
  manageWeb: boolean;
  source: '' | 'property' | 'tenant' | 'found';
  /** Teams found several timesheets: the person picks one. */
  choices?: Array<{ url: string; title: string }>;
  error?: string;
}

const NOMETA = { headers: { Accept: 'application/json;odata=nometadata' } };
let iconsReady = false;

export default class BrillientaTimesheetWebPart extends BaseClientSideWebPart<IBrillientaTimesheetWebPartProps> {
  private _web: ITimesheetWeb | undefined;
  private _inTeams = false;

  protected async onInit(): Promise<void> {
    await super.onInit();
    if (!iconsReady) {
      initializeIcons(undefined, { disableWarnings: true });
      iconsReady = true;
    }
    this._inTeams = !!(this.context.sdks && this.context.sdks.microsoftTeams);
    this._web = await this._resolveWeb();
  }

  public render(): void {
    const ctx = this.context;
    const perms = ctx.pageContext.web.permissions;
    const aad = ctx.pageContext.aadInfo;
    const host = ctx.pageContext.web.absoluteUrl;
    const web: ITimesheetWeb = this._web || { url: host, title: ctx.pageContext.web.title, manageLists: false, manageWeb: false, source: '' };
    const onHost = sameSite(web.url, host);
    const teams = ctx.sdks && ctx.sdks.microsoftTeams;
    const props: IAppProps = {
      http: ctx.spHttpClient,
      webUrl: web.url,
      siteTitle: web.title,
      user: {
        displayName: ctx.pageContext.user.displayName,
        email: ctx.pageContext.user.email,
        loginName: ctx.pageContext.user.loginName,
        isExternal: !!ctx.pageContext.user.isExternalGuestUser
      },
      permissions: onHost
        ? {
            manageLists: perms.hasPermission(SPPermission.manageLists),
            manageWeb: perms.hasPermission(SPPermission.manageWeb) || perms.hasPermission(SPPermission.managePermissions)
          }
        : { manageLists: web.manageLists, manageWeb: web.manageWeb },
      tenantId: aad && aad.tenantId ? aad.tenantId.toString() : undefined,
      prefix: this._prefix(),
      appVersion: this.manifest.version,
      isEditMode: this.displayMode === DisplayMode.Edit,
      graphFactory: ctx.msGraphClientFactory,
      host: {
        inTeams: this._inTeams,
        siteSource: web.source,
        hostWebUrl: host,
        siteError: web.error,
        siteChoices: web.choices,
        openTeamsLink: teams ? (url: string) => teams.teamsJs.app.openLink(url) : undefined
      }
    };
    ReactDom.render(React.createElement(App, props), this.domElement);
  }

  /** Every list is named LIST_PREFIX + its name (BrillTS_Settings, BrillTS_TimeEntries…). */
  private _prefix(): string {
    return LIST_PREFIX;
  }

  private _origin(): string {
    return new URL(this.context.pageContext.web.absoluteUrl).origin;
  }

  private async _getJson<T>(url: string): Promise<T> {
    const res = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, NOMETA);
    if (!res.ok) {
      const e = new Error(`${res.status} ${res.statusText}`) as Error & { status?: number };
      e.status = res.status;
      throw e;
    }
    return (await res.json()) as T;
  }

  private async _hasTimesheet(webUrl: string): Promise<boolean> {
    const title = listTitle(this._prefix(), 'settings').replace(/'/g, "''");
    try {
      await this._getJson(`${webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(title)}')?$select=Id`);
      return true;
    } catch {
      return false;
    }
  }

  private async _tenantSite(): Promise<string> {
    try {
      const r = await this._getJson<{ Value?: string }>(`${this.context.pageContext.web.absoluteUrl}/_api/web/GetStorageEntity('${TENANT_SITE_KEY}')`);
      return normalizeSiteUrl((r && r.Value) || '', this._origin());
    } catch {
      return '';
    }
  }

  /** This browser's pick among several timesheets (Teams). */
  private _remembered(): string {
    try {
      return window.localStorage.getItem(TEAMS_CHOICE_KEY) || '';
    } catch {
      return '';
    }
  }

  /**
   * Teams with nothing set up: timesheets the person can open, found with SharePoint search (the Settings list is kept
   * in search). One → that one; several → this browser's pick, else the one marked "Use this site in Teams", else the
   * person chooses. Search can take a few hours to list a new timesheet.
   */
  private async _findTimesheet(): Promise<{ url: string } | { choose: Array<{ url: string; title: string }> } | undefined> {
    const host = this.context.pageContext.web.absoluteUrl;
    const origin = this._origin();
    const settings = listTitle(this._prefix(), 'settings');
    const remembered = normalizeSiteUrl(this._remembered(), origin);
    if (remembered && (await this._hasTimesheet(remembered))) return { url: remembered };
    let sites: string[] = [];
    try {
      const q = encodeURIComponent(`'${settingsSearchQuery(settings)}'`);
      const r = await this._getJson<{ PrimaryQueryResult?: { RelevantResults?: { Table?: { Rows?: ISearchRow[] } } } }>(
        `${host}/_api/search/query?querytext=${q}&selectproperties='Title,Path,SPWebUrl'&rowlimit=50&trimduplicates=false`
      );
      const rows = (r.PrimaryQueryResult && r.PrimaryQueryResult.RelevantResults && r.PrimaryQueryResult.RelevantResults.Table && r.PrimaryQueryResult.RelevantResults.Table.Rows) || [];
      sites = sitesFromSearch(rows, settings, origin).slice(0, 10);
    } catch {
      return undefined;
    }
    const title = settings.replace(/'/g, "''");
    const found = (
      await Promise.all(
        sites.map(async (url) => {
          try {
            const r = await this._getJson<{ value: Array<{ SettingValue: string | null }> }>(
              `${url}/_api/web/lists/getByTitle('${encodeURIComponent(title)}')/items?$select=SettingValue&$filter=Title eq '${TEAMS_DEFAULT_SETTING}'&$top=1`
            );
            return { url, marked: !!r.value && !!r.value[0] && r.value[0].SettingValue === '1' };
          } catch {
            return undefined; // not readable, or not a timesheet after all
          }
        })
      )
    ).filter((f): f is { url: string; marked: boolean } => !!f);
    const pick = pickSite(found, remembered);
    if (!pick || 'url' in pick) return pick;
    const choose = await Promise.all(
      pick.choose.map(async (url) => {
        try {
          return { url, title: (await this._getJson<{ Title: string }>(`${url}/_api/web?$select=Title`)).Title || url };
        } catch {
          return { url, title: url };
        }
      })
    );
    return { choose };
  }

  /**
   * The timesheet site: the web part setting, else (in Teams, when the host site has none) the tenant-wide timesheet
   * site, else one found with SharePoint search; otherwise the host site.
   */
  private async _resolveWeb(): Promise<ITimesheetWeb> {
    const pc = this.context.pageContext.web;
    const host: ITimesheetWeb = { url: pc.absoluteUrl, title: pc.title, manageLists: false, manageWeb: false, source: '' };
    const typed = (this.properties.siteUrl || '').trim();
    let url = '';
    let source: ITimesheetWeb['source'] = '';
    if (typed) {
      url = normalizeSiteUrl(typed, this._origin());
      if (!url) return { ...host, error: strings.SiteUrlInvalid.replace('{0}', typed) };
      source = 'property';
    } else if (this._inTeams && !(await this._hasTimesheet(pc.absoluteUrl))) {
      url = await this._tenantSite();
      source = 'tenant';
      if (!url) {
        const found = await this._findTimesheet();
        if (found && 'choose' in found) return { ...host, choices: found.choose };
        url = found ? found.url : '';
        source = 'found';
      }
    }
    if (!url || sameSite(url, pc.absoluteUrl)) return host;
    try {
      const w = await this._getJson<{ Title: string; EffectiveBasePermissions: { High: string; Low: string } }>(`${url}/_api/web?$select=Title,EffectiveBasePermissions`);
      const p = new SPPermission({ High: Number(w.EffectiveBasePermissions.High), Low: Number(w.EffectiveBasePermissions.Low) });
      return {
        url,
        title: w.Title,
        manageLists: p.hasPermission(SPPermission.manageLists),
        manageWeb: p.hasPermission(SPPermission.manageWeb) || p.hasPermission(SPPermission.managePermissions),
        source
      };
    } catch (e) {
      const status = (e as { status?: number }).status;
      const why = status === 403 || status === 401 ? strings.SiteNoAccess : (e as Error).message;
      return { ...host, url, source, error: strings.SiteUnreachable.replace('{0}', url).replace('{1}', why) };
    }
  }

  protected onDisplayModeChanged(): void {
    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    this._resolveWeb()
      .then((w) => {
        this._web = w;
        this.render();
      })
      .catch(() => this.render());
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.SiteGroupName,
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteUrlLabel,
                  description: strings.SiteUrlDescription,
                  placeholder: 'https://contoso.sharepoint.com/sites/timesheet',
                  onGetErrorMessage: (v: string) => (!v || !v.trim() || normalizeSiteUrl(v, this._origin()) ? '' : strings.SiteUrlInvalid.replace('{0}', v)),
                  deferredValidationTime: 400
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
