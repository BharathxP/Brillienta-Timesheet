// Sends a notification on every channel that is switched on: the in-app bell (always), email (Outlook through
// Microsoft Graph, or the email queue that a Power Automate flow sends), a Teams chat message, and the Power Automate
// webhook. A channel that fails never stops the others or the action that caused it.

import { ITimesheetConfig, TemplateKey } from '../models/types';
import { TimesheetService } from './TimesheetService';
import { GraphService } from './GraphService';
import { emailLayout, renderTemplate, toPlainText } from './templates';
import { sanitizeHtml } from './html';

export interface IRecipient {
  email: string;
  name: string;
}

export interface ISendOptions {
  template: TemplateKey;
  to: IRecipient[];
  cc?: string[];
  vars: Record<string, string | number | undefined>;
  /** App route the link opens, e.g. "approvals" (made absolute with the page address). */
  path: string;
  cta?: string;
}

export const WEBHOOK_EVENTS: Array<{ key: string; text: string }> = [
  { key: 'timesheet.submitted', text: 'Timesheet submitted' },
  { key: 'timesheet.approved', text: 'Timesheet approved' },
  { key: 'timesheet.rejected', text: 'Timesheet rejected' },
  { key: 'project.created', text: 'Project created' },
  { key: 'project.ending', text: 'Project end date approaching' },
  { key: 'task.assigned', text: 'Task assigned' },
  { key: 'reminder.sent', text: 'Reminder sent' }
];

export class Notifier {
  private readonly _svc: TimesheetService;
  private readonly _config: () => ITimesheetConfig;
  private readonly _graph?: GraphService;
  /** Page address the links in emails point to (no hash). */
  public linkBase: string;
  /** Mail and Teams are on when the plan includes Teams & Outlook connectivity. */
  public teamsOutlook = true;
  public onProblem: (message: string) => void = () => undefined;

  public constructor(svc: TimesheetService, config: () => ITimesheetConfig, linkBase: string, graph?: GraphService) {
    this._svc = svc;
    this._config = config;
    this._graph = graph;
    this.linkBase = linkBase;
  }

  public link(path: string): string {
    return `${this.linkBase}#/${path.replace(/^#?\/?/, '')}`;
  }

  /** Subject and branded HTML for one recipient. */
  public compose(template: TemplateKey, vars: Record<string, string | number | undefined>, path: string, cta?: string, config?: ITimesheetConfig): { subject: string; html: string; text: string; teams: boolean; enabled: boolean } {
    const cfg = config || this._config();
    const t = cfg.templates[template];
    const all = { org: cfg.orgName, app: cfg.appName, link: this.link(path), ...vars };
    const subject = renderTemplate(t.subject, all, false);
    const body = sanitizeHtml(renderTemplate(t.body, all, true));
    const html = emailLayout({ orgName: cfg.orgName || cfg.appName, brandColor: cfg.brandColor, title: subject, bodyHtml: body, cta, link: all.link, footerHtml: `Sent by ${cfg.appName}${cfg.orgName ? ' · ' + cfg.orgName : ''}` });
    return { subject, html, text: toPlainText(body), teams: t.teams, enabled: t.enabled };
  }

  public async send(o: ISendOptions): Promise<void> {
    const cfg = this._config();
    const to = o.to.filter((r, i, a) => !!r.email && a.map((x) => x.email.toLowerCase()).indexOf(r.email.toLowerCase()) === i);
    if (!to.length) return;
    const notices: Array<{ recipient: string; title: string; body: string; link: string; kind: string }> = [];
    for (const r of to) {
      const m = this.compose(o.template, { recipient: r.name.split(' ')[0] || r.name, ...o.vars }, o.path, o.cta);
      notices.push({ recipient: r.email, title: m.subject, body: m.text, link: o.path, kind: o.template });
      if (!m.enabled || o.template === 'ticketRaised') continue;
      await this._email([r.email], o.cc || [], m.subject, m.html, o.template, cfg);
      if (m.teams && cfg.integrations.teamsChat && this.teamsOutlook && this._graph) {
        try {
          await this._graph.teamsMessage(r.email, `<p><strong>${m.subject}</strong></p><p>${m.text.replace(/\n/g, '<br/>')}</p><p><a href="${this.link(o.path)}">Open ${cfg.appName}</a></p>`);
        } catch (e) {
          this.onProblem(`Teams message to ${r.email} not sent: ${(e as Error).message}`);
        }
      }
    }
    try {
      await this._svc.notify(notices);
    } catch (e) {
      this.onProblem('In-app notification not saved: ' + (e as Error).message);
    }
  }

  /** Email straight to an address (support tickets). */
  public async sendRaw(to: string[], template: TemplateKey, vars: Record<string, string | number | undefined>): Promise<void> {
    const m = this.compose(template, vars, '');
    await this._email(to, [], m.subject, m.html, template, this._config());
  }

  private async _email(to: string[], cc: string[], subject: string, html: string, kind: string, cfg: ITimesheetConfig): Promise<void> {
    try {
      if (cfg.integrations.mailMode === 'graph' && this.teamsOutlook && this._graph) await this._graph.sendMail(to, cc, subject, html);
      else await this._svc.queueEmail({ to, cc, subject, html, kind });
    } catch (e) {
      // Outlook not allowed (API access not approved): fall back to the queue so the flow can still send it.
      try {
        await this._svc.queueEmail({ to, cc, subject, html, kind });
      } catch {
        this.onProblem(`Email to ${to.join(', ')} not sent: ${(e as Error).message}`);
      }
    }
  }

  /** Tell the Power Automate flow (HTTP trigger) about an event. Fire and forget. */
  public event(name: string, data: Record<string, unknown>): void {
    const cfg = this._config();
    const url = cfg.integrations.webhookUrl;
    if (!url || cfg.integrations.webhookEvents.indexOf(name) < 0 || !/^https:\/\//i.test(url)) return;
    try {
      // text/plain + no-cors: a simple request, so the browser does not need CORS from the flow endpoint.
      void fetch(url, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify({ event: name, at: new Date().toISOString(), org: cfg.orgName, data }) }).catch(() => undefined);
    } catch {
      // ignore
    }
  }
}
