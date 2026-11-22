// Microsoft Graph: people search and import, external (guest) invitations, Outlook email, Teams chat messages and
// Planner. Every call is delegated (it runs as the signed-in person) with the permissions an admin approved for the
// solution in the SharePoint admin center (API access).

import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';

export interface IGraphPerson {
  id: string;
  name: string;
  email: string;
  department: string;
  jobTitle: string;
  external: boolean;
  managerEmail?: string;
}

export interface IPlannerPlan {
  id: string;
  title: string;
  owner: string;
}

export interface IPlannerTask {
  id: string;
  title: string;
  percentComplete: number;
  dueDate: string;
  assigneeIds: string[];
  bucketId: string;
  etag: string;
}

interface IGraphUser {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName?: string;
  department?: string;
  jobTitle?: string;
  userType?: string;
}

const USER_SELECT = 'id,displayName,mail,userPrincipalName,department,jobTitle,userType';

function toPerson(u: IGraphUser): IGraphPerson {
  return {
    id: u.id,
    name: u.displayName,
    email: u.mail || u.userPrincipalName || '',
    department: u.department || '',
    jobTitle: u.jobTitle || '',
    external: u.userType === 'Guest' || /#EXT#/i.test(u.userPrincipalName || '')
  };
}

export class GraphService {
  private readonly _factory: MSGraphClientFactory;
  private _client: Promise<MSGraphClientV3> | undefined;
  private _me: Promise<{ id: string; mail: string }> | undefined;
  private readonly _chats: Record<string, string> = {};

  public constructor(factory: MSGraphClientFactory) {
    this._factory = factory;
  }

  public client(): Promise<MSGraphClientV3> {
    if (!this._client) this._client = this._factory.getClient('3');
    return this._client;
  }

  private async _get<T>(path: string, headers?: Record<string, string>): Promise<T> {
    const c = await this.client();
    let req = c.api(path);
    Object.keys(headers || {}).forEach((h) => (req = req.header(h, (headers as Record<string, string>)[h])));
    return req.get();
  }

  private async _all<T>(path: string, max: number = 2000, headers?: Record<string, string>): Promise<T[]> {
    const out: T[] = [];
    let next: string | undefined = path;
    while (next && out.length < max) {
      const page: { value: T[]; '@odata.nextLink'?: string } = await this._get(next, headers);
      out.push(...(page.value || []));
      next = page['@odata.nextLink'];
    }
    return out;
  }

  public me(): Promise<{ id: string; mail: string }> {
    if (!this._me) this._me = this._get<{ id: string; mail: string; userPrincipalName: string }>('/me?$select=id,mail,userPrincipalName').then((m) => ({ id: m.id, mail: m.mail || m.userPrincipalName }));
    return this._me;
  }

  // ---------------------------------------------------------------- people

  /** People whose name or email starts with the text (members and guests). */
  public async searchPeople(text: string, top: number = 12): Promise<IGraphPerson[]> {
    const q = text.replace(/"/g, '').trim();
    if (!q) return [];
    try {
      const res = await this._get<{ value: IGraphUser[] }>(`/users?$search="displayName:${q}" OR "mail:${q}"&$select=${USER_SELECT}&$top=${top}&$count=true`, { ConsistencyLevel: 'eventual' });
      return res.value.map(toPerson);
    } catch {
      const esc = q.replace(/'/g, "''");
      const res = await this._get<{ value: IGraphUser[] }>(`/users?$filter=startswith(displayName,'${esc}') or startswith(mail,'${esc}')&$select=${USER_SELECT}&$top=${top}`);
      return res.value.map(toPerson);
    }
  }

  /** One person with their manager's email (import, "fill from Microsoft 365"). */
  public async person(email: string): Promise<IGraphPerson | undefined> {
    try {
      const u = await this._get<IGraphUser>(`/users/${encodeURIComponent(email)}?$select=${USER_SELECT}`);
      const p = toPerson(u);
      try {
        const m = await this._get<IGraphUser>(`/users/${u.id}/manager?$select=mail,userPrincipalName`);
        p.managerEmail = m.mail || m.userPrincipalName || '';
      } catch {
        // no manager
      }
      return p;
    } catch {
      return undefined;
    }
  }

  /** Members of a Microsoft 365 group / team (bulk add people). */
  public async groupMembers(groupId: string): Promise<IGraphPerson[]> {
    const rows = await this._all<IGraphUser>(`/groups/${groupId}/transitiveMembers/microsoft.graph.user?$select=${USER_SELECT}&$top=999`, 5000);
    return rows.map(toPerson).filter((p) => !!p.email);
  }

  /** Groups the signed-in person belongs to (group pickers for people import and Planner). */
  public async myGroups(): Promise<Array<{ id: string; name: string }>> {
    const rows = await this._all<{ id: string; displayName: string; groupTypes?: string[] }>('/me/memberOf/microsoft.graph.group?$select=id,displayName,groupTypes&$top=200', 1000);
    return rows.map((g) => ({ id: g.id, name: g.displayName })).sort((a, b) => a.name.localeCompare(b.name));
  }

  /** Invite an external person as a Microsoft 365 guest (needs User.Invite.All, or guest inviter rights). */
  public async inviteGuest(email: string, name: string, redirectUrl: string, message: string): Promise<{ id: string; redeemUrl: string }> {
    const c = await this.client();
    const res = await c.api('/invitations').post({
      invitedUserEmailAddress: email,
      invitedUserDisplayName: name,
      inviteRedirectUrl: redirectUrl,
      sendInvitationMessage: true,
      invitedUserMessageInfo: { customizedMessageBody: message }
    });
    return { id: res.invitedUser && res.invitedUser.id, redeemUrl: res.inviteRedeemUrl };
  }

  // ---------------------------------------------------------------- Outlook and Teams

  /** Send an email from the signed-in person's mailbox (Outlook connectivity). */
  public async sendMail(to: string[], cc: string[], subject: string, html: string): Promise<void> {
    const c = await this.client();
    const addr = (list: string[]): Array<{ emailAddress: { address: string } }> => list.map((a) => ({ emailAddress: { address: a } }));
    await c.api('/me/sendMail').post({ message: { subject, body: { contentType: 'HTML', content: html }, toRecipients: addr(to), ccRecipients: addr(cc) }, saveToSentItems: false });
  }

  /** Post a message to a person in a one-to-one Teams chat with the signed-in person. */
  public async teamsMessage(toEmail: string, html: string): Promise<void> {
    const me = await this.me();
    if (me.mail && me.mail.toLowerCase() === toEmail.toLowerCase()) return;
    const c = await this.client();
    let chatId = this._chats[toEmail.toLowerCase()];
    if (!chatId) {
      const chat = await c.api('/chats').post({
        chatType: 'oneOnOne',
        members: [
          { '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: ['owner'], 'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${me.id}')` },
          { '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: ['owner'], 'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${toEmail}')` }
        ]
      });
      chatId = this._chats[toEmail.toLowerCase()] = chat.id;
    }
    await c.api(`/chats/${chatId}/messages`).post({ body: { contentType: 'html', content: html } });
  }

  // ---------------------------------------------------------------- Planner

  /** Plans of the groups the signed-in person belongs to. */
  public async plans(): Promise<IPlannerPlan[]> {
    const groups = await this.myGroups();
    const out: IPlannerPlan[] = [];
    for (const g of groups.slice(0, 60)) {
      try {
        const res = await this._get<{ value: Array<{ id: string; title: string }> }>(`/groups/${g.id}/planner/plans?$select=id,title`);
        res.value.forEach((p) => out.push({ id: p.id, title: p.title, owner: g.name }));
      } catch {
        // group without Planner
      }
    }
    return out;
  }

  public async planTasks(planId: string): Promise<IPlannerTask[]> {
    const rows = await this._all<{ id: string; title: string; percentComplete: number; dueDateTime?: string; assignments?: Record<string, unknown>; bucketId: string; '@odata.etag': string }>(
      `/planner/plans/${planId}/tasks`,
      2000
    );
    return rows.map((t) => ({
      id: t.id,
      title: t.title,
      percentComplete: t.percentComplete,
      dueDate: t.dueDateTime ? t.dueDateTime.substring(0, 10) : '',
      assigneeIds: Object.keys(t.assignments || {}),
      bucketId: t.bucketId,
      etag: t['@odata.etag']
    }));
  }

  /** Emails of Entra user ids (Planner assignees). */
  public async emailsOf(ids: string[]): Promise<Record<string, string>> {
    const out: Record<string, string> = {};
    const unique = ids.filter((x, i, a) => a.indexOf(x) === i);
    for (let i = 0; i < unique.length; i += 15) {
      const chunk = unique.slice(i, i + 15);
      const res = await this._get<{ value: IGraphUser[] }>(`/users?$filter=id in (${chunk.map((x) => `'${x}'`).join(',')})&$select=id,mail,userPrincipalName`);
      res.value.forEach((u) => (out[u.id] = u.mail || u.userPrincipalName || ''));
    }
    return out;
  }

  /** Mark a Planner task complete (or reopen it). */
  public async setPlannerProgress(taskId: string, percentComplete: number): Promise<void> {
    const c = await this.client();
    const t = await c.api(`/planner/tasks/${taskId}`).get();
    await c.api(`/planner/tasks/${taskId}`).header('If-Match', t['@odata.etag']).patch({ percentComplete });
  }
}
