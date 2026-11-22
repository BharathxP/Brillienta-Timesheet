import * as React from 'react';
import { DefaultButton, Dropdown, MessageBar, MessageBarType, PrimaryButton, TextField } from '@fluentui/react';
import { useApp } from '../AppContext';
import { ITicket } from '../../models/types';
import { ticketLimit } from '../../services/licensing/plans';
import { formatDateTime } from '../../services/format';
import { Card, Empty, Kpi, Loading, PageHeader, StatusBadge, useBusy, useLoader } from '../common/Ui';

const CATEGORIES = ['Question', 'Problem / bug', 'Setup help', 'Licence & billing', 'Feature request'];

/** Support: tickets to Brillienta (4 a year on Standard, unlimited on Pro and Enterprise) and the account manager. */
export function Support(): JSX.Element {
  const { svc, me, license, notifier, config, has, appVersion, tenantId } = useApp();
  const list = useLoader(() => svc.tickets(), []);
  const [t, setT] = React.useState({ title: '', details: '', category: CATEGORIES[0], priority: 'Normal' as ITicket['priority'] });
  const [busy, run] = useBusy();
  const limit = ticketLimit(license.plan);
  const year = new Date().getFullYear();
  const thisYear = (list.data || []).filter((x) => new Date(x.created).getFullYear() === year);
  const left = limit ? Math.max(0, limit - thisYear.length) : Infinity;
  const am = (license.license && license.license.accountManager) || (config.support.accountManager ? { name: config.support.accountManager, email: config.support.accountManagerEmail, phone: config.support.phone } : undefined);
  const send = (): Promise<void> =>
    run(async () => {
      if (!t.title.trim() || !t.details.trim()) throw new Error('Describe the problem: a subject and the details.');
      if (left <= 0) throw new Error(`Your plan includes ${limit} support tickets a year and they are used. Upgrade to Pro for unlimited support.`);
      await svc.addTicket({ ...t, raisedByEmail: me.email });
      await notifier.sendRaw([config.support.email || 'support@brillienta.com'], 'ticketRaised', {
        person: me.title,
        email: me.email,
        org: `${config.orgName} (tenant ${tenantId || '?'}, v${appVersion}, ${license.plan})`,
        subject: t.title,
        priority: t.priority,
        category: t.category,
        details: t.details
      });
      setT({ title: '', details: '', category: CATEGORIES[0], priority: 'Normal' });
      list.reload();
    }, 'Ticket sent to Brillienta support.');
  return (
    <>
      <PageHeader title="Support" sub="Ask Brillienta for help. Replies come by email to the person who raised the ticket." />
      <div className="ts-kpis">
        <Kpi label="Tickets this year" value={thisYear.length} icon="Headset" tint={2} />
        <Kpi label="Left this year" value={limit ? left : 'Unlimited'} icon="Ticket" tint={1} tone={limit && !left ? 'danger' : undefined} />
        <Kpi label="Updates" value={has('unlimitedSupport') ? 'Free with support' : 'Included'} icon="Sync" tint={5} />
      </div>
      <div className="ts-split">
        <Card title="New ticket" icon="Edit">
          {limit > 0 && (
            <MessageBar messageBarType={left ? MessageBarType.info : MessageBarType.warning} styles={{ root: { marginBottom: 12, borderRadius: 8 } }}>
              The Standard plan includes {limit} support tickets a year ({left} left in {year}). Pro and Enterprise include unlimited support.
            </MessageBar>
          )}
          <div className="ts-stack">
            <TextField label="Subject" required value={t.title} onChange={(_, v) => setT({ ...t, title: v || '' })} />
            <div className="ts-form-grid">
              <Dropdown label="Category" selectedKey={t.category} options={CATEGORIES.map((c) => ({ key: c, text: c }))} onChange={(_, o) => o && setT({ ...t, category: String(o.key) })} />
              <Dropdown label="Priority" selectedKey={t.priority} options={['Low', 'Normal', 'High', 'Urgent'].map((c) => ({ key: c, text: c }))} onChange={(_, o) => o && setT({ ...t, priority: o.key as ITicket['priority'] })} />
            </div>
            <TextField label="Details" required multiline rows={6} value={t.details} onChange={(_, v) => setT({ ...t, details: v || '' })} placeholder="What happened, what you expected, and the steps to see it." />
            <div>
              <PrimaryButton text="Send ticket" iconProps={{ iconName: 'Send' }} onClick={() => void send()} disabled={busy || (limit > 0 && !left)} />
            </div>
          </div>
        </Card>
        <div className="ts-stack">
          <Card title="Your account manager" icon="Contact">
            {has('accountManager') && am ? (
              <div className="ts-stack" style={{ gap: 4 }}>
                <strong>{am.name}</strong>
                {am.email && (
                  <a href={`mailto:${am.email}`} data-interception="off">
                    {am.email}
                  </a>
                )}
                {am.phone && <span>{am.phone}</span>}
              </div>
            ) : has('accountManager') ? (
              <div className="ts-muted ts-small">Brillienta assigns your dedicated account manager when the Enterprise licence starts; their details appear here.</div>
            ) : (
              <div className="ts-muted ts-small">A dedicated account manager comes with Enterprise.</div>
            )}
          </Card>
          <Card title="Contact" icon="Mail">
            <div className="ts-small">
              Email{' '}
              <a href={`mailto:${config.support.email}`} data-interception="off">
                {config.support.email}
              </a>
            </div>
            <div className="ts-small ts-muted" style={{ marginTop: 6 }}>
              Include your tenant id: <code>{tenantId || 'unknown'}</code>
            </div>
          </Card>
        </div>
      </div>
      <Card title="Tickets" icon="History" style={{ marginTop: 16 }}>
        {list.loading ? (
          <Loading />
        ) : !(list.data || []).length ? (
          <Empty icon="Headset" title="No tickets yet" />
        ) : (
          <div className="ts-table-wrap">
            <table className="ts-table">
              <thead>
                <tr>
                  <th>Subject</th>
                  <th>Category</th>
                  <th>Priority</th>
                  <th>Status</th>
                  <th>Raised</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {(list.data || []).map((x) => (
                  <tr key={x.id} style={{ cursor: 'default' }}>
                    <td>
                      <div className="ts-subject">{x.title}</div>
                      {x.reply && <div className="ts-small">Reply: {x.reply}</div>}
                    </td>
                    <td className="ts-small">{x.category}</td>
                    <td className="ts-small">{x.priority}</td>
                    <td>
                      <StatusBadge status={x.status} />
                    </td>
                    <td className="ts-small ts-muted">
                      {formatDateTime(x.created)} · {x.raisedBy}
                    </td>
                    <td style={{ textAlign: 'right' }}>
                      {x.status !== 'Closed' && (
                        <DefaultButton
                          text="Close"
                          onClick={() =>
                            void run(async () => {
                              await svc.updateTicket(x.id, { TicketStatus: 'Closed' });
                              list.reload();
                            })
                          }
                        />
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </>
  );
}
