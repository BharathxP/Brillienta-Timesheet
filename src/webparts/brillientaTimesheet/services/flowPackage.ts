// The Power Automate flow that sends Brillienta Timesheet's emails, as an import package (Power Automate › My flows ›
// Import › Import Package (Legacy)). Pure (unit tested): the admin screen fills in the site and list and zips it.
//
//   When an item is created (SharePoint, <prefix>EmailQueue)
//   └ Check recipient (optional: only addresses at the allowed domains)
//       ├ yes: Send an email (V2) — or Send an email from a shared mailbox (V2)
//       │      ├ succeeded → Update item: Status = Sent
//       │      └ failed / timed out → Update item: Status = Failed, Error = …
//       └ no:  Update item: Status = Failed, Error = "outside the allowed domains"
//
// Standard connectors only (SharePoint, Office 365 Outlook). The connections are chosen while importing, so the
// emails come from whoever imports the flow (or the shared mailbox, if chosen).

import { zipFiles } from './zip';

export interface IFlowOptions {
  /** Site that holds the queue list, e.g. https://contoso.sharepoint.com/sites/people */
  siteUrl: string;
  /** Id (GUID) of the queue list. */
  listId: string;
  listTitle: string;
  displayName: string;
  /** Send from this shared mailbox instead of the importing account ('' = the account). */
  sharedMailbox: string;
  /** Only send to addresses at these domains (empty = any address). */
  allowedDomains: string[];
  /** Pass the Reply to column on (only when a reply-to address is set in the app). */
  useReplyTo: boolean;
}

const SP = 'shared_sharepointonline';
const OUTLOOK = 'shared_office365';
const api = (name: string): string => `/providers/Microsoft.PowerApps/apis/${name}`;
const T = (field: string): string => `@triggerOutputs()?['body/${field}']`;

export function newGuid(random: () => number = Math.random): string {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = Math.floor(random() * 16);
    return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16);
  });
}

/** Domains as typed ("contoso.com, @fabrikam.com") → ["contoso.com", "fabrikam.com"]. */
export function parseDomains(text: string): string[] {
  return (text || '')
    .split(/[\s,;]+/)
    .map((d) => d.trim().toLowerCase().replace(/^@/, ''))
    .filter((d, i, all) => /^[a-z0-9-]+(\.[a-z0-9-]+)+$/.test(d) && all.indexOf(d) === i);
}

function updateItem(o: IFlowOptions, status: 'Sent' | 'Failed', error: string, runAfter: Record<string, string[]>): Record<string, unknown> {
  const parameters: Record<string, string> = {
    dataset: o.siteUrl,
    table: o.listId,
    id: T('ID'),
    'item/Title': T('Title'),
    'item/MailStatus/Value': status
  };
  if (error) parameters['item/MailError'] = error;
  return {
    runAfter,
    type: 'OpenApiConnection',
    inputs: {
      host: { apiId: api(SP), connectionName: SP, operationId: 'PatchItem' },
      parameters,
      authentication: "@parameters('$authentication')"
    }
  };
}

/** The workflow definition (triggers and actions). */
export function flowDefinition(o: IFlowOptions): Record<string, unknown> {
  const shared = !!o.sharedMailbox;
  const send = shared ? 'Send_an_email_from_a_shared_mailbox_(V2)' : 'Send_an_email_(V2)';
  const email: Record<string, string> = {
    'emailMessage/To': T('MailTo'),
    'emailMessage/Subject': T('Title'),
    // The column holds the complete email (HTML), sent as it is.
    'emailMessage/Body': T('MailBody'),
    'emailMessage/Importance': 'Normal'
  };
  if (shared) email['emailMessage/MailboxAddress'] = o.sharedMailbox;
  if (o.useReplyTo) email['emailMessage/ReplyTo'] = T('MailReplyTo');
  const sendAndMark: Record<string, unknown> = {
    [send]: {
      runAfter: {},
      type: 'OpenApiConnection',
      inputs: {
        host: { apiId: api(OUTLOOK), connectionName: OUTLOOK, operationId: shared ? 'SharedMailboxSendEmailV2' : 'SendEmailV2' },
        parameters: email,
        authentication: "@parameters('$authentication')"
      }
    },
    Mark_as_sent: updateItem(o, 'Sent', '', { [send]: ['Succeeded'] }),
    Mark_as_failed: updateItem(o, 'Failed', `Sending failed (status @{outputs('${send}')?['statusCode']}). See the run history of the "${o.displayName}" flow in Power Automate.`, {
      [send]: ['Failed', 'TimedOut']
    })
  };
  let actions: Record<string, unknown> = sendAndMark;
  if (o.allowedDomains.length) {
    actions = {
      Check_recipient: {
        runAfter: {},
        type: 'If',
        expression: {
          and: [
            { not: { contains: [T('MailTo'), ';'] } },
            { or: o.allowedDomains.map((d) => ({ endsWith: [`@toLower(trim(${T('MailTo').substring(1)}))`, '@' + d] })) }
          ]
        },
        actions: sendAndMark,
        else: { actions: { Mark_as_rejected: updateItem(o, 'Failed', `Not sent: the address is outside the allowed domains (${o.allowedDomains.join(', ')}).`, {}) } }
      }
    };
  }
  return {
    $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
    contentVersion: '1.0.0.0',
    parameters: {
      $connections: { defaultValue: {}, type: 'Object' },
      $authentication: { defaultValue: {}, type: 'SecureObject' }
    },
    triggers: {
      When_an_item_is_created: {
        recurrence: { interval: 1, frequency: 'Minute' },
        splitOn: "@triggerOutputs()?['body/value']",
        type: 'OpenApiConnection',
        inputs: {
          host: { apiId: api(SP), connectionName: SP, operationId: 'GetOnNewItems' },
          parameters: { dataset: o.siteUrl, table: o.listId },
          authentication: "@parameters('$authentication')"
        }
      }
    },
    actions
  };
}

/** Files of the legacy import package (paths inside the zip → JSON text). */
export function flowPackageFiles(o: IFlowOptions, now: Date = new Date(), random: () => number = Math.random): Array<{ path: string; content: string }> {
  const flowId = newGuid(random);
  const res = { flow: newGuid(random), spApi: newGuid(random), spConn: newGuid(random), olApi: newGuid(random), olConn: newGuid(random) };
  const description = `Sends Brillienta Timesheet's emails: every item added to ${o.listTitle} on ${o.siteUrl} is emailed and marked Sent or Failed.`;
  const apiResource = (name: string, displayName: string): Record<string, unknown> => ({
    id: api(name),
    name,
    type: 'Microsoft.PowerApps/apis',
    suggestedCreationType: 'Existing',
    details: { displayName },
    configurableBy: 'System',
    hierarchy: 'Child',
    dependsOn: []
  });
  const connResource = (displayName: string, apiRes: string): Record<string, unknown> => ({
    type: 'Microsoft.PowerApps/apis/connections',
    suggestedCreationType: 'Existing',
    creationType: 'Existing',
    details: { displayName },
    configurableBy: 'User',
    hierarchy: 'Child',
    dependsOn: [apiRes]
  });
  const manifest = {
    schema: '1.0',
    details: { displayName: o.displayName, description, createdTime: now.toISOString(), packageTelemetryId: newGuid(random), creator: 'Brillienta Timesheet', sourceEnvironment: '' },
    resources: {
      [res.flow]: {
        id: `/providers/Microsoft.Flow/flows/${flowId}`,
        name: flowId,
        type: 'Microsoft.Flow/flows',
        suggestedCreationType: 'New',
        creationType: 'Existing, New, Update',
        details: { displayName: o.displayName },
        configurableBy: 'User',
        hierarchy: 'Root',
        dependsOn: [res.spApi, res.spConn, res.olApi, res.olConn]
      },
      [res.spApi]: apiResource(SP, 'SharePoint'),
      [res.spConn]: connResource('SharePoint', res.spApi),
      [res.olApi]: apiResource(OUTLOOK, 'Office 365 Outlook'),
      [res.olConn]: connResource('Office 365 Outlook', res.olApi)
    }
  };
  const definition = {
    name: flowId,
    id: `/providers/Microsoft.Flow/flows/${flowId}`,
    type: 'Microsoft.Flow/flows',
    properties: {
      apiId: '/providers/Microsoft.PowerApps/apis/shared_logicflows',
      displayName: o.displayName,
      definition: flowDefinition(o),
      connectionReferences: {
        [SP]: { connectionName: SP, source: 'Embedded', id: api(SP), tier: 'NotSpecified' },
        [OUTLOOK]: { connectionName: OUTLOOK, source: 'Embedded', id: api(OUTLOOK), tier: 'NotSpecified' }
      },
      flowFailureAlertSubscribed: false,
      isManaged: false
    }
  };
  const json = (v: unknown): string => JSON.stringify(v, null, 2);
  return [
    { path: 'manifest.json', content: json(manifest) },
    { path: 'Microsoft.Flow/flows/manifest.json', content: json({ packageSchemaVersion: '1.0', flowAssets: { assetPaths: [flowId] } }) },
    { path: `Microsoft.Flow/flows/${flowId}/definition.json`, content: json(definition) },
    { path: `Microsoft.Flow/flows/${flowId}/apisMap.json`, content: json({ [SP]: res.spApi, [OUTLOOK]: res.olApi }) },
    { path: `Microsoft.Flow/flows/${flowId}/connectionsMap.json`, content: json({ [SP]: res.spConn, [OUTLOOK]: res.olConn }) }
  ];
}

/** The import package as zip bytes. */
export function flowPackageZip(o: IFlowOptions, now: Date = new Date()): Uint8Array {
  return zipFiles(flowPackageFiles(o, now), now);
}
