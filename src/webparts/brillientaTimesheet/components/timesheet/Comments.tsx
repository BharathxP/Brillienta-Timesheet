import * as React from 'react';
import { Icon } from '@fluentui/react';
import { ISheetComment } from '../../models/types';
import { formatDateTime } from '../../services/format';

const ICONS: Record<string, string> = { submitted: 'Send', approved: 'CheckMark', 'auto-approved': 'Robot', rejected: 'Undo', reopened: 'Refresh', comment: 'Comment' };
const VERBS: Record<string, string> = { submitted: 'submitted', approved: 'approved', 'auto-approved': 'approved automatically', rejected: 'sent it back', reopened: 'reopened it', comment: 'commented' };

/** The approval history of a week (who submitted, approved, rejected, with their comments). */
export function Comments(props: { comments: ISheetComment[] }): JSX.Element {
  return (
    <div className="ts-timeline">
      {props.comments
        .slice()
        .reverse()
        .map((c, i) => (
          <div key={i} className="ts-timeline-item">
            <span className="ts-timeline-dot">
              <Icon iconName={ICONS[c.action || 'comment'] || 'Comment'} />
            </span>
            <div>
              <div>
                <strong>{c.by}</strong> {VERBS[c.action || 'comment'] || ''}
              </div>
              {c.text && <div>{c.text}</div>}
              <div className="ts-muted ts-small">{formatDateTime(c.at)}</div>
            </div>
          </div>
        ))}
    </div>
  );
}
