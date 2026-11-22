import * as React from 'react';
import { DefaultButton, Icon, IContextualMenuItem, TextField } from '@fluentui/react';
import { useApp } from '../AppContext';
import { Avatar } from './Ui';
import { IGraphPerson } from '../../services/GraphService';
import { IXlsxSheet, xlsxWorkbook } from '../../services/xlsx';
import { download, slug, toCsv } from '../../services/format';
import { sameEmail } from '../../services/access';

interface IOption {
  email: string;
  name: string;
  sub: string;
  person?: IGraphPerson;
}

/**
 * Pick one person: people already in the timesheet first, then Microsoft 365 (Graph) when allowed. Anyone can also
 * type a full email address (external users).
 */
export function PersonPicker(props: { label?: string; placeholder?: string; onPick: (p: { email: string; name: string; person?: IGraphPerson }) => void; exclude?: string[]; onlyKnown?: boolean }): JSX.Element {
  const { ref, graph } = useApp();
  const [q, setQ] = React.useState('');
  const [remote, setRemote] = React.useState<IGraphPerson[]>([]);
  const [open, setOpen] = React.useState(false);
  const [active, setActive] = React.useState(0);
  const box = React.useRef<HTMLDivElement>(null);
  const exclude = props.exclude || [];
  React.useEffect(() => {
    if (!graph || props.onlyKnown || q.trim().length < 2) {
      setRemote([]);
      return;
    }
    let alive = true;
    const t = window.setTimeout(() => {
      graph
        .searchPeople(q)
        .then((r) => alive && setRemote(r))
        .catch(() => alive && setRemote([]));
    }, 250);
    return () => {
      alive = false;
      window.clearTimeout(t);
    };
  }, [q, graph, props.onlyKnown]);
  React.useEffect(() => {
    const onClick = (e: MouseEvent): void => {
      if (box.current && !box.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener('mousedown', onClick);
    return () => document.removeEventListener('mousedown', onClick);
  }, []);
  const needle = q.trim().toLowerCase();
  const known: IOption[] = ref.users
    .filter((u) => u.active && (!needle || u.name.toLowerCase().indexOf(needle) >= 0 || u.email.toLowerCase().indexOf(needle) >= 0))
    .slice(0, 8)
    .map((u) => ({ email: u.email, name: u.name, sub: [u.jobTitle, u.department, u.external ? 'External' : ''].filter((x) => !!x).join(' · ') || u.email }));
  const options: IOption[] = known.concat(
    remote.filter((r) => !known.some((k) => sameEmail(k.email, r.email))).map((r) => ({ email: r.email, name: r.name, sub: [r.jobTitle, r.department, r.external ? 'Guest' : ''].filter((x) => !!x).join(' · ') || r.email, person: r }))
  );
  const typed = /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(q.trim()) && !options.some((o) => sameEmail(o.email, q.trim()));
  if (typed && !props.onlyKnown) options.push({ email: q.trim(), name: q.trim(), sub: 'Use this email address' });
  const shown = options.filter((o) => !exclude.some((e) => sameEmail(e, o.email)));
  const pick = (o: IOption): void => {
    props.onPick({ email: o.email, name: o.name, person: o.person });
    setQ('');
    setOpen(false);
  };
  return (
    <div className="ts-picker" ref={box}>
      <TextField
        label={props.label}
        placeholder={props.placeholder || 'Search people by name or email'}
        value={q}
        iconProps={{ iconName: 'Search' }}
        onChange={(_, v) => {
          setQ(v || '');
          setOpen(true);
          setActive(0);
        }}
        onFocus={() => setOpen(true)}
        onKeyDown={(e) => {
          if (e.key === 'ArrowDown') setActive((a) => Math.min(shown.length - 1, a + 1));
          else if (e.key === 'ArrowUp') setActive((a) => Math.max(0, a - 1));
          else if (e.key === 'Enter' && shown[active]) {
            e.preventDefault();
            pick(shown[active]);
          } else if (e.key === 'Escape') setOpen(false);
        }}
      />
      {open && (needle || known.length > 0) && shown.length > 0 && (
        <div className="ts-picker-menu" role="listbox">
          {shown.map((o, i) => (
            <button key={o.email} role="option" aria-selected={i === active} className={'ts-picker-item' + (i === active ? ' is-active' : '')} onMouseEnter={() => setActive(i)} onClick={() => pick(o)}>
              <Avatar name={o.name} email={o.email} />
              <span style={{ minWidth: 0 }}>
                <div>{o.name}</div>
                <div className="ts-picker-sub">{o.sub}</div>
              </span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

/** Several people (project managers, members, assignees, approvers) as removable chips plus a picker. */
export function PeopleField(props: { label: string; value: string[]; onChange: (v: string[]) => void; hint?: string; onlyKnown?: boolean }): JSX.Element {
  const { ref } = useApp();
  const name = (e: string): string => (ref.users.filter((u) => sameEmail(u.email, e))[0] || { name: e }).name;
  return (
    <div>
      <PersonPicker label={props.label} onPick={(p) => props.onChange(props.value.concat(p.email))} exclude={props.value} onlyKnown={props.onlyKnown} />
      {props.value.length > 0 && (
        <div className="ts-chips" style={{ marginTop: 8 }}>
          {props.value.map((e) => (
            <span key={e} className="ts-chip" title={e}>
              {name(e)}
              <button aria-label={`Remove ${name(e)}`} onClick={() => props.onChange(props.value.filter((x) => x !== e))}>
                <Icon iconName="Cancel" />
              </button>
            </span>
          ))}
        </div>
      )}
      {props.hint && <div className="ts-muted ts-small" style={{ marginTop: 4 }}>{props.hint}</div>}
    </div>
  );
}

/** Excel (default) or CSV export of one or more sheets. */
export function ExportButton(props: { name: string; sheets: () => IXlsxSheet[]; disabled?: boolean; text?: string }): JSX.Element {
  const { has, toast } = useApp();
  if (!has('export')) return <></>;
  const excel = (): void => {
    try {
      download(`${slug(props.name)}.xlsx`, xlsxWorkbook(props.sheets()), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } catch (e) {
      toast('Export failed: ' + (e as Error).message, 'error');
    }
  };
  const csv = (): void => {
    const s = props.sheets()[0];
    download(`${slug(props.name)}.csv`, toCsv([s.columns as Array<string | number | boolean | undefined>].concat(s.rows.map((r) => r.map((v) => (v === null ? '' : v)) as Array<string | number | boolean | undefined>))), 'text/csv;charset=utf-8');
  };
  const items: IContextualMenuItem[] = [
    { key: 'xlsx', text: 'Excel workbook (.xlsx)', iconProps: { iconName: 'ExcelDocument' }, onClick: excel },
    { key: 'csv', text: 'CSV (first table)', iconProps: { iconName: 'TextDocument' }, onClick: csv }
  ];
  return <DefaultButton text={props.text || 'Export'} iconProps={{ iconName: 'Download' }} disabled={props.disabled} split menuProps={{ items }} onClick={excel} />;
}
