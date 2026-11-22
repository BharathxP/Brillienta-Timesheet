import * as React from 'react';
import { useApp } from '../AppContext';
import { IReportFilter, ITimeEntry, ITimesheet } from '../../models/types';
import { IFact, ILookups, toFacts } from '../../services/reports';
import { addDays, weekStartOf } from '../../services/dates';
import { billRate, costRate, sameEmail } from '../../services/access';

export interface IFactsState {
  facts: IFact[];
  sheets: ITimesheet[];
  entries: ITimeEntry[];
  lookups: ILookups;
  loading: boolean;
  error?: string;
  reload: () => void;
}

/** Load the weeks overlapping the filter's dates and turn their lines into report facts. */
export function useFacts(filter: Partial<IReportFilter>, opts?: { projectId?: number }): IFactsState {
  const { svc, ref, config, has } = useApp();
  const [raw, setRaw] = React.useState<{ sheets: ITimesheet[]; entries: ITimeEntry[] }>({ sheets: [], entries: [] });
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | undefined>();
  const [tick, setTick] = React.useState(0);
  const from = filter.from || '';
  const to = filter.to || '';
  const people = (filter.userEmails || []).join(',');
  React.useEffect(() => {
    let alive = true;
    setLoading(true);
    setError(undefined);
    const load = opts && opts.projectId ? svc.entriesOfProject(opts.projectId) : svc.weeks(from ? addDays(weekStartOf(from, config.weekStart), -6) : '0000', to || '9999', people ? people.split(',') : undefined);
    load
      .then((r) => alive && setRaw(r))
      .catch((e) => alive && setError((e as Error).message))
      .finally(() => alive && setLoading(false));
    return () => {
      alive = false;
    };
  }, [from, to, people, opts && opts.projectId, tick]);
  const lookups: ILookups = React.useMemo(
    () => ({
      ...ref,
      // Lines not submitted yet have no captured rates: use today's rates.
      rateFor: (e: ITimeEntry) => {
        const u = ref.users.filter((x) => sameEmail(x.email, e.userEmail))[0];
        const p = ref.projects.filter((x) => x.id === e.projectId)[0];
        return { cost: costRate(config.costMode, u, p, config.defaultCostRate), bill: billRate(u, p, config.defaultBillRate, has('userRates')) };
      }
    }),
    [ref, config]
  );
  const facts = React.useMemo(() => toFacts(raw.entries, raw.sheets, lookups, filter, config.weekStart), [raw, lookups, JSON.stringify(filter)]);
  return { facts, sheets: raw.sheets, entries: raw.entries, lookups, loading, error, reload: () => setTick((t) => t + 1) };
}
