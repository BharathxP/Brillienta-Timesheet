import * as React from 'react';

// Categorical palette (validated for colour-vision deficiency: adjacent pairs ΔE ≥ 9). Assigned in this fixed order,
// never cycled — more than eight categories fold into 'Other'. Every chart also shows labels and values.
export const PALETTE = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#4a3aa7', '#e34948'];
/** Single-series magnitude colour. */
export const SERIES_COLOR = '#2a78d6';

export interface IDatum {
  label: string;
  value: number;
  color?: string;
}

export function BarList(props: { items: IDatum[]; format?: (v: number) => string; onClick?: (d: IDatum) => void; empty?: string }): JSX.Element {
  const max = Math.max(1, ...props.items.map((i) => i.value));
  if (!props.items.length) return <div className="ts-muted ts-small">{props.empty || 'No data'}</div>;
  return (
    <div className="ts-bars">
      {props.items.map((d, i) => (
        <div
          key={d.label + i}
          className="ts-bar-row"
          onClick={props.onClick ? () => props.onClick && props.onClick(d) : undefined}
          style={props.onClick ? { cursor: 'pointer' } : undefined}
        >
          <span className="ts-bar-label" title={d.label}>
            {d.label}
          </span>
          <span className="ts-bar-track">
            <span className="ts-bar-fill" style={{ display: 'block', width: `${(d.value / max) * 100}%`, background: d.color || SERIES_COLOR }} />
          </span>
          <span className="ts-bar-value">{props.format ? props.format(d.value) : d.value}</span>
        </div>
      ))}
    </div>
  );
}

export interface ISeries {
  label: string;
  color: string;
  values: number[];
}

/** Multi-series line/area chart over evenly spaced categories (e.g. days). */
export function LineChart(props: { labels: string[]; series: ISeries[]; height?: number }): JSX.Element {
  const w = 640;
  const h = props.height || 200;
  const pad = { l: 30, r: 10, t: 10, b: 24 };
  const n = props.labels.length;
  const max = Math.max(1, ...props.series.map((s) => Math.max(0, ...s.values)));
  const niceMax = Math.ceil(max / 4) * 4 || 4;
  const x = (i: number): number => pad.l + (n <= 1 ? 0 : (i / (n - 1)) * (w - pad.l - pad.r));
  const y = (v: number): number => pad.t + (1 - v / niceMax) * (h - pad.t - pad.b);
  const step = Math.max(1, Math.ceil(n / 8));
  return (
    <div className="ts-chart">
      <svg viewBox={`0 0 ${w} ${h}`} role="img" aria-label={props.series.map((s) => s.label).join(', ')}>
        {[0, 0.25, 0.5, 0.75, 1].map((f) => (
          <g key={f}>
            <line x1={pad.l} x2={w - pad.r} y1={y(niceMax * f)} y2={y(niceMax * f)} stroke="currentColor" strokeOpacity={0.08} />
            <text x={pad.l - 6} y={y(niceMax * f) + 3} textAnchor="end">
              {Math.round(niceMax * f)}
            </text>
          </g>
        ))}
        {props.labels.map((l, i) =>
          i % step === 0 ? (
            <text key={l + i} x={x(i)} y={h - 6} textAnchor="middle">
              {l}
            </text>
          ) : null
        )}
        {props.series.map((s) => {
          const pts = s.values.map((v, i) => `${x(i)},${y(v)}`).join(' ');
          const area = `${x(0)},${y(0)} ${pts} ${x(n - 1)},${y(0)}`;
          return (
            <g key={s.label}>
              <polygon points={area} fill={s.color} fillOpacity={0.08} />
              <polyline points={pts} fill="none" stroke={s.color} strokeWidth={2} strokeLinejoin="round" />
              {s.values.map((v, i) => (
                <circle key={i} cx={x(i)} cy={y(v)} r={n > 40 ? 0 : 2.5} fill={s.color}>
                  <title>{`${props.labels[i]}: ${v} ${s.label}`}</title>
                </circle>
              ))}
            </g>
          );
        })}
      </svg>
      <div className="ts-legend">
        {props.series.map((s) => (
          <span key={s.label}>
            <span className="ts-dot" style={{ background: s.color }} /> {s.label}
          </span>
        ))}
      </div>
    </div>
  );
}

/** Fixts-order categorical colour; the catch-all buckets are neutral grey. */
function sliceColor(d: IDatum, i: number): string {
  if (d.color) return d.color;
  if (d.label === 'Other' || d.label === 'Not set') return '#898781';
  return PALETTE[Math.min(i, PALETTE.length - 1)];
}

export function Donut(props: { items: IDatum[]; size?: number; center?: string; centerSub?: string }): JSX.Element {
  const size = props.size || 150;
  const r = size / 2 - 12;
  const c = 2 * Math.PI * r;
  const total = props.items.reduce((s, d) => s + d.value, 0);
  let offset = 0;
  return (
    <div className="ts-donut">
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} role="img" aria-label={props.items.map((d) => `${d.label}: ${d.value}`).join(', ') || 'No data'}>
        <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke="currentColor" strokeOpacity={0.08} strokeWidth={18} />
        {total > 0 &&
          props.items.map((d, i) => {
            const len = (d.value / total) * c;
            const el = (
              <circle
                key={d.label}
                cx={size / 2}
                cy={size / 2}
                r={r}
                fill="none"
                stroke={sliceColor(d, i)}
                strokeWidth={18}
                // 2px surface gap between slices.
                strokeDasharray={props.items.length > 1 ? `${Math.max(0.5, len - 2)} ${c - Math.max(0.5, len - 2)}` : `${len} ${c - len}`}
                strokeDashoffset={-offset}
                transform={`rotate(-90 ${size / 2} ${size / 2})`}
              >
                <title>{`${d.label}: ${d.value}`}</title>
              </circle>
            );
            offset += len;
            return el;
          })}
        <text x="50%" y="48%" textAnchor="middle" style={{ fontSize: 22, fontWeight: 600, fill: 'var(--ts-text)' }}>
          {props.center !== undefined ? props.center : total}
        </text>
        <text x="50%" y="62%" textAnchor="middle" style={{ fontSize: 11, fill: 'var(--ts-muted)' }}>
          {props.centerSub || 'total'}
        </text>
      </svg>
      <div className="ts-legend" style={{ flexDirection: 'column', gap: 4 }}>
        {props.items.map((d, i) => (
          <span key={d.label}>
            <span className="ts-dot" style={{ background: sliceColor(d, i) }} /> {d.label}
            <strong style={{ marginLeft: 4 }}>{d.value}</strong>
            {total > 0 && <span className="ts-muted">({Math.round((d.value / total) * 100)}%)</span>}
          </span>
        ))}
      </div>
    </div>
  );
}

/** Vertical bars over categories (e.g. months), optionally with a second series drawn as thinner bars. */
export function ColumnChart(props: { labels: string[]; series: ISeries[]; height?: number }): JSX.Element {
  const w = 640;
  const h = props.height || 200;
  const pad = { l: 30, r: 10, t: 12, b: 24 };
  const n = props.labels.length;
  const max = Math.max(1, ...props.series.map((s) => Math.max(0, ...s.values)));
  const niceMax = Math.ceil(max / 4) * 4 || 4;
  const slot = (w - pad.l - pad.r) / Math.max(1, n);
  const groupW = Math.min(34, slot * 0.7);
  const barW = groupW / Math.max(1, props.series.length);
  const y = (v: number): number => pad.t + (1 - v / niceMax) * (h - pad.t - pad.b);
  const step = Math.max(1, Math.ceil(n / 12));
  return (
    <div className="ts-chart">
      <svg viewBox={`0 0 ${w} ${h}`} role="img" aria-label={props.series.map((s) => s.label).join(', ')}>
        {[0, 0.25, 0.5, 0.75, 1].map((f) => (
          <g key={f}>
            <line x1={pad.l} x2={w - pad.r} y1={y(niceMax * f)} y2={y(niceMax * f)} stroke="currentColor" strokeOpacity={0.08} />
            <text x={pad.l - 6} y={y(niceMax * f) + 3} textAnchor="end">
              {Math.round(niceMax * f)}
            </text>
          </g>
        ))}
        {props.labels.map((l, i) => {
          const cx = pad.l + slot * i + slot / 2;
          return (
            <g key={l + i}>
              {props.series.map((s, j) => {
                const v = s.values[i] || 0;
                const x = cx - groupW / 2 + j * barW;
                return (
                  <rect key={s.label} x={x + 1} y={y(v)} width={Math.max(2, barW - 2)} height={Math.max(0, y(0) - y(v))} rx={3} fill={s.color}>
                    <title>{`${l}: ${v} ${s.label}`}</title>
                  </rect>
                );
              })}
              {i % step === 0 && (
                <text x={cx} y={h - 6} textAnchor="middle">
                  {l}
                </text>
              )}
            </g>
          );
        })}
      </svg>
      {props.series.length > 1 && (
        <div className="ts-legend">
          {props.series.map((s) => (
            <span key={s.label}>
              <span className="ts-dot" style={{ background: s.color }} /> {s.label}
            </span>
          ))}
        </div>
      )}
    </div>
  );
}
