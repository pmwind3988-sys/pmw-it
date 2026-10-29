// The time charts: "Raised vs closed" (two lines) and "Still open" (one area).
// The plot is an SVG stretched to the card; its labels, dots and tooltip are
// positioned in percent over it, so text never stretches with the picture.
import * as React from 'react';
import { Bucket, Grain } from '../stats/stats';
import {
  anchorShift, axisTicks, labelEvery, moveIndex, niceMax, PALETTE, snapIndex, svgArea, svgPoints, svgY, VIEW_H,
  VIEW_W, xPct, yPct
} from './chartMath';
import styles from './Dashboard.module.scss';

export interface TrendPlotProps {
  buckets: Bucket[];
  kind: 'lines' | 'area';
  /** Read out by screen readers for the focusable plot. */
  label: string;
}

const tipWords = (b: Bucket, kind: 'lines' | 'area'): string =>
  kind === 'lines'
    ? `${b.raised} raised · ${b.closed} closed · ${b.openAtEnd} still open`
    : `${b.openAtEnd} still open`;

function indexOfMax(values: number[]): number {
  let best = 0;
  values.forEach((v, i) => {
    if (v > values[best]) best = i;
  });
  return best;
}

export function TrendPlot(props: TrendPlotProps): React.ReactElement {
  const { buckets, kind, label } = props;
  const n = buckets.length;
  const [hover, setHover] = React.useState<number | undefined>(undefined);

  const raised = buckets.map((b) => b.raised);
  const closed = buckets.map((b) => b.closed);
  const open = buckets.map((b) => b.openAtEnd);
  const top = niceMax(Math.max(0, ...(kind === 'lines' ? raised.concat(closed) : open)));
  const ticks = axisTicks(top);
  const every = labelEvery(n);
  // Clamped while drawing: a shorter window must not leave a crosshair past its last point.
  const at = hover !== undefined && hover < n ? hover : undefined;

  if (n === 0) return <p className={styles.empty}>No tickets match these filters.</p>;

  const onMove = (e: React.MouseEvent<HTMLDivElement>): void => {
    const box = e.currentTarget.getBoundingClientRect();
    setHover(snapIndex((e.clientX - box.left) / box.width, n));
  };
  const onKey = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
    e.preventDefault();
    const step = e.key === 'ArrowLeft' ? -1 : 1;
    setHover((cur) => moveIndex(cur, n, step));
  };

  const last = n - 1;
  const peakOpen = indexOfMax(open);
  const peakRaised = indexOfMax(raised);
  const shownLabels = buckets.map((b, i) => ({ b, i })).filter((x) => x.i % every === 0);
  const hoverBucket = at !== undefined ? buckets[at] : undefined;
  const crossX = at !== undefined ? xPct(at, n) : 0;

  return (
    <div className={styles.chartBlock}>
      <div className={styles.plotWrap}>
        <div className={styles.yAxis}>
          {ticks.map((t) => (
            <span key={t} className={styles.yLabel} style={{ top: `${yPct(t, top)}%` }}>
              {t}
            </span>
          ))}
        </div>

        <div
          className={styles.plot}
          role="group"
          aria-label={label}
          tabIndex={0}
          onMouseMove={onMove}
          onMouseLeave={() => setHover(undefined)}
          onFocus={() => setHover(last)}
          onBlur={() => setHover(undefined)}
          onKeyDown={onKey}
        >
          <svg
            className={styles.svg}
            viewBox={`0 0 ${VIEW_W} ${VIEW_H}`}
            preserveAspectRatio="none"
            aria-hidden="true"
            focusable="false"
          >
            {ticks.slice(1).map((t) => (
              <line
                key={t}
                x1={0}
                x2={VIEW_W}
                y1={svgY(t, top)}
                y2={svgY(t, top)}
                stroke={PALETTE.hairline}
                strokeWidth={1}
                vectorEffect="non-scaling-stroke"
              />
            ))}
            <line
              x1={0}
              x2={VIEW_W}
              y1={VIEW_H}
              y2={VIEW_H}
              stroke={PALETTE.baseline}
              strokeWidth={1}
              vectorEffect="non-scaling-stroke"
            />
            {kind === 'area' && <path d={svgArea(open, top)} fill={PALETTE.open} />}
            {kind === 'lines' && (
              <polyline
                points={svgPoints(closed, top)}
                fill="none"
                stroke={PALETTE.closed}
                strokeWidth={2}
                strokeLinejoin="round"
                strokeLinecap="round"
                vectorEffect="non-scaling-stroke"
              />
            )}
            {kind === 'lines' && (
              <polyline
                points={svgPoints(raised, top)}
                fill="none"
                stroke={PALETTE.raised}
                strokeWidth={2}
                strokeLinejoin="round"
                strokeLinecap="round"
                vectorEffect="non-scaling-stroke"
              />
            )}
            {kind === 'area' && (
              <polyline
                points={svgPoints(open, top)}
                fill="none"
                stroke={PALETTE.openLine}
                strokeWidth={2}
                strokeLinejoin="round"
                strokeLinecap="round"
                vectorEffect="non-scaling-stroke"
              />
            )}
          </svg>

          {at !== undefined && <div className={styles.crosshair} style={{ left: `${crossX}%` }} />}

          {kind === 'lines' && (
            <>
              <span className={styles.dot} style={{ left: `${xPct(last, n)}%`, top: `${yPct(raised[last], top)}%`, background: PALETTE.raised }} />
              <span className={styles.dot} style={{ left: `${xPct(last, n)}%`, top: `${yPct(closed[last], top)}%`, background: PALETTE.closed }} />
              <span
                className={styles.tag}
                style={{
                  left: `${xPct(peakRaised, n)}%`,
                  top: `${yPct(raised[peakRaised], top)}%`,
                  transform: `translate(${anchorShift(xPct(peakRaised, n))}, -150%)`
                }}
              >
                {buckets[peakRaised].label}: {raised[peakRaised]} raised
              </span>
            </>
          )}

          {kind === 'area' && (
            <>
              <span className={styles.dot} style={{ left: `${xPct(last, n)}%`, top: `${yPct(open[last], top)}%`, background: PALETTE.openLine }} />
              <span
                className={styles.tag}
                style={{
                  left: `${xPct(last, n)}%`,
                  top: `${yPct(open[last], top)}%`,
                  transform: `translate(${anchorShift(xPct(last, n))}, -150%)`
                }}
              >
                {open[last]} now
              </span>
              {peakOpen !== last && (
                <span
                  className={styles.tag}
                  style={{
                    left: `${xPct(peakOpen, n)}%`,
                    top: `${yPct(open[peakOpen], top)}%`,
                    transform: `translate(${anchorShift(xPct(peakOpen, n))}, -150%)`
                  }}
                >
                  {open[peakOpen]} peak
                </span>
              )}
            </>
          )}

          {hoverBucket !== undefined && at !== undefined && (
            <>
              {kind === 'lines' && (
                <>
                  <span className={styles.dot} style={{ left: `${crossX}%`, top: `${yPct(raised[at], top)}%`, background: PALETTE.raised }} />
                  <span className={styles.dot} style={{ left: `${crossX}%`, top: `${yPct(closed[at], top)}%`, background: PALETTE.closed }} />
                </>
              )}
              <div
                className={styles.tip}
                style={{ left: `${crossX}%`, transform: crossX > 60 ? 'translateX(calc(-100% - 10px))' : 'translateX(10px)' }}
              >
                <strong>{hoverBucket.label}</strong>
                <span>{tipWords(hoverBucket, kind)}</span>
              </div>
            </>
          )}
        </div>
      </div>

      <div className={styles.xAxis}>
        {shownLabels.map((x) => (
          <span
            key={x.i}
            className={styles.xLabel}
            style={{ left: `${xPct(x.i, n)}%`, transform: `translateX(${anchorShift(xPct(x.i, n))})` }}
          >
            {x.b.label}
          </span>
        ))}
      </div>
    </div>
  );
}

/** Table twin of the time charts: the same numbers, row by row. */
export function TrendTable(props: { buckets: Bucket[]; grain: Grain; kind: 'lines' | 'area' }): React.ReactElement {
  const { buckets, grain, kind } = props;
  const unit = grain === 'month' ? 'Month' : grain === 'week' ? 'Week starting' : 'Day';
  return (
    <table className={styles.table}>
      <thead>
        <tr>
          <th scope="col">{unit}</th>
          {kind === 'lines' && <th scope="col">Raised</th>}
          {kind === 'lines' && <th scope="col">Closed</th>}
          <th scope="col">Still open</th>
        </tr>
      </thead>
      <tbody>
        {buckets.map((b) => (
          <tr key={b.label + b.start.getTime()}>
            <th scope="row">{b.label}</th>
            {kind === 'lines' && <td>{b.raised}</td>}
            {kind === 'lines' && <td>{b.closed}</td>}
            <td>{b.openAtEnd}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}
