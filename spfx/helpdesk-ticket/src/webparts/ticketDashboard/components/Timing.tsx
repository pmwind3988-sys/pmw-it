// When tickets arrive, how long they take to close, how old the open ones are,
// and which records are missing what. Each block has a plain-table twin.
import * as React from 'react';
import { Picker } from '../stats/drill';
import { Gap, Heat, HEAT_DAYS, HEAT_HOURS, slotText } from '../stats/stats';
import { barPct, heatColour, PALETTE, percentText, ticketsText } from './chartMath';
import { Mark, Segment } from './Marks';
import styles from './Dashboard.module.scss';

const pickingClass = (pick: Picker): string => (pick.picked !== undefined ? styles.picking : '');

// ---------------------------------------------------------------- when they arrive

export function HeatMap(props: { heat: Heat; pick: Picker }): React.ReactElement {
  const { heat, pick } = props;
  const days = heat.showWeekend ? HEAT_DAYS : HEAT_DAYS.slice(0, 5);
  return (
    <div className={styles.heatScroll}>
      <div className={`${styles.heatGrid} ${pickingClass(pick)}`} role="group" aria-label="Tickets raised by weekday and hour">
        <span />
        {HEAT_HOURS.map((h) => (
          <span key={h} className={styles.heatHead}>
            {h}
          </span>
        ))}
        {days.map((day, d) => (
          <React.Fragment key={day}>
            <span className={styles.heatDay}>{day}</span>
            {heat.grid[d].map((count, c) => (
              <Segment
                key={c}
                pick={pick}
                drill={{ kind: 'heat', day: d, col: c }}
                count={count}
                label={`${slotText(d, c)}: ${ticketsText(count)}`}
                className={styles.heatCell}
                style={{ background: heatColour(count, heat.max, PALETTE.heatRamp, PALETTE.empty) }}
              />
            ))}
          </React.Fragment>
        ))}
      </div>
    </div>
  );
}

export function HeatLegend(props: { heat: Heat }): React.ReactElement {
  const { heat } = props;
  return (
    <div className={styles.heatLegend}>
      <span>Fewer</span>
      {PALETTE.heatRamp.map((c) => (
        <span key={c} className={styles.heatKey} style={{ background: c }} aria-hidden="true" />
      ))}
      <span>More</span>
      {heat.busiest !== undefined && (
        <span className={styles.busiest}>
          Busiest: <strong>{slotText(heat.busiest.day, heat.busiest.col)}</strong>
        </span>
      )}
    </div>
  );
}

export function HeatTable(props: { heat: Heat }): React.ReactElement {
  const { heat } = props;
  const days = heat.showWeekend ? HEAT_DAYS : HEAT_DAYS.slice(0, 5);
  return (
    <table className={styles.table}>
      <thead>
        <tr>
          <th scope="col">Day</th>
          {HEAT_HOURS.map((h) => (
            <th key={h} scope="col">
              {h}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {days.map((day, d) => (
          <tr key={day}>
            <th scope="row">{day}</th>
            {heat.grid[d].map((count, c) => (
              <td key={c}>{count}</td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

// ---------------------------------------------------------------- how long to close

export interface SpeedRow {
  label: string;
  count: number;
}

export function SpeedBars(props: { rows: SpeedRow[]; pick: Picker }): React.ReactElement {
  const { rows, pick } = props;
  const top = rows.reduce((m, r) => Math.max(m, r.count), 0);
  return (
    <div className={styles.speedBlock}>
      <div className={`${styles.speedBars} ${pickingClass(pick)}`}>
        {rows.map((r, i) => {
          const bar = (
            <>
              <span className={styles.speedCount}>{r.count}</span>
              <span
                className={styles.speedBar}
                style={{
                  height: r.count > 0 ? `${Math.max(2, Math.round((r.count / top) * 120))}px` : '0px',
                  background: i < 3 ? PALETTE.raised : PALETTE.calm
                }}
              />
            </>
          );
          return r.count > 0 ? (
            <Mark key={r.label} pick={pick} drill={{ kind: 'speed', index: i }} label={`${r.label}: ${ticketsText(r.count)}`} className={styles.speedCol}>
              {bar}
            </Mark>
          ) : (
            <div key={r.label} className={styles.speedCol}>
              {bar}
            </div>
          );
        })}
      </div>
      <div className={styles.speedLabels}>
        {rows.map((r) => (
          <span key={r.label} className={styles.speedLabel}>
            {r.label}
          </span>
        ))}
      </div>
    </div>
  );
}

export function SpeedTable(props: { rows: SpeedRow[] }): React.ReactElement {
  return (
    <table className={styles.table}>
      <thead>
        <tr>
          <th scope="col">Time to close</th>
          <th scope="col">Tickets</th>
        </tr>
      </thead>
      <tbody>
        {props.rows.map((r) => (
          <tr key={r.label}>
            <th scope="row">{r.label}</th>
            <td>{r.count}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

// ---------------------------------------------------------------- how old the open ones are

export const AGE_COLOURS = [PALETTE.calm, PALETTE.calmer, PALETTE.deep, PALETTE.alert, PALETTE.critical];

export interface AgeRow {
  label: string;
  count: number;
}

export function AgeBars(props: { rows: AgeRow[]; pick: Picker }): React.ReactElement {
  const { rows, pick } = props;
  const top = rows.reduce((m, r) => Math.max(m, r.count), 0);
  return (
    <div className={`${styles.rows} ${pickingClass(pick)}`}>
      {rows.map((r, i) => {
        const inner = (
          <>
            <span className={styles.rowName}>{r.label}</span>
            <span className={styles.ageTrack}>
              <span className={styles.ageFill} style={{ width: `${barPct(r.count, top)}%`, background: AGE_COLOURS[i] }} />
            </span>
            <span className={styles.rowNum}>{r.count}</span>
          </>
        );
        return r.count > 0 ? (
          <Mark key={r.label} pick={pick} drill={{ kind: 'age', index: i }} label={`Open ${r.label.toLowerCase()}: ${ticketsText(r.count)}`} className={`${styles.ageRow} ${styles.markRow}`}>
            {inner}
          </Mark>
        ) : (
          <div key={r.label} className={styles.ageRow}>
            {inner}
          </div>
        );
      })}
    </div>
  );
}

export function AgeTable(props: { rows: AgeRow[] }): React.ReactElement {
  return (
    <table className={styles.table}>
      <thead>
        <tr>
          <th scope="col">Open for</th>
          <th scope="col">Open tickets</th>
        </tr>
      </thead>
      <tbody>
        {props.rows.map((r) => (
          <tr key={r.label}>
            <th scope="row">{r.label}</th>
            <td>{r.count}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

// ---------------------------------------------------------------- record gaps

export function GapBars(props: { gaps: Gap[]; pick: Picker }): React.ReactElement {
  const { pick } = props;
  return (
    <div className={`${styles.gapList} ${pickingClass(pick)}`}>
      {props.gaps.map((g, i) => (
        <div key={g.label} className={styles.gap}>
          <div className={styles.gapHead}>
            <span>{g.label}</span>
            <strong>{percentText(g.have, g.of)}</strong>
          </div>
          <div className={styles.gapTrack}>
            <div className={styles.gapFill} style={{ width: `${g.of === 0 ? 0 : barPct(g.have, g.of)}%` }} />
          </div>
          <span className={styles.gapNote}>{g.note}</span>
          {g.of - g.have > 0 && (
            <Mark
              pick={pick}
              drill={{ kind: 'gap', index: i }}
              label={`${g.missing}: ${ticketsText(g.of - g.have)}`}
              className={styles.gapButton}
            >
              {`See the ${ticketsText(g.of - g.have)}: ${g.missing.toLowerCase()}`}
            </Mark>
          )}
        </div>
      ))}
    </div>
  );
}

export function GapTable(props: { gaps: Gap[] }): React.ReactElement {
  return (
    <table className={styles.table}>
      <thead>
        <tr>
          <th scope="col">Record</th>
          <th scope="col">Filled in</th>
          <th scope="col">Out of</th>
          <th scope="col">Share</th>
        </tr>
      </thead>
      <tbody>
        {props.gaps.map((g) => (
          <tr key={g.label}>
            <th scope="row">{g.label}</th>
            <td>{g.have}</td>
            <td>{g.of}</td>
            <td>{percentText(g.have, g.of)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}
