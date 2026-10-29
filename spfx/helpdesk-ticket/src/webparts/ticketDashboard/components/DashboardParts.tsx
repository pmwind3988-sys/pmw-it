// The pieces every section is built from: a chart card with its "Show as table"
// switch, the three story cards on top, and the key-figure tiles.
import * as React from 'react';
import { useInstanceId } from '../../helpdeskTicket/components/hooks';
import { Story } from '../stats/stats';
import styles from './Dashboard.module.scss';

export interface ChartCardProps {
  title: string;
  subtitle: string;
  /** Layout class: how wide the card sits in its row. */
  className: string;
  /** Nothing in view: the card says so instead of drawing an empty chart. */
  empty: boolean;
  chart: React.ReactNode;
  table: React.ReactNode;
  legend?: React.ReactNode;
  footer?: React.ReactNode;
}

export function ChartCard(props: ChartCardProps): React.ReactElement {
  const { title, subtitle, className, empty, chart, table, legend, footer } = props;
  const [asTable, setAsTable] = React.useState(false);
  const headingId = useInstanceId('dash-card');
  return (
    <section aria-labelledby={headingId} className={`${styles.card} ${styles.section} ${className}`}>
      <div className={styles.cardHead}>
        <div className={styles.cardText}>
          <h2 id={headingId} className={styles.cardTitle}>
            {title}
          </h2>
          <p className={styles.cardSub}>{subtitle}</p>
        </div>
        {!empty && (
          <div className={styles.cardTools}>
            {legend}
            <button
              type="button"
              className={styles.toggle}
              aria-pressed={asTable}
              onClick={() => setAsTable(!asTable)}
            >
              Show as table
            </button>
          </div>
        )}
      </div>

      {empty && <p className={styles.empty}>No tickets match these filters.</p>}
      {!empty && asTable && <div className={styles.tableBox}>{table}</div>}
      {!empty && !asTable && chart}
      {!empty && footer}
    </section>
  );
}

export function LegendItem(props: { colour: string; label: string; square?: boolean }): React.ReactElement {
  const { colour, label, square } = props;
  return (
    <span className={styles.legendItem}>
      <span
        className={square ? styles.swatchSquare : styles.swatchLine}
        style={{ background: colour }}
        aria-hidden="true"
      />
      {label}
    </span>
  );
}

export function StoryCard(props: { story: Story }): React.ReactElement {
  const { story } = props;
  return (
    <div className={styles.story}>
      <div className={styles.storyBadge}>{story.badge}</div>
      <div className={styles.storyText}>
        <span className={styles.storyKicker}>{story.kicker}</span>
        <span className={styles.storyBody}>{story.text}</span>
      </div>
    </div>
  );
}

export interface KpiTileProps {
  label: string;
  value: string;
  note: string;
  noteColour?: string;
  /** When set the whole tile is a link that opens in a new tab. */
  href?: string;
}

export function KpiTile(props: KpiTileProps): React.ReactElement {
  const { label, value, note, noteColour, href } = props;
  const body = (
    <>
      <span className={styles.kpiLabel}>{label}</span>
      <span className={styles.kpiValue}>{value}</span>
      <span className={styles.kpiNote} style={{ color: noteColour }}>
        {note}
      </span>
    </>
  );
  if (href !== undefined) {
    return (
      <a className={`${styles.kpi} ${styles.kpiLink}`} href={href} target="_blank" rel="noopener noreferrer">
        {body}
      </a>
    );
  }
  return <div className={styles.kpi}>{body}</div>;
}
