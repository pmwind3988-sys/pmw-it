// One tappable thing on a chart: a bar, a segment, a square, a row. It looks
// exactly as it did; tapping it opens the tickets behind it in TicketPreview.
import * as React from 'react';
import { Drill, isPicked, Picker } from '../stats/drill';
import styles from './Dashboard.module.scss';

export interface MarkProps {
  pick: Picker;
  drill: Drill;
  /** Read out, and shown on hover: what it is and how many. */
  label: string;
  className: string;
  style?: React.CSSProperties;
  children?: React.ReactNode;
}

export function Mark(props: MarkProps): React.ReactElement {
  const { pick, drill, label, className, style, children } = props;
  return (
    <button
      type="button"
      className={`${className} ${styles.mark}`}
      style={style}
      aria-label={children === undefined ? label : undefined}
      title={label}
      aria-pressed={isPicked(pick, drill)}
      onClick={() => pick.onPick(drill)}
    >
      {children}
    </button>
  );
}

/**
 * A bar segment: a Mark when it has tickets, a plain span when it has none
 * (nothing to list, and nothing to tab to).
 */
export function Segment(props: MarkProps & { count: number }): React.ReactElement {
  const { count, className, style, label } = props;
  if (count === 0) return <span className={className} style={style} title={label} aria-hidden="true" />;
  return <Mark {...props} />;
}
