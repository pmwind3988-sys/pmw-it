// The filter bar above the table: who, when, which statuses, and what text.
import * as React from 'react';
import { DATE_PRESETS, DateChoice, DatePreset, describe as describeRange } from '../queue/dateRange';
import { QueueFilter, STATUSES } from '../queue/queueFilter';
import { PersonRef, QueueTicket } from '../queue/ticket';
import { useInstanceId } from '../../helpdeskTicket/components/hooks';
import { AssigneePicker } from './AssigneePicker';
import { dayKey, monthKey, OPEN_STATUSES, sameSet } from './queueFormat';
import styles from './TicketQueue.module.scss';

export interface QueueFiltersProps {
  filter: QueueFilter;
  me: PersonRef;
  rows: QueueTicket[];
  /** The assignee button's words, worked out by the screen. */
  assigneeText: string;
  searchText: string;
  onSearchText(text: string): void;
  onChange(change: Partial<QueueFilter>): void;
  onReset(): void;
}

export function QueueFilters(props: QueueFiltersProps): React.ReactElement {
  const { filter, me, rows, assigneeText, searchText, onSearchText, onChange, onReset } = props;
  const ids = {
    assign: useInstanceId('qf-assign'),
    raised: useInstanceId('qf-raised'),
    month: useInstanceId('qf-month'),
    from: useInstanceId('qf-from'),
    to: useInstanceId('qf-to'),
    search: useInstanceId('qf-search')
  };
  const now = new Date();
  const date = filter.date;

  // Picking a preset fills in the one value it needs, from today, so the inputs never start empty.
  const choose = (preset: DatePreset): void => {
    let next: DateChoice = { preset };
    if (preset === 'month') next = { preset, month: date.month ?? monthKey(now) };
    if (preset === 'range') {
      next = {
        preset,
        start: date.start ?? dayKey(new Date(now.getFullYear(), now.getMonth(), 1)),
        end: date.end ?? dayKey(now)
      };
    }
    onChange({ date: next });
  };

  const selected = filter.statuses.length ? filter.statuses : STATUSES;
  const toggleStatus = (status: string): void => {
    const next = selected.indexOf(status) !== -1 ? selected.filter((s) => s !== status) : [...selected, status];
    // At least one status stays ticked; clearing the last one would only mean "all".
    if (!next.length) return;
    onChange({ statuses: STATUSES.filter((s) => next.indexOf(s) !== -1) });
  };

  return (
    <div className={styles.filters}>
      <div className={styles.field}>
        <span id={ids.assign} className={styles.fieldLabel}>
          Assigned to
        </span>
        <AssigneePicker
          ids={filter.assignees}
          me={me}
          rows={rows}
          summary={assigneeText}
          onChange={(list) => onChange({ assignees: list })}
        />
      </div>

      <div className={styles.field}>
        <label htmlFor={ids.raised} className={styles.fieldLabel}>
          Raised
        </label>
        <select
          id={ids.raised}
          className={styles.control}
          value={date.preset}
          onChange={(e) => choose(e.target.value as DatePreset)}
        >
          {DATE_PRESETS.map((p) => (
            <option key={p.id} value={p.id}>
              {p.label}
            </option>
          ))}
        </select>

        {date.preset === 'month' && (
          <>
            <label htmlFor={ids.month} className={styles.visuallyHidden}>
              Month
            </label>
            <input
              id={ids.month}
              type="month"
              className={styles.control}
              value={date.month ?? ''}
              onChange={(e) => onChange({ date: { ...date, month: e.target.value } })}
            />
          </>
        )}

        {date.preset === 'range' && (
          <div className={styles.datePair}>
            <label className={styles.inlineLabel} htmlFor={ids.from}>
              From
              <input
                id={ids.from}
                type="date"
                className={styles.control}
                value={date.start ?? ''}
                onChange={(e) => onChange({ date: { ...date, start: e.target.value } })}
              />
            </label>
            <label className={styles.inlineLabel} htmlFor={ids.to}>
              To
              <input
                id={ids.to}
                type="date"
                className={styles.control}
                value={date.end ?? ''}
                onChange={(e) => onChange({ date: { ...date, end: e.target.value } })}
              />
            </label>
          </div>
        )}

        <p className={styles.hint}>{describeRange(date, now)}</p>
      </div>

      <div className={styles.fieldWide} role="group" aria-label="Status">
        <span className={styles.fieldLabel}>Status</span>
        <div className={styles.chips}>
          <button
            type="button"
            className={styles.chip}
            aria-pressed={sameSet(filter.statuses, OPEN_STATUSES)}
            onClick={() => onChange({ statuses: OPEN_STATUSES.slice() })}
          >
            Open
          </button>
          <button
            type="button"
            className={styles.chip}
            aria-pressed={filter.statuses.length === 0}
            onClick={() => onChange({ statuses: [] })}
          >
            All
          </button>
          {STATUSES.map((s) => (
            <button
              key={s}
              type="button"
              className={styles.chip}
              aria-pressed={selected.indexOf(s) !== -1}
              onClick={() => toggleStatus(s)}
            >
              {s}
            </button>
          ))}
        </div>
      </div>

      <div className={styles.fieldWide}>
        <label htmlFor={ids.search} className={styles.visuallyHidden}>
          Search tickets
        </label>
        <input
          id={ids.search}
          type="search"
          className={styles.searchInput}
          placeholder="Search number, description, person…"
          value={searchText}
          onChange={(e) => onSearchText(e.target.value)}
        />
        <button type="button" className={styles.linkBtn} onClick={onReset}>
          Reset filters
        </button>
      </div>
    </div>
  );
}
