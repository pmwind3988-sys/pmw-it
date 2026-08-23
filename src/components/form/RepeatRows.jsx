import { Plus, Trash2 } from '../ui/Icons';

/**
 * A list of rows somebody can add to and remove from.
 *
 * There is no ceiling. The reference form this replaces offered exactly three
 * item slots, which is a limit of the tool it was built in rather than a rule —
 * somebody needing a fourth thing should not have to submit twice.
 *
 * `renderRow(row, index, setRow)` draws one row's controls. Removing is refused
 * when only one row is left rather than hidden, because a list that can empty
 * itself leaves nothing to type into and no obvious way back.
 */
export default function RepeatRows({
  rows, onChange, renderRow, newRow, addLabel = 'Add another', title,
}) {
  const setRow = (index) => (next) => onChange(
    rows.map((row, position) => (position === index ? next : row)),
  );

  const remove = (index) => onChange(rows.filter((unused, position) => position !== index));

  return (
    <div className="ff-rows">
      {title && <p className="ff-rows-title">{title}</p>}

      {rows.map((row, index) => (
        // Rows have no id of their own and reorder only by removal, so the
        // index is a stable enough key here — and the alternative is putting an
        // id on data that goes on to be serialised.
        <div className="ff-row" key={index}>
          <span className="ff-row-index">{index + 1}</span>
          <div className="ff-row-body">{renderRow(row, index, setRow(index))}</div>
          <button
            type="button"
            className="ff-row-remove"
            onClick={() => remove(index)}
            disabled={rows.length <= 1}
            aria-label={`Remove row ${index + 1}`}
          >
            <Trash2 size={14} />
          </button>
        </div>
      ))}

      <button
        type="button"
        className="ff-add"
        onClick={() => onChange([...rows, newRow()])}
      >
        <Plus size={14} /> {addLabel}
      </button>
    </div>
  );
}
