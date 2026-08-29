import { ChevronLeft, ChevronRight } from './Icons';
import { PAGE_SIZES } from './paginate';

/**
 * The bar under a paged table.
 *
 * It says where in the list you are before it offers to move you, because
 * "26-50 of 312" is the thing somebody actually wants from it; the arrows are
 * the smaller half of the job. The size picker is here rather than in the
 * filters: how much to load at once is a fact about this screen and this
 * connection, not about which rows are wanted.
 */
export default function Pager({ page, onPage, size, onSize, label = 'rows' }) {
  if (page.total === 0) return null;

  return (
    <div className="ui-pager">
      <span className="ui-pager-count">
        {page.from}–{page.to} of {page.total} {label}
      </span>

      <div className="ui-pager-moves">
        <button
          type="button"
          className="ui-pager-step"
          disabled={page.page <= 1}
          onClick={() => onPage(page.page - 1)}
          aria-label="The page before"
        >
          <ChevronLeft size={15} />
        </button>
        <span className="ui-pager-where">Page {page.page} of {page.pages}</span>
        <button
          type="button"
          className="ui-pager-step"
          disabled={page.page >= page.pages}
          onClick={() => onPage(page.page + 1)}
          aria-label="The next page"
        >
          <ChevronRight size={15} />
        </button>
      </div>

      <label className="ui-pager-size">
        <span>Show</span>
        <select value={size} onChange={(event) => onSize(Number(event.target.value))}>
          {PAGE_SIZES.map((option) => (
            <option key={option} value={option}>{option === 0 ? 'All' : option}</option>
          ))}
        </select>
      </label>
    </div>
  );
}
