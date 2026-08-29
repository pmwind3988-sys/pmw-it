/**
 * One page of a long list.
 *
 * A register of two thousand rows rendered in full is two thousand table rows
 * in the document, each with its own links, chips and per-item arithmetic. The
 * browser lays every one of them out before the first is visible, which on a
 * phone is the difference between a list that opens and a list that appears to
 * have hung. Only the page being looked at is built.
 *
 * Pure, and deliberately forgiving about the page number: filters change under
 * a reader's feet, and asking for page 9 of a list that just became one page
 * long has to answer with the rows that exist rather than with nothing.
 */

/** Every option the size picker offers. `0` means "all of them". */
export const PAGE_SIZES = [25, 50, 100, 0];

export function pageCount(total, size) {
  if (!size || size <= 0) return 1;
  return Math.max(1, Math.ceil(total / size));
}

export function paginate(items = [], page = 1, size = 25) {
  const total = items.length;
  const pages = pageCount(total, size);
  // Clamped rather than trusted: a page number can outlive the list it counts.
  const current = Math.min(Math.max(1, Math.trunc(Number(page) || 1)), pages);

  if (!size || size <= 0) {
    return { rows: items, page: 1, pages: 1, total, from: total ? 1 : 0, to: total };
  }

  const start = (current - 1) * size;
  const rows = items.slice(start, start + size);

  return {
    rows,
    page: current,
    pages,
    total,
    // Said in the reader's numbering, from one: "showing 26-50 of 312".
    from: total ? start + 1 : 0,
    to: start + rows.length,
  };
}
