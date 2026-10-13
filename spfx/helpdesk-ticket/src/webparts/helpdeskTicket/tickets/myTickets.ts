// Reading the person's own tickets back for "My tickets".

export type Status = 'New' | 'In progress' | 'On Hold' | 'Completed' | 'Cancelled';

export interface TicketRow {
  id: number;
  /** The flow's number (yymmdd + domain + sequence). Blank for a moment after a submit. */
  number: string;
  title: string;
  category: string;
  priority: string;
  status: string;
  created: string;
  modified: string;
}

const CLOSED = ['Completed', 'Cancelled'];

export function isClosed(status: string): boolean {
  return CLOSED.indexOf(status) !== -1;
}

/** One list item, as `items?$select=` returns it with odata=nometadata. */
export function fromListItem(raw: Record<string, unknown>): TicketRow {
  const s = (v: unknown): string => (typeof v === 'string' ? v : '');
  return {
    id: Number(raw.Id ?? raw.ID),
    number: s(raw.Ticket_x0020_ID),
    title: s(raw.Title),
    category: s(raw.Category),
    priority: s(raw.Priority),
    // A blank status reads as New, which is the column's default.
    status: s(raw.Status) || 'New',
    created: s(raw.Created),
    modified: s(raw.Modified)
  };
}

export const SELECT = 'Id,Title,Ticket_x0020_ID,Category,Priority,Status,Created,Modified';

/**
 * Only what this person raised. The list already limits ordinary users to
 * their own rows, but IT staff can read every row, so the filter is explicit
 * or an IT member's "My tickets" would be the whole helpdesk.
 */
export function myTicketsQuery(userId: number, top = 200): string {
  return `$select=${SELECT}&$filter=AuthorId eq ${userId}&$orderby=Created desc&$top=${top}`;
}

export function matchesSearch(t: TicketRow, q: string): boolean {
  const needle = q.trim().toLowerCase();
  if (!needle) return true;
  return [t.number, t.title, t.category].some((v) => v.toLowerCase().indexOf(needle) !== -1);
}

export function splitOpenClosed(rows: TicketRow[]): { open: TicketRow[]; closed: TicketRow[] } {
  return { open: rows.filter((r) => !isClosed(r.status)), closed: rows.filter((r) => isClosed(r.status)) };
}
