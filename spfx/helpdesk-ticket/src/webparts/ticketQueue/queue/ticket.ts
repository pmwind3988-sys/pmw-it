// One Tickets list row as the queue reads it, and the columns it can export.

export interface PersonRef {
  id: number;
  name: string;
  email: string;
}

export interface QueueTicket {
  id: number;
  number: string;
  title: string;
  description: string;
  category: string;
  priority: string;
  status: string;
  assignedTo?: PersonRef;
  raisedBy?: PersonRef;
  department: string;
  cc: PersonRef[];
  created: string;
  modified: string;
  completedDate: string;
  completedBy?: PersonRef;
  rootCause: string;
  resolution: string;
  comments: string;
  holdCount: number | undefined;
  totalHoldDays: number | undefined;
  kbCandidate: boolean;
}

/** The columns read, and the people columns expanded, in one request. */
export const QUEUE_SELECT = [
  'Id', 'Title', 'Ticket_x0020_ID', 'Description', 'Category', 'Priority', 'Status', 'Department',
  'Created', 'Modified', 'Completed_x0020_Date', 'Root_x0020_Cause', 'Resolution_x0020_Summary',
  'Comments', 'HoldCount', 'TotalHoldDays', 'KB_x0020_Candidate',
  'Assignedto0/Id', 'Assignedto0/Title', 'Assignedto0/EMail',
  'Author/Id', 'Author/Title', 'Author/EMail',
  'Cc/Id', 'Cc/Title', 'Cc/EMail',
  'CompletedBy/Id', 'CompletedBy/Title', 'CompletedBy/EMail'
].join(',');

export const QUEUE_EXPAND = 'Assignedto0,Author,Cc,CompletedBy';

type Raw = Record<string, unknown>;

const str = (v: unknown): string => (typeof v === 'string' ? v : '');
const num = (v: unknown): number | undefined =>
  typeof v === 'number' ? v : v === undefined || v === null || v === '' ? undefined : Number(v);

function person(v: unknown): PersonRef | undefined {
  const p = v as Raw | null | undefined;
  if (!p || typeof p.Id !== 'number') return undefined;
  return { id: p.Id, name: str(p.Title), email: str(p.EMail) };
}

/**
 * Rich-text columns (Issue description, Comments, Resolution) arrive as HTML.
 * A spreadsheet cell and a table cell want the words: keep line breaks, drop
 * the markup, decode the entities SharePoint writes.
 */
export function htmlToText(html: string): string {
  if (!html) return '';
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li|h\d)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;|&#160;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&#x27;/g, "'")
    .replace(/&#58;/g, ':')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .replace(/&amp;/g, '&')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

export function fromQueueItem(raw: Raw): QueueTicket {
  const cc = (raw.Cc as Raw[] | { results?: Raw[] } | null) ?? [];
  const ccList = Array.isArray(cc) ? cc : cc.results ?? [];
  return {
    id: Number(raw.Id ?? raw.ID),
    number: str(raw.Ticket_x0020_ID),
    title: str(raw.Title),
    description: htmlToText(str(raw.Description)),
    category: str(raw.Category),
    priority: str(raw.Priority),
    status: str(raw.Status) || 'New',
    assignedTo: person(raw.Assignedto0),
    raisedBy: person(raw.Author),
    department: str(raw.Department),
    cc: ccList.map(person).filter((p): p is PersonRef => !!p),
    created: str(raw.Created),
    modified: str(raw.Modified),
    completedDate: str(raw.Completed_x0020_Date),
    completedBy: person(raw.CompletedBy),
    rootCause: str(raw.Root_x0020_Cause),
    resolution: htmlToText(str(raw.Resolution_x0020_Summary)),
    comments: htmlToText(str(raw.Comments)),
    holdCount: num(raw.HoldCount),
    totalHoldDays: num(raw.TotalHoldDays),
    kbCandidate: raw.KB_x0020_Candidate === true
  };
}
