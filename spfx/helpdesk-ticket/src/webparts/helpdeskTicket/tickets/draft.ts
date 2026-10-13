// A ticket being written, what it still needs, and the row it becomes.

import { categoryValue, effectiveKind, areaById, DEFAULT_PRIORITY, Kind, Priority } from './categories';

export const TITLE_MAX = 255;

export interface Person {
  /** SharePoint's site user id, what a User column stores. */
  id?: number;
  email: string;
  name: string;
}

export interface Draft {
  area: string;
  kind: Kind | '';
  priority: Priority;
  title: string;
  details: string;
  cc: Person[];
}

export function emptyDraft(): Draft {
  return { area: '', kind: '', priority: DEFAULT_PRIORITY, title: '', details: '', cc: [] };
}

/** What the person still has to answer, in the order the form asks it. */
export function missingOf(d: Draft): string[] {
  const missing: string[] = [];
  if (!areaById(d.area)) missing.push("What it's about");
  else if (!effectiveKind(d.area, d.kind)) missing.push('Broken, or need something');
  if (!d.title.trim()) missing.push('Short description');
  return missing;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * "Issue description" is a RICH-TEXT column, so plain text written straight in
 * loses its line breaks and lets a typed "<" be read as markup. Escape it, and
 * give each paragraph its own <p> and each line break a <br>.
 */
export function detailsToHtml(text: string): string {
  const trimmed = text.replace(/\r\n?/g, '\n').trim();
  if (!trimmed) return '';
  return trimmed
    .split(/\n{2,}/)
    .map((para) => `<p>${escapeHtml(para).replace(/\n/g, '<br>')}</p>`)
    .join('');
}

/**
 * The row a submit writes. Only the columns the person answers -- plus the
 * department, which comes from their profile -- and never Ticket ID, Assigned
 * to or Status: the Power Automate flow fills those when the row arrives, and
 * who raised it is SharePoint's own Created By.
 *
 * Cc people must already carry their site user id (see ensureUser).
 */
export function toTicketItem(d: Draft, department: string): Record<string, unknown> {
  const item: Record<string, unknown> = {
    Title: d.title.trim().slice(0, TITLE_MAX),
    Category: categoryValue(d.area, d.kind),
    Priority: d.priority
  };
  const html = detailsToHtml(d.details);
  if (html) item.Description = html;
  if (department.trim()) item.Department = department.trim().slice(0, 255);
  const ccIds = d.cc.map((p) => p.id).filter((id): id is number => typeof id === 'number');
  if (ccIds.length) item.CcId = ccIds;
  return item;
}
