// What a person picks, and the Tickets list value it files as.
//
// The list's Category column holds "<area> - <kind>" (e.g. "SAP - Request"),
// sixteen values in all. Asking for one of sixteen at once is why most tickets
// arrived with no category, so the form asks two easy questions instead and
// builds the value. Every pair here must be one of the column's choices --
// SharePoint refuses a choice it has never heard of -- and categories.test.ts
// pins the list against the column as read on 2026-10-13.

export type Kind = 'Incident' | 'Request' | 'Change';

export interface Area {
  /** The prefix the column stores. */
  id: string;
  /** What the person reads. */
  label: string;
  /** The kinds the column has for this area, in the order offered. */
  kinds: Kind[];
}

export const AREAS: Area[] = [
  { id: 'Hardware', label: 'Laptop, PC or printer', kinds: ['Incident', 'Request'] },
  { id: 'SAP', label: 'SAP', kinds: ['Incident', 'Request', 'Change'] },
  { id: 'Application/Software', label: 'Software & apps', kinds: ['Incident', 'Request'] },
  { id: 'Email', label: 'Email & Teams', kinds: ['Incident', 'Request'] },
  { id: 'Network', label: 'Internet & Wi-Fi', kinds: ['Incident'] },
  { id: 'Access', label: 'Access & passwords', kinds: ['Incident', 'Request'] },
  { id: 'Security', label: 'Security or a scam email', kinds: ['Incident', 'Request'] },
  { id: 'How-To/Training', label: 'How do I…?', kinds: ['Request'] },
  { id: 'Onboarding/Offboarding', label: 'Someone joining or leaving', kinds: ['Request'] }
];

export const KIND_LABELS: Record<Kind, string> = {
  Incident: "Something's broken",
  Request: 'I need something',
  Change: 'Change something in SAP'
};

export function areaById(id: string): Area | undefined {
  return AREAS.find((a) => a.id === id);
}

/**
 * The kind that stands for this area: the one picked if the area offers it,
 * the only one if there is only one, otherwise none yet.
 */
export function effectiveKind(areaId: string, picked: Kind | ''): Kind | '' {
  const area = areaById(areaId);
  if (!area) return '';
  if (area.kinds.length === 1) return area.kinds[0];
  return picked && area.kinds.indexOf(picked) !== -1 ? picked : '';
}

/** The Category value to store, or '' while the two answers are incomplete. */
export function categoryValue(areaId: string, picked: Kind | ''): string {
  const kind = effectiveKind(areaId, picked);
  return kind ? `${areaId} - ${kind}` : '';
}

// Priority: the person sets it, IT may change it afterwards in the list.
// The column's default is Normal, so the form starts there. "Undefined" is a
// column choice nobody should pick and is not offered.
export type Priority = 'Low' | 'Normal' | 'High' | 'Critical';

export const PRIORITIES: { id: Priority; label: string }[] = [
  { id: 'Low', label: 'It can wait' },
  { id: 'Normal', label: 'Slowing me down' },
  { id: 'High', label: "I can't work" },
  { id: 'Critical', label: "Several of us can't work" }
];

export const DEFAULT_PRIORITY: Priority = 'Normal';
