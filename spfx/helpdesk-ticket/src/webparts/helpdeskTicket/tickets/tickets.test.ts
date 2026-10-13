import { AREAS, categoryValue, effectiveKind, DEFAULT_PRIORITY, PRIORITIES } from './categories';
import { emptyDraft, missingOf, toTicketItem, detailsToHtml, Draft } from './draft';
import { fromListItem, isClosed, myTicketsQuery, matchesSearch, splitOpenClosed } from './myTickets';

// The Tickets list's Category choices, read from the tenant on 2026-10-13.
const COLUMN_CHOICES = [
  'SAP - Incident', 'SAP - Request', 'SAP - Change',
  'Hardware - Incident', 'Hardware - Request',
  'Application/Software - Incident', 'Application/Software - Request',
  'Network - Incident',
  'Access - Incident', 'Access - Request',
  'Security - Incident', 'Security - Request',
  'Email - Incident', 'Email - Request',
  'How-To/Training - Request',
  'Onboarding/Offboarding - Request'
];

describe('categories', () => {
  it('offers exactly the column choices, every one of them', () => {
    const offered: string[] = [];
    AREAS.forEach((a) => a.kinds.forEach((k) => offered.push(`${a.id} - ${k}`)));
    expect([...offered].sort()).toEqual([...COLUMN_CHOICES].sort());
  });

  it('needs no second answer when an area has one kind', () => {
    expect(effectiveKind('Network', '')).toBe('Incident');
    expect(categoryValue('How-To/Training', '')).toBe('How-To/Training - Request');
  });

  it('waits for the second answer when an area has several', () => {
    expect(categoryValue('SAP', '')).toBe('');
    expect(categoryValue('SAP', 'Change')).toBe('SAP - Change');
  });

  it('drops a kind the new area does not have', () => {
    expect(categoryValue('Hardware', 'Change')).toBe('');
  });

  it('starts priority at Normal, the column default, and never offers Undefined', () => {
    expect(DEFAULT_PRIORITY).toBe('Normal');
    expect(PRIORITIES.map((p) => p.id)).toEqual(['Low', 'Normal', 'High', 'Critical']);
  });
});

describe('missingOf', () => {
  it('lists everything on an empty draft', () => {
    expect(missingOf(emptyDraft())).toEqual(["What it's about", 'Short description']);
  });

  it('asks for the kind only once an area with several is picked', () => {
    const d: Draft = { ...emptyDraft(), area: 'SAP', title: 'x' };
    expect(missingOf(d)).toEqual(['Broken, or need something']);
    expect(missingOf({ ...d, kind: 'Request' })).toEqual([]);
  });

  it('does not count a description of spaces', () => {
    expect(missingOf({ ...emptyDraft(), area: 'Network', title: '   ' })).toEqual(['Short description']);
  });
});

describe('detailsToHtml', () => {
  it('escapes markup so a typed "<" stays text', () => {
    expect(detailsToHtml('a <b> & "c"')).toBe('<p>a &lt;b&gt; &amp; &quot;c&quot;</p>');
  });

  it('keeps line breaks and paragraphs', () => {
    expect(detailsToHtml('one\ntwo\r\n\r\nthree')).toBe('<p>one<br>two</p><p>three</p>');
  });

  it('writes nothing for blank text', () => {
    expect(detailsToHtml('  \n ')).toBe('');
  });
});

describe('toTicketItem', () => {
  const d: Draft = {
    area: 'Email', kind: 'Incident', priority: 'High',
    title: '  Outlook asks for my password  ', details: 'Since Monday',
    cc: [{ id: 12, email: 'a@pmw', name: 'A' }, { email: 'b@pmw', name: 'B' }]
  };

  it('writes the answered columns and the department', () => {
    expect(toTicketItem(d, 'Finance')).toEqual({
      Title: 'Outlook asks for my password',
      Category: 'Email - Incident',
      Priority: 'High',
      Description: '<p>Since Monday</p>',
      Department: 'Finance',
      CcId: [12]
    });
  });

  it('never writes what the flow or SharePoint owns', () => {
    const item = toTicketItem(d, 'Finance');
    for (const k of ['Ticket_x0020_ID', 'Status', 'Assignedto0Id', 'AuthorId']) expect(item).not.toHaveProperty(k);
  });

  it('leaves out empty optional columns', () => {
    const item = toTicketItem({ ...d, details: '', cc: [] }, ' ');
    expect(item).not.toHaveProperty('Description');
    expect(item).not.toHaveProperty('Department');
    expect(item).not.toHaveProperty('CcId');
  });

  it('cuts the title to the column length', () => {
    expect((toTicketItem({ ...d, title: 'x'.repeat(300) }, '').Title as string).length).toBe(255);
  });
});

describe('my tickets', () => {
  it('reads a list item, a blank status as New', () => {
    expect(fromListItem({ Id: 5, Title: 't', Ticket_x0020_ID: '261013PMW001', Status: null }).status).toBe('New');
    expect(fromListItem({ Id: 5, Title: 't', Ticket_x0020_ID: '261013PMW001' }).number).toBe('261013PMW001');
  });

  it('counts Completed and Cancelled as closed, On Hold as open', () => {
    expect(isClosed('Completed')).toBe(true);
    expect(isClosed('Cancelled')).toBe(true);
    expect(isClosed('On Hold')).toBe(false);
  });

  it('filters on the author, so IT staff see only their own', () => {
    expect(myTicketsQuery(42)).toContain('$filter=AuthorId eq 42');
  });

  it('searches number, title and category', () => {
    const r = fromListItem({ Id: 1, Title: 'Printer jams', Ticket_x0020_ID: '261009PMW003', Category: 'Hardware - Incident' });
    expect(matchesSearch(r, 'pmw003')).toBe(true);
    expect(matchesSearch(r, 'jams')).toBe(true);
    expect(matchesSearch(r, 'sap')).toBe(false);
  });

  it('splits open from closed', () => {
    const rows = ['New', 'Completed', 'On Hold'].map((s, i) => fromListItem({ Id: i, Status: s }));
    const { open, closed } = splitOpenClosed(rows);
    expect(open.length).toBe(2);
    expect(closed.length).toBe(1);
  });
});
