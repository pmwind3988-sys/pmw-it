// Everything the web part asks of SharePoint. It runs on the helpdesk page as
// the person looking at it, through SPHttpClient, so there is no sign-in of its
// own and no app permission to grant: it can do exactly what the old list form
// could, and nothing else.

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Person } from '../tickets/draft';
import { fromListItem, myTicketsQuery, TicketRow } from '../tickets/myTickets';

export const LIST_TITLE = 'Tickets';

const JSON_HEADERS = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata',
  'odata-version': ''
};

export class TicketsApi {
  constructor(private readonly http: SPHttpClient, private readonly webUrl: string) {}

  private list(path = ''): string {
    return `${this.webUrl}/_api/web/lists/getbytitle('${LIST_TITLE}')${path}`;
  }

  private async ok(res: SPHttpClientResponse, what: string): Promise<SPHttpClientResponse> {
    if (res.ok) return res;
    let detail = '';
    try {
      const body = await res.json();
      detail = body?.['odata.error']?.message?.value ?? body?.error?.message ?? '';
    } catch {
      /* not JSON */
    }
    throw new Error(`${what} failed (${res.status})${detail ? `: ${detail}` : ''}`);
  }

  /** The person's department as their Microsoft 365 profile has it, or ''. */
  async myDepartment(): Promise<string> {
    try {
      const res = await this.http.get(
        `${this.webUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=UserProfileProperties`,
        SPHttpClient.configurations.v1,
        { headers: JSON_HEADERS }
      );
      if (!res.ok) return '';
      const body = await res.json();
      const props: { Key: string; Value: string }[] = body.UserProfileProperties ?? [];
      const dep = props.find((p) => p.Key === 'Department') ?? props.find((p) => p.Key === 'SPS-Department');
      return dep?.Value ?? '';
    } catch {
      // A missing profile must not stop a ticket; the field just stays blank.
      return '';
    }
  }

  /** Directory people matching what was typed, for "Copy someone in". */
  async searchPeople(query: string): Promise<Person[]> {
    if (query.trim().length < 2) return [];
    const res = await this.ok(
      await this.http.post(
        `${this.webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
        SPHttpClient.configurations.v1,
        {
          headers: JSON_HEADERS,
          body: JSON.stringify({
            queryParams: {
              QueryString: query.trim(),
              MaximumEntitySuggestions: 8,
              AllowEmailAddresses: false,
              AllowMultipleEntities: false,
              PrincipalSource: 15,
              PrincipalType: 1
            }
          })
        }
      ),
      'People search'
    );
    const body = await res.json();
    // The picker answers with a JSON STRING inside its JSON.
    const raw: { Key: string; DisplayText: string; EntityData?: { Email?: string } }[] = JSON.parse(body.value ?? '[]');
    return raw
      .map((e) => ({ name: e.DisplayText, email: e.EntityData?.Email ?? '', key: e.Key }))
      .filter((p) => p.email)
      .map(({ name, email }) => ({ name, email }));
  }

  /** The site user id a User column stores, adding the person to the site's user list if needed. */
  async ensureUser(email: string): Promise<number> {
    const res = await this.ok(
      await this.http.post(`${this.webUrl}/_api/web/ensureuser`, SPHttpClient.configurations.v1, {
        headers: JSON_HEADERS,
        body: JSON.stringify({ logonName: `i:0#.f|membership|${email}` })
      }),
      `Finding ${email}`
    );
    return (await res.json()).Id;
  }

  async createTicket(item: Record<string, unknown>): Promise<number> {
    const res = await this.ok(
      await this.http.post(this.list('/items'), SPHttpClient.configurations.v1, {
        headers: JSON_HEADERS,
        body: JSON.stringify(item)
      }),
      'Saving the ticket'
    );
    return (await res.json()).Id;
  }

  /** One file onto a saved ticket's "Associated files". */
  async addAttachment(itemId: number, file: File): Promise<void> {
    const name = encodeURIComponent(file.name.replace(/'/g, "''"));
    await this.ok(
      await this.http.post(
        this.list(`/items(${itemId})/AttachmentFiles/add(FileName='${name}')`),
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' }, body: file }
      ),
      `Attaching ${file.name}`
    );
  }

  /** The flow's ticket number once it has written it, or '' if it has not yet. */
  async ticketNumber(itemId: number): Promise<string> {
    const res = await this.http.get(
      this.list(`/items(${itemId})?$select=Ticket_x0020_ID`),
      SPHttpClient.configurations.v1,
      { headers: JSON_HEADERS }
    );
    if (!res.ok) return '';
    return (await res.json()).Ticket_x0020_ID ?? '';
  }

  async myTickets(userId: number): Promise<TicketRow[]> {
    const res = await this.ok(
      await this.http.get(this.list(`/items?${myTicketsQuery(userId)}`), SPHttpClient.configurations.v1, {
        headers: JSON_HEADERS
      }),
      'Loading your tickets'
    );
    const body = await res.json();
    return (body.value ?? []).map(fromListItem);
  }

  /** Where the old SharePoint form shows one ticket; "My tickets" links each number there. */
  displayFormUrl(itemId: number): string {
    return `${this.webUrl}/Lists/${LIST_TITLE}/DispForm.aspx?ID=${itemId}`;
  }
}
