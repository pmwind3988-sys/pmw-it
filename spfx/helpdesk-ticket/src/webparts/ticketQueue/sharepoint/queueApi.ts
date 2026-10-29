// Reading the whole Tickets list for the queue. It runs as the signed-in
// owner, so it sees what that owner may see in the list and nothing more.

import { SPHttpClient } from '@microsoft/sp-http';
import { fromQueueItem, QUEUE_EXPAND, QUEUE_SELECT, QueueTicket } from '../queue/ticket';

const LIST_TITLE = 'Tickets';
const HEADERS = { Accept: 'application/json;odata=nometadata', 'odata-version': '' };
/** Pages of 2000; a few requests read the whole list today (~650 rows). */
const PAGE = 2000;
/** A stop far past any real list, so a paging bug cannot loop forever. */
const MAX_PAGES = 50;

export class QueueApi {
  constructor(private readonly http: SPHttpClient, private readonly webUrl: string) {}

  /**
   * Every ticket, newest first. Filtering happens in the browser: the date
   * presets move with the clock and the assignee list needs every name, so one
   * read serves every change of filter without asking SharePoint again.
   */
  async allTickets(): Promise<QueueTicket[]> {
    let url: string | undefined =
      `${this.webUrl}/_api/web/lists/getbytitle('${LIST_TITLE}')/items` +
      `?$select=${QUEUE_SELECT}&$expand=${QUEUE_EXPAND}&$orderby=Id desc&$top=${PAGE}`;
    const rows: QueueTicket[] = [];
    for (let i = 0; url && i < MAX_PAGES; i++) {
      const res = await this.http.get(url, SPHttpClient.configurations.v1, { headers: HEADERS });
      if (!res.ok) {
        let detail = '';
        try {
          const body = await res.json();
          detail = body?.['odata.error']?.message?.value ?? '';
        } catch {
          /* not JSON */
        }
        throw new Error(`Loading tickets failed (${res.status})${detail ? `: ${detail}` : ''}`);
      }
      const body = await res.json();
      (body.value ?? []).forEach((raw: Record<string, unknown>) => rows.push(fromQueueItem(raw)));
      url = body['odata.nextLink'];
    }
    return rows;
  }

  displayFormUrl(id: number): string {
    return `${this.webUrl}/Lists/${LIST_TITLE}/DispForm.aspx?ID=${id}`;
  }

  editFormUrl(id: number): string {
    return `${this.webUrl}/Lists/${LIST_TITLE}/EditForm.aspx?ID=${id}`;
  }
}
