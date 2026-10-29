import { QueueApi } from '../../ticketQueue/sharepoint/queueApi';
import { PersonRef } from '../../ticketQueue/queue/ticket';

export interface ITicketDashboardProps {
  api: QueueApi;
  me: PersonRef;
  /** Full Control on the site. Anyone else gets nothing rendered at all. */
  isOwner: boolean;
  /** The page is being edited: a non-owner editor sees a note instead of a blank. */
  isEditing: boolean;
}
