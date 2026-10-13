import { TicketsApi } from '../sharepoint/ticketsApi';

export type View = 'raise' | 'mine';

export interface SignedInUser {
  name: string;
  email: string;
  /** Site user id; "My tickets" filters Created By on it. */
  id: number;
}

export interface IHelpdeskTicketProps {
  api: TicketsApi;
  user: SignedInUser;
  startView: View;
}
