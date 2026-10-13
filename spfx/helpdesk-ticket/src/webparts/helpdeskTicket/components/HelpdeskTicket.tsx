// Root of the web part: the two screens, the address-bar view and the breadcrumb.
import * as React from 'react';
import { IHelpdeskTicketProps, View } from './IHelpdeskTicketProps';
import { MyTickets } from './MyTickets';
import { RaiseTicket } from './RaiseTicket';
import { useMountedRef } from './hooks';
import styles from './HelpdeskTicket.module.scss';

export default function HelpdeskTicket(props: IHelpdeskTicketProps): React.ReactElement {
  const { api, user, startView } = props;
  const [view, setViewState] = React.useState<View>(startView);
  const [department, setDepartment] = React.useState('');
  const mounted = useMountedRef();

  React.useEffect(() => {
    api.myDepartment().then((d) => {
      if (mounted.current) setDepartment(d);
    }).catch(() => undefined);
  }, [api, mounted]);

  // The form keeps its answers when the person looks at "My tickets" and back,
  // so both screens stay mounted and only one is shown.
  const setView = React.useCallback((next: View): void => {
    setViewState(next);
    try {
      const url = new URL(window.location.href);
      url.searchParams.set('view', next);
      window.history.replaceState(window.history.state, '', url.toString());
    } catch {
      // The address bar only mirrors the screen; a refusal changes nothing.
    }
  }, []);

  return (
    <div className={styles.root}>
      <nav aria-label="Breadcrumb" className={styles.crumbs}>
        <span>IT helpdesk</span>
        <span aria-hidden="true">›</span>
        <span aria-current="page">{view === 'raise' ? 'Raise a ticket' : 'My tickets'}</span>
      </nav>
      <div hidden={view !== 'raise'}>
        <RaiseTicket api={api} user={user} department={department} onShowMine={() => setView('mine')} />
      </div>
      <div hidden={view !== 'mine'}>
        <MyTickets api={api} userId={user.id} active={view === 'mine'} onRaise={() => setView('raise')} />
      </div>
    </div>
  );
}
