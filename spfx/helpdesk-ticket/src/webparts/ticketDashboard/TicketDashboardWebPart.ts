import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import TicketDashboard from './components/TicketDashboard';
import { ITicketDashboardProps } from './components/ITicketDashboardProps';
import { QueueApi } from '../ticketQueue/sharepoint/queueApi';

export default class TicketDashboardWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const { user, legacyPageContext, web } = this.context.pageContext;
    // Site owners only, the same test as the ticket queue (Full Control carries
    // Manage Permissions). It hides the dashboard; the Tickets list's own
    // permissions are what protect the data.
    const isOwner = web.permissions.hasPermission(SPPermission.managePermissions);
    const element: React.ReactElement<ITicketDashboardProps> = React.createElement(TicketDashboard, {
      api: new QueueApi(this.context.spHttpClient, web.absoluteUrl),
      me: { id: legacyPageContext?.userId as number, name: user.displayName, email: user.email },
      isOwner,
      isEditing: this.displayMode === DisplayMode.Edit
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
