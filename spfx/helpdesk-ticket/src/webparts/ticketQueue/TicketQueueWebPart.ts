import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import TicketQueue from './components/TicketQueue';
import { ITicketQueueProps } from './components/ITicketQueueProps';
import { QueueApi } from './sharepoint/queueApi';

export default class TicketQueueWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const { user, legacyPageContext, web } = this.context.pageContext;
    // "Owner" = Full Control on this site, the same test the IT portal uses for
    // its admin buttons: only Full Control carries Manage Permissions.
    // This HIDES the queue from everyone else; what it shows is still guarded by
    // the Tickets list's own permissions, which are what actually protect it.
    const isOwner = web.permissions.hasPermission(SPPermission.managePermissions);
    const element: React.ReactElement<ITicketQueueProps> = React.createElement(TicketQueue, {
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
