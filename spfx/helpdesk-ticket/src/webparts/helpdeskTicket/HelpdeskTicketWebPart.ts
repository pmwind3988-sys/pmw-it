import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneChoiceGroup } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HelpdeskTicketWebPartStrings';
import HelpdeskTicket from './components/HelpdeskTicket';
import { IHelpdeskTicketProps, View } from './components/IHelpdeskTicketProps';
import { TicketsApi } from './sharepoint/ticketsApi';

export interface IHelpdeskTicketWebPartProps {
  /** Which screen the web part opens on; `?view=mine` in the address overrides it. */
  startView: View;
}

export default class HelpdeskTicketWebPart extends BaseClientSideWebPart<IHelpdeskTicketWebPartProps> {
  public render(): void {
    const { user, legacyPageContext, web } = this.context.pageContext;
    const fromAddress = new URLSearchParams(window.location.search).get('view');
    const element: React.ReactElement<IHelpdeskTicketProps> = React.createElement(HelpdeskTicket, {
      api: new TicketsApi(this.context.spHttpClient, web.absoluteUrl),
      user: {
        name: user.displayName,
        email: user.email,
        // The site user id, what the list's Created By holds.
        id: legacyPageContext?.userId as number
      },
      startView: fromAddress === 'mine' || fromAddress === 'raise' ? fromAddress : this.properties.startView || 'raise'
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('startView', {
                  label: strings.StartViewLabel,
                  options: [
                    { key: 'raise', text: strings.StartViewRaise },
                    { key: 'mine', text: strings.StartViewMine }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
