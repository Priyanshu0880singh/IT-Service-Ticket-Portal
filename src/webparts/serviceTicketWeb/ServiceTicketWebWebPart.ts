import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import ServiceTicketWeb, { IServiceTicketWebProps } from './components/ServiceTicketWeb';

export interface IServiceTicketWebWebPartProps {
  // No configurable properties needed - the component gets everything it
  // needs from `context`. Keep this interface even if empty so the
  // property pane below still type-checks.
}

export default class ServiceTicketWebWebPart extends BaseClientSideWebPart<IServiceTicketWebWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IServiceTicketWebProps> = React.createElement(
      ServiceTicketWeb,
      {
        context: this.context
      }
    );

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
          header: { description: 'IT Service Desk' },
          groups: [
            {
              groupName: 'Settings',
              groupFields: []
            }
          ]
        }
      ]
    };
  }
}
