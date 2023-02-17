import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ProjectsWebPartStrings';
import Projects from './components/Projects';
import { IProjectsProps } from './components/IProjectsProps';

export interface IProjectsWebPartProps {
  description: string;
}

export default class ProjectsWebPart extends BaseClientSideWebPart<IProjectsWebPartProps> {

  public render(): void {
     // Getting Item ID from URL Parameter -   
     const queryParams = new URLSearchParams(window.location.search);
     const itemID = queryParams.get('ItemID');
     const element: React.ReactElement<IProjectsProps> = React.createElement(
       Projects,
       {
         AppContext: this.context,
         ItemID: itemID
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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
