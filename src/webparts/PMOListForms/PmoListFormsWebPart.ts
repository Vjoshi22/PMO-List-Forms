import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'PmoListFormsWebPartStrings';
import PmoListForms from './components/PmoListForms';
import { IPmoListFormsProps } from './components/IPmoListFormsProps';

export interface IPmoListFormsWebPartProps {
  description: string;
  currentContext: WebPartContext;
}
export interface ISPList{
  ProjectID: string;
  ProjectID_SalesCRM: string;
  Project_x0020_Name: string;
  Client_x0020_Name: string;
  Delivery_x0020_Manager: string;
  Project_x0020_Manager: string;
  Project_x0020_Type: string;
  Project_x0020_Mode: string;
  PlannedStart: string;
  Planned_x0020_End: string;
  Project_x0020_Description: string;
  Region: string;
  Project_x0020_Budget: string;
  Status: string;
  Actual_x0020_Start:string;
  Actual_x0020_End:string;
  Revised_x0020_Budget:string;
  Total_x0020_Cost:string
}

export default class PmoListFormsWebPart extends BaseClientSideWebPart <IPmoListFormsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPmoListFormsProps> = React.createElement(
      PmoListForms,
      {
        description: this.properties.description,
        currentContext: this.context
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
