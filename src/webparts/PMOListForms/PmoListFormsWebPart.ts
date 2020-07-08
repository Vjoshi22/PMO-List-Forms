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
import PmoListEditForm from "./components/PmoListEditForm";
import { IPmoListFormsProps } from './components/IPmoListFormsProps';

export interface IPmoListFormsWebPartProps {
  description: string;
  currentContext: WebPartContext;
}
var renderPMOForm: any;
export default class PmoListFormsWebPart extends BaseClientSideWebPart <IPmoListFormsWebPartProps> {
  
  public render(): void {
    if((/edit/.test(window.location.href))){
      renderPMOForm = PmoListEditForm 
    }
    if((/new/.test(window.location.href))){
      renderPMOForm = PmoListForms
    }
    const element: React.ReactElement<IPmoListFormsProps> = React.createElement(
      renderPMOForm,
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
