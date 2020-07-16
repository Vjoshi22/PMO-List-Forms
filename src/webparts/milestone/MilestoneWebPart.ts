import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart,WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'MilestoneWebPartStrings';
import MilestoneNew from './components/MilestoneNew';
import MilestoneEdit from './components/MilestoneEdit';
import { IMilestoneProps } from './components/IMilestoneProps';

export interface IMilestoneWebPartProps {
  description: string;
  currentContext: WebPartContext;
}

export let allchoiceColumns: any[] = ["Phase", "MilestoneStatus"];

export default class MilestoneWebPart extends BaseClientSideWebPart <IMilestoneWebPartProps> {

  public render(): void {
    let renderPMOForm: any;
    if((/edit/.test(window.location.href))){
      renderPMOForm = MilestoneEdit
    }
    else{
      renderPMOForm = MilestoneNew
    }
    
    const element: React.ReactElement<IMilestoneProps> = React.createElement(
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
