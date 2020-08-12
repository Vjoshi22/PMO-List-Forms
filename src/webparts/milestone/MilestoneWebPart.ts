import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart,WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'MilestoneWebPartStrings';
import MilestoneNew from './components/MilestoneNew';
import MilestoneEdit from './components/MilestoneEdit';
import { IMilestoneProps } from './components/IMilestoneProps';
import CheckBrowser from '../../checkBrowser';

export interface IMilestoneWebPartProps {
  description: string;
  currentContext: WebPartContext;
  customGridRequired: string;
  listGUID:string;
  ProjectMasterGUID:string;
}

export let allchoiceColumns: any[] = ["Phase", "MilestoneStatus"];

export default class MilestoneWebPart extends BaseClientSideWebPart <IMilestoneWebPartProps> {

  public render(): void {
    let renderPMOForm: any;
    
    let userAgentString = navigator.userAgent;
    let IExplorerAgent =
      userAgentString.indexOf("MSIE") > -1 ||
      userAgentString.indexOf("rv:") > -1;
    //checking the current browser is IE, if IE then asking the user to use modern browsers
    if (IExplorerAgent) {
      renderPMOForm = CheckBrowser;
    } else {
      if((/edit/.test(window.location.href))){
        renderPMOForm = MilestoneEdit
      }
      else{
        renderPMOForm = MilestoneNew
      }
    }
    
    
    const element: React.ReactElement<IMilestoneProps> = React.createElement(
      renderPMOForm,
      {
        description: this.properties.description,
        currentContext: this.context,
        customGridRequired: this.properties.customGridRequired,
        listGUID:this.properties.listGUID,
        ProjectMasterGUID: this.properties.ProjectMasterGUID
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

  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
                }),
                PropertyPaneTextField('listGUID', {
                  label: 'Enter the list GUID'
                }),
                PropertyPaneTextField('ProjectMasterGUID', {
                  label: 'Enter the Project Master GUID'
                })
              ]
            },
            {
              groupName: "Custom Grid",
              groupFields: [
                PropertyPaneToggle('customGridRequired', {
                  label: "Custom Grid Required"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
