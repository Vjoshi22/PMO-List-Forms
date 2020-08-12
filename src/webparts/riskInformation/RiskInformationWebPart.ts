import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart,WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'RiskInformationWebPartStrings';
import RiskInformationNew from './components/RiskInformationNew';
import RiskInformationEdit from './components/RiskInformationEdit';
import { IRiskInformationProps } from './components/IRiskInformationProps';
import CheckBrowser from '../../checkBrowser';

export interface IRiskInformationWebPartProps {
  description: string;
  currentContext: WebPartContext;
  customGridRequired: string;
}

export var allchoiceColumns: any[] = ["RiskCategory", "RiskStatus", "RiskResponse", "RiskImpact", "RiskProbability"];

export default class RiskInformationWebPart extends BaseClientSideWebPart <IRiskInformationWebPartProps> {

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
        renderPMOForm = RiskInformationEdit 
      }
      if((/new/.test(window.location.href))){
        renderPMOForm = RiskInformationNew
      }
    }

    
    const element: React.ReactElement<IRiskInformationProps> = React.createElement(
      renderPMOForm,
      {
        description: this.properties.description,
        currentContext: this.context,
        customGridRequired: this.properties.customGridRequired
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
