import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'IssueInformationWebPartStrings';
import CreateIssue from './components/CreateIssue';
import UpdateIssue from './components/UpdateIssue';
import { IIssueInformationProps } from './components/IIssueInformationProps';
import { SPComponentLoader } from "@microsoft/sp-loader";

export interface IIssueInformationWebPartProps {
  description: string;
  currentContext: WebPartContext;
  customGridRequired: string;
  listGUID:string;
  ProjectMasterGUID:string;

}
var renderIssueForm: any;

let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);

export default class IssueInformationWebPart extends BaseClientSideWebPart <IIssueInformationWebPartProps> {

  public render(): void {
    //SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.0/css/bootstrap.css");
    if((/edit/.test(window.location.href))){
      renderIssueForm = UpdateIssue 
    }
    if((/new/.test(window.location.href))){
      renderIssueForm = CreateIssue
    }
    const element: React.ReactElement<IIssueInformationProps> = React.createElement(
      renderIssueForm,
      {
        description: this.properties.description,
        currentContext: this.context,
        customGridRequired: this.properties.customGridRequired,
        listGUID: this.properties.listGUID,
        ProjectMasterGUID:this.properties.ProjectMasterGUID
        
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
                  label: 'Enter the Project Master List GUID'
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
