import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'IssueInformationWebPartStrings';
import CreateIssue from './components/CreateIssue';
import UpdateIssue from './components/UpdateIssue';
import { IIssueInformationProps } from './components/IIssueInformationProps';

export interface IIssueInformationWebPartProps {
  description: string;
  currentContext: WebPartContext;
}
var renderIssueForm: any;

export default class IssueInformationWebPart extends BaseClientSideWebPart <IIssueInformationWebPartProps> {

  public render(): void {
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
