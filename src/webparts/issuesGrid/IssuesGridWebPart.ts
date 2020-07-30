import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import styles from './IssuesGridWebPart.module.scss';
import * as strings from 'IssuesGridWebPartStrings';
import { _getallItems, _populateGrid } from './components/getItemsIssues';
// import { _getParameterValues } from '../PMOListForms/components/getQueryString';

let cssURL = "https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css";
SPComponentLoader.loadCss(cssURL);

export interface ISPIssueInformationList{
  ID: string;
  ProjectID: string;
  IssueCategory: string;
  IssueDescription: string;
  NextStepsOrResolution: string;
  IssueStatus: string;
  IssuePriority: string;
  Assignedteam: string;
  Assginedperson: string;
  IssueReportedOn: string;
  IssueClosedOn: string;
  RequiredDate: string;
}

export interface IIssuesGridWebPartProps {
  description: string;
}

export default class IssuesGridWebPart extends BaseClientSideWebPart <IIssuesGridWebPartProps> {

  public render(): void {
    //this.domElement.innerHTML = ``;
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");

    let listGuid = 'A373C7C3-3379-49C9-B3B1-AC87C2166DC0';
    this.domElement.innerHTML = `<div class="dataGrid"></div>`;

    let _ProjectId = this._getParameterValues('FilterValue1')

    let url = `/_api/web/lists('${listGuid}')/items?$select=*&$filter=ProjectID eq '` + _ProjectId + `'&$orderby=Id desc`;
    let currentContext = this.context;
    _getallItems(url, currentContext, currentContext.pageContext.web.absoluteUrl).then((results) => {
      _populateGrid(results);
      //_customStyle();
    });
  }
  private _getParameterValues(param) {
    var url = window.location.href.slice(window.location.href.indexOf('&') + 1).split('&');
    for (var i = 0; i < url.length; i++) {
        var urlparam = url[i].split('=');
        if (urlparam[0] == param) {
            return urlparam[1];
        }
    }
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
