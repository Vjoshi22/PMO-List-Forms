import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import styles from './ProjectGridWebPart.module.scss';
import * as strings from 'ProjectsGridWebPartStrings';
import { _getallItems, _populateGrid } from "./Components/getItemsProject";

import "datatables.net";
import "datatables.net-dt";
import "datatables.net-responsive";
//import styles from './ProjectsGridWebPart.module.scss';
//import * as strings from 'ProjectsGridWebPartStrings';

let cssURL = "https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css";
SPComponentLoader.loadCss(cssURL);
//SPComponentLoader.loadCss("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");

export interface IProjectsGridWebPartProps {
  description: string;
  listName: string;
  listGUID: string;
}

export interface ISPProjectsList {
  ProjectID: string,
  ProjectName: string;
  ClientName: string;
  ProjectManager: string;
  ProjectType: string;
  ProjectMode: string;
  PlannedStart: string;
  PlannedCompletion: string;
  ProjectDescription: string;
  ProjectLocation: string;
  // ProjectBudget: string;
  ProjectStatus: string;
  ProjectProgress: string;
  ActualStartDate: string; //edit only
  ActualEndDate: string; //edit only
  RevisedBudget: string; //edit only
  TotalCost: string; //edit only
  InvoicedAmount: string; //edit only
  ProjectScope: string; // Project Scope edit only
  ProjectSchedule: string; //project scheduled edit only
  ProjectResource: string;
  ProjectCost: string; //only in edit
  //peoplepicker
  DeliveryManager: string;
  //date
  startDate: any;
  disable_RMSID: boolean;
  disable_plannedCompletion: boolean;
  endDate: any;
}

export default class ProjectsGridWebPart extends BaseClientSideWebPart<IProjectsGridWebPartProps> {  
  public render(): void {
    // SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css");
    // SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/js/bootstrap.min.js");
     SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");
    // SPComponentLoader.loadCss("//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css");
    
    let listGuid = '2c3ffd4e-1b73-4623-898d-8e3a1bb60b91';
    this.domElement.innerHTML = `<div class="dataGrid"></div>`;

    let url = `/_api/web/lists('${this.properties.listGUID}')/items?$select=*&$orderby=ID desc`;
    let currentContext = this.context;
    _getallItems(url, currentContext, currentContext.pageContext.web.absoluteUrl).then((results) => {
      _populateGrid(results, currentContext);
      //_customStyle();
    });
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
                }),
                PropertyPaneTextField('listGUID', {
                  label: 'Enter the List GUID'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
