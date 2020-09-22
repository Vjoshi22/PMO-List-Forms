import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import styles from './MilestonesGridWebPart.module.scss';
import * as strings from 'MilestonesGridWebPartStrings';
import { _getallItems, _populateGrid } from './components/getItemsMilestone';

let cssURL = "https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css";
SPComponentLoader.loadCss(cssURL);

export interface ISPMilestoneList{
  ID: string;
  ProjectID: string;    
  Phase: string;
  Milestone: string;
  PlannedStart: string;
  PlannedEnd:string;
  MilestoneStatus: string;
  Remarks: string;
  Created?: string;
  Modified?: string;
  ActualStart: string;
  ActualEnd: string;      
}

export interface IMilestonesGridWebPartProps {
  description: string;
  listGUID:string;
}

export default class MilestonesGridWebPart extends BaseClientSideWebPart <IMilestonesGridWebPartProps> {

  public render(): void {
    //this.domElement.innerHTML = ``;
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");

    let listGuid = 'e163f102-1cc9-4cc5-97b6-c5296811b444';
    this.domElement.innerHTML = `<div class="dataGrid"></div>`;

    let _ProjectId = this._getParameterValues('FilterValue1')

    let url = `/_api/web/lists('${this.properties.listGUID}')/items?$select=*&$filter=ProjectID eq '` + _ProjectId + `'&$orderby=Id desc`;
    let currentContext = this.context;
    _getallItems(url, currentContext, currentContext.pageContext.web.absoluteUrl).then((results) => {
      _populateGrid(results, currentContext);
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
