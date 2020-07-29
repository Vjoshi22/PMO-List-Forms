import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from "@microsoft/sp-loader";
import styles from './RisksGridWebPart.module.scss';
import * as strings from 'RisksGridWebPartStrings';
import { _getallItems, _populateGrid } from './components/getItemsRisk';
import { _customStyle } from "./components/customCssRisk";

let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css";
SPComponentLoader.loadCss(cssURL);

export interface IRiskInformationList {
  Title?: string;  
  RiskID: string;//ID in SP
  ProjectID: string;
  RiskName: string;
  RiskDescription: string;
  RiskCategory: string;
  RiskIdentifiedOn: string;
  RiskClosedOn: string;
  RiskStatus: string;
  RiskOwner: string;
  RiskResponse: string;
  RiskImpact: string;
  RiskProbability: string;
  RiskRank: string;
  Remarks: string;
}
export interface IRisksGridWebPartProps {
  description: string;
}

export default class RisksGridWebPart extends BaseClientSideWebPart <IRisksGridWebPartProps> {

  public render(): void {
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");

    let listGuid = 'b94d8766-9e5a-41ae-afc6-b00a0bbe0149';
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
