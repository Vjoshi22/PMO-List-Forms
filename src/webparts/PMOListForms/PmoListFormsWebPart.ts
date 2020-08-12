import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import {
  BaseClientSideWebPart,
  WebPartContext,
} from "@microsoft/sp-webpart-base";

import * as strings from "PmoListFormsWebPartStrings";
import PmoListForms from "./components/PmoListForms";
import PmoListEditForm from "./components/PmoListEditForm";
import CheckBrowser from "../../checkBrowser";
import { IPmoListFormsProps } from "./components/IPmoListFormsProps";
import {
  SPHttpClient,
  ISPHttpClientOptions,
  SPHttpClientConfiguration,
  SPHttpClientResponse,
  HttpClientResponse,
} from "@microsoft/sp-http";

export interface IPmoListFormsWebPartProps {
  description: string;
  currentContext: WebPartContext;
  customGridRequired: string;
  listGUID: string;
}
var renderPMOForm: any;

export default class PmoListFormsWebPart extends BaseClientSideWebPart<
  IPmoListFormsWebPartProps
> {
  public render(): void {
    // //to fetch the choice column options
    // allchoiceColumns.forEach(element => {
    //   this.retrieveAllChoicesFromListField(this.context.pageContext.web.absoluteUrl, element);
    // });
    let userAgentString = navigator.userAgent;
    let IExplorerAgent =
      userAgentString.indexOf("MSIE") > -1 ||
      userAgentString.indexOf("rv:") > -1;
    //checking the current browser is IE, if IE then asking the user to use modern browsers
    if (IExplorerAgent) {
      renderPMOForm = CheckBrowser;
    } else {
      if (/edit/.test(window.location.href)) {
        renderPMOForm = PmoListEditForm;
      }
      if (/new/.test(window.location.href)) {
        renderPMOForm = PmoListForms;
      }
    }
    const element: React.ReactElement<IPmoListFormsProps> = React.createElement(
      renderPMOForm,
      {
        description: this.properties.description,
        currentContext: this.context,
        customGridRequired: this.properties.customGridRequired,
        listGUID: this.properties.listGUID
      }
    );

    ReactDom.render(element, this.domElement);
  }
  private retrieveAllChoicesFromListField(
    siteColUrl: string,
    entitySetName: string
  ): void {
    const endPoint: string = `${siteColUrl}/_vti_bin/listdata.svc/${entitySetName}`;

    this.context.spHttpClient
      .get(endPoint, SPHttpClient.configurations.v1)
      .then((response: HttpClientResponse) => {
        if (response.ok) {
          response.json().then(
            (jsonResponse: JSON) => {
              for (const result of jsonResponse["d"]["results"]) {
                console.log(result["Value"]);
              }
            },
            (err: any): void => {
              console.warn(`Failed to fulfill Promise\r\n\t${err}`);
            }
          );
        } else {
          console.warn(
            `List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`
          );
        }
      });
  }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
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
            },
            {
              groupName: "Custom Required",
              groupFields: [
                PropertyPaneToggle("customGridRequired", {
                  label: "Custom Grid",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
