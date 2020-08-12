import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
} from "@microsoft/sp-property-pane";
import {
  BaseClientSideWebPart,
  WebPartContext,
} from "@microsoft/sp-webpart-base";

import * as strings from "TileNavigaitonPmoWebPartStrings";
import TileNavigaitonPmo from "./components/TileNavigaitonPmo";
import { ITileNavigaitonPmoProps } from "./components/ITileNavigaitonPmoProps";
import { arr_distinctParentVal } from "../tileNavigaitonPmo/components/TileNavigaitonPmo";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
import CheckBrowser from "../../checkBrowser";

export interface ITileNavigaitonPmoWebPartProps {
  description: string;
  currentContext: WebPartContext;
  lists: string | string[];
  tileName: string;
}
var renderTiles: any;

export default class TileNavigaitonPmoWebPart extends BaseClientSideWebPart<
  ITileNavigaitonPmoWebPartProps
> {
  //array to store the dynamic values of the property pane dropdown
  private parentItemDropdownOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    let userAgentString = navigator.userAgent;
    let IExplorerAgent =
      userAgentString.indexOf("MSIE") > -1 ||
      userAgentString.indexOf("rv:") > -1;
    //checking the current browser is IE, if IE then asking the user to use modern browsers
    if (IExplorerAgent) {
      renderTiles = CheckBrowser;
    } else {
      renderTiles = TileNavigaitonPmo;
    }
    const element: React.ReactElement<ITileNavigaitonPmoProps> = React.createElement(
      renderTiles,
      {
        description: this.properties.description,
        currentContext: this.context,
        lists: this.properties.lists,
        tileName: this.properties.tileName,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    if (this.parentItemDropdownOptions.length <= 0) {
      let _count = 0;
      arr_distinctParentVal.forEach((element) => {
        this.parentItemDropdownOptions.push({ key: element, text: element });
      });
    }

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
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
            {
              groupName: "Select the Navigation",
              groupFields: [
                PropertyPaneDropdown("tileName", {
                  label: "Select Tile Name",
                  options: this.parentItemDropdownOptions,
                  disabled: false,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
