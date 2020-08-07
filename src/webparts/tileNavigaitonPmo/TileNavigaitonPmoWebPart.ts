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

export interface ITileNavigaitonPmoWebPartProps {
  description: string;
  currentContext: WebPartContext;
  listGUID: string;
  tileName: string;
}

export default class TileNavigaitonPmoWebPart extends BaseClientSideWebPart<
  ITileNavigaitonPmoWebPartProps
> {
  //array to store the dynamic values of the property pane dropdown
  private parentItemDropdownOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<ITileNavigaitonPmoProps> = React.createElement(
      TileNavigaitonPmo,
      {
        description: this.properties.description,
        currentContext: this.context,
        listGUID: this.properties.listGUID,
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
  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
                PropertyPaneTextField("listGUID", {
                  label: 'Enter the List GUID',
                })
              ]
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
