import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'HomePageWebPartStrings';
import HomePage, { arr_distinctParentVal } from './components/HomePage';
import { IHomePageProps } from './components/IHomePageProps';

export interface IHomePageWebPartProps {
  description: string;
  currenContext: WebPartContext;
  tileName: string;
}

export default class HomePageWebPart extends BaseClientSideWebPart <IHomePageWebPartProps> {
  
  //array to store the dynamic values of the property pane dropdown
  private parentItemDropdownOptions: IPropertyPaneDropdownOption[] =[];

  public render(): void {
    const element: React.ReactElement<IHomePageProps> = React.createElement(
      HomePage,
      {
        description: this.properties.description,
        currentContext: this.context,
        tileName:this.properties.tileName
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
    if(this.parentItemDropdownOptions.length<=0){
      let _count =0;
    arr_distinctParentVal.forEach(element => {
      this.parentItemDropdownOptions.push({key:element,text:element});
    });
  }
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
              groupName: "Select the Navigation",
              groupFields: [
                PropertyPaneDropdown('tileName', {
                  label: 'Select Tile Name',
                  options: this.parentItemDropdownOptions,
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
