import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import { SPCommonOperations } from 'spfxhelper';

import * as strings from 'DemoDynamicValuesInWebPartPropertiesWebPartStrings';
import DemoDynamicValuesInWebPartProperties from './components/DemoDynamicValuesInWebPartProperties';
import { IDemoDynamicValuesInWebPartPropertiesProps } from './components/IDemoDynamicValuesInWebPartPropertiesProps';

export interface IDemoDynamicValuesInWebPartPropertiesWebPartProps {
  selectedList:string;
}

export default class DemoDynamicValuesInWebPartPropertiesWebPart extends BaseClientSideWebPart<IDemoDynamicValuesInWebPartPropertiesWebPartProps> {

   // Property that will hold all the options value
   private ddlListOptions: IPropertyPaneDropdownOption[] = [];

   // Property that will be used flag if all the values are filled
   private receivedLists: boolean = false;
   
  public render(): void {
    const element: React.ReactElement<IDemoDynamicValuesInWebPartPropertiesProps > = React.createElement(
      DemoDynamicValuesInWebPartProperties,
      {
        description: this.properties.selectedList
      }
    );

    ReactDom.render(element, this.domElement);
  }

  
  /** Property to get the current web URL **/
  private get webUrl(): string {
    return this.context.pageContext.web.absoluteUrl;
  }

  /** Property to get the SPCommonOperations object, which will help in fetching values **/
  private get oSPCommonOps(): SPCommonOperations {
    return SPCommonOperations.getInstance(this.context.spHttpClient as any, this.webUrl);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** Method to get all the lists from SharePoint */
  private get getAllLists(): Promise<IPropertyPaneDropdownOption[]> {

    // Initialize the variable to geather all values
    let options: IPropertyPaneDropdownOption[] = [];

    // Call the method to fetch record from SharePoint based on query
    return this.oSPCommonOps.queryGETResquest(`${this.webUrl}/_api/web/lists`).then(response => {

      // Check if the response is success ?
      if (response.ok) {
        // Iterate over each response and get the title to fill in as ddl options
        response.result.value.forEach(item => {
          options.push({ key: item.Title, text: item.Title });
        });

        // return the fetched records
        return Promise.resolve(options);
      }
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    // Check if the values are recieved
    if (!this.receivedLists) {

      // Call the method to get all the lists Titles
      this.getAllLists.then(resp=>{
        // Fill the values in the variable assigned
        this.ddlListOptions = resp;
        // update the flag so it is not called again
        this.receivedLists =true;
        // Refresh the property pane, to reflect the changes
        this.context.propertyPane.refresh();
      });
    }

    return {
      pages: [
        {
          header: {
            description: `Demo By Sumit Kanchan`
          },
          groups: [
            {
              groupName: 'Demo for Dynamic values in Web Part Properties',
              groupFields: [
                PropertyPaneDropdown('selectedList',{
                  label: `select list`,
                  options: this.ddlListOptions
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
