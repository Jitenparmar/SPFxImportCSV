import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'ImportCsvWebPartStrings';
import ImportCsv from './components/ImportCsv';
import { IImportCsvProps } from './components/IImportCsvProps';
import { SPService } from "../../Services/SPService";
export interface IImportCsvWebPartProps {
  description: string;
}

export default class ImportCsvWebPart extends BaseClientSideWebPart <IImportCsvWebPartProps> {
  private SpServiceInstance:SPService;
  public render(): void {
    const element: React.ReactElement<IImportCsvProps> = React.createElement(
      ImportCsv,
      {
        description: this.properties.description,
        context: this.context,
        spHttpClient: this.context.spHttpClient,  
        siteUrl: this.context.pageContext.web.absoluteUrl,
        SPServiceInstance:this.SpServiceInstance,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public async onInit(){
    await super.onInit();
    this.SpServiceInstance = new SPService();
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
