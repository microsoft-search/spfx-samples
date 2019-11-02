import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { 
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  BaseClientSideWebPart,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'GraphSearchApiWebPartStrings';
import GraphSearchApi from './components/GraphSearchApi';
import { IGraphSearchApiProps } from './components/IGraphSearchApiProps';
import { ClientMode } from './components/ClientMode';
//require('bootstrap');
//require('jsoneditor');

require('../../../src/jsoneditor/jsoneditor.css');

export interface IGraphSearchApiWebPartProps {
  clientMode: ClientMode;
}

export default class GraphSearchApiWebPart extends BaseClientSideWebPart<IGraphSearchApiWebPartProps> {

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.min.css');
    
    let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    SPComponentLoader.loadCss(cssURL);

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IGraphSearchApiProps > = React.createElement(
      GraphSearchApi,
      {
        clientMode: this.properties.clientMode,
        context: this.context,
        description : this.description
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
                PropertyPaneChoiceGroup('clientMode', {
                  label: strings.ClientModeLabel,
                  options: [
                    { key: ClientMode.aad, text: "AadHttpClient"},
                    { key: ClientMode.graph, text: "MSGraphClient"},
                  ]
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
