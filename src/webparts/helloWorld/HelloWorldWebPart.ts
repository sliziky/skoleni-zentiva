import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'HelloWorldWebPartStrings';
import { HelloWorld } from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';

export interface IHelloWorldWebPartProps {
  description: string;
  slider: number;
  context: WebPartContext;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        description: this.properties.description,
        context: this.context,
        slider: this.properties.slider
      }
    );
    // <HelloWorld description={this.properties.description} context={this.context}/> 
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneSlider('slider', {
                  min: 0,
                  max: 20,
                  label: 'Number of people'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
