import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-webpart-base";

import * as strings from "PersonnelAppraisalWebPartStrings";
import PersonnelAppraisal from "./components/PersonnelAppraisal";
import { IPersonnelAppraisalProps } from "./components/IPersonnelAppraisalProps";

import { sp } from "@pnp/sp";

export interface IPersonnelAppraisalWebPartProps {
  description: string;
  selectedDepartment: string;
 }

export default class PersonnelAppraisalWebPart extends BaseClientSideWebPart<
  IPersonnelAppraisalWebPartProps
> {

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }


  public render(): void {
    // Ensure props match the expected IPersonnelAppraisalProps interface
    const element: React.ReactElement<IPersonnelAppraisalProps> = React.createElement(
      PersonnelAppraisal,
      {
        description: this.properties.description, // Pass the description prop correctly
        context: this.context, // Pass context explicitly
        selectedDepartment: this.properties.selectedDepartment // Ensure this property exists in your WebPart properties
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // public get dataVersion(): Version {
  //   return Version.parse("1.0");
  // }

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
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
