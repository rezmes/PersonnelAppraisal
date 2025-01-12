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

export interface IPersonnelAppraisalWebPartProps {
  description: string;
  employeeListName: string;
  evaluationResultsListName: string;
  evaluationPeriodListName: string;
  questionBankListName: string;
}

export default class PersonnelAppraisalWebPart extends BaseClientSideWebPart<IPersonnelAppraisalWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPersonnelAppraisalProps> = React.createElement(
      PersonnelAppraisal,
      {
        description: this.properties.description,
        context: this.context,
        employeeListName: this.properties.employeeListName,
        questionBankListName: this.properties.questionBankListName,
        evaluationPeriodListName: this.properties.evaluationPeriodListName,
        evaluationResultsListName: this.properties.evaluationResultsListName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
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
                PropertyPaneTextField("employeeListName", {
                  label: "Employee List Name",
                }),
                PropertyPaneTextField("evaluationResultsListName", {
                  label: "Evaluation Results List Name",
                }),
                PropertyPaneTextField("evaluationPeriodListName", {
                  label: "Evaluation Period List Name",
                }),
                PropertyPaneTextField("questionBankListName", {
                  label: "Question Bank List Name",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
