// import * as React from 'react';
// import * as ReactDom from 'react-dom';
// import { Version } from '@microsoft/sp-core-library';
// import {
//   BaseClientSideWebPart,
//   IPropertyPaneConfiguration,
//   PropertyPaneTextField
// } from '@microsoft/sp-webpart-base';

// import * as strings from 'PersonnelAppraisalWebPartStrings';
// import PersonnelAppraisal from './components/PersonnelAppraisal';
// import { IPersonnelAppraisalProps } from './components/IPersonnelAppraisalProps';

// export interface IPersonnelAppraisalWebPartProps {
//   description: string;
// }

// export default class PersonnelAppraisalWebPart extends BaseClientSideWebPart<IPersonnelAppraisalWebPartProps> {

//   public render(): void {
//     const element: React.ReactElement<IPersonnelAppraisalProps > = React.createElement(
//       PersonnelAppraisal,
//       {
//         description: this.properties.description
//       }
//     );

//     ReactDom.render(element, this.domElement);
//   }

//   protected onDispose(): void {
//     ReactDom.unmountComponentAtNode(this.domElement);
//   }

//   protected get dataVersion(): Version {
//     return Version.parse('1.0');
//   }

//   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
//     return {
//       pages: [
//         {
//           header: {
//             description: strings.PropertyPaneDescription
//           },
//           groups: [
//             {
//               groupName: strings.BasicGroupName,
//               groupFields: [
//                 PropertyPaneTextField('description', {
//                   label: strings.DescriptionFieldLabel
//                 })
//               ]
//             }
//           ]
//         }
//       ]
//     };
//   }
// }
import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import * as strings from "PersonnelAppraisalWebPartStrings";
import PersonnelAppraisal from "./components/PersonnelAppraisal";
import { IPersonnelAppraisalProps } from "./components/IPersonnelAppraisalProps";

export interface IPersonnelAppraisalWebPartProps {
  description: string;
}

export default class PersonnelAppraisalWebPart extends BaseClientSideWebPart<IPersonnelAppraisalWebPartProps> {
  public render(): void {
    const element = React.createElement<PersonnelAppraisal>(
      PersonnelAppraisal,
      {
        description: this.properties.description
      } as IPersonnelAppraisalProps // Explicitly type the props
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  public get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
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
