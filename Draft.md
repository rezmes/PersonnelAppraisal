# Root cause finding

It doesn't fetch any question.
I will send all of my code and you check it needs to any change or not. I prefer not making change in anything, becuase in the last version it was woking.

#### _SharePoint 2019 - On-premises_

#### dev.env. : `SPFx@1.4.1 ( node@8.17.0 , react@15.6.2, typescript@2.4.2 ,update and upgrade are not options)`

```tsx PersonnelAppraisal
import * as React from "react";
import {
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Label,
  Dialog,
  DialogType,
  DialogFooter,
  IDropdownOption,
} from "office-ui-fabric-react";
import { IPersonnelAppraisalProps } from "./IPersonnelAppraisalProps";
import EmployeeDropdown from "./EmployeeDropdown";
import QuestionTable from "./QuestionTable";
import EvaluationPeriod from "./EvaluationPeriod";
import "./PersonnelAppraisal.module.scss";

// Define and export the interface separately
export interface IEmployeeOption extends IDropdownOption {
  department: string;
  departmentGuid: string;
}

// Define and export the interface separately
export interface IAppraisalFormState {
  employees: IEmployeeOption[];
  selectedEmployee: string | number | undefined;
  questions: { id: number; text: string; weight: number }[];
  scores: { [questionId: number]: number };
  isLoading: boolean;
  errorMessage: string | null;
  isDialogHidden: boolean;
  evaluationPeriod: string;
}

// Use the interfaces in the class definition
export default class PersonnelAppraisal extends React.Component<
  IPersonnelAppraisalProps,
  IAppraisalFormState
> {
  constructor(props: IPersonnelAppraisalProps) {
    super(props);

    this.state = {
      employees: [],
      selectedEmployee: undefined,
      questions: [],
      scores: {},
      isLoading: false,
      errorMessage: null,
      isDialogHidden: true,
      evaluationPeriod: "",
    };
  }

  componentDidMount(): void {
    this.loadEmployees();
  }

  private handlePeriodLoaded = (period: string): void => {
    this.setState({ evaluationPeriod: period });
  };

  private async loadEmployees(): Promise<void> {
    try {
      this.setState({ isLoading: true });
      const response = await fetch(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/web/currentUser`,
        {
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }
      );
      const currentUser = await response.json();
      console.log("Current User:", currentUser);

      const encodedLoginName = encodeURIComponent(currentUser.d.LoginName);
      const listName = encodeURIComponent(this.props.employeeListName);

      // Corrected filter for Evaluator
      const employeesUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=ID,Title,FirstName,FieldValuesAsText/MechDepartment,Evaluator/Name&$expand=FieldValuesAsText,Evaluator&$filter=Evaluator/Name eq '${encodedLoginName}'`;
      console.log("Employees URL with corrected filter:", employeesUrl);

      const employeesResponse = await fetch(employeesUrl, {
        headers: {
          Accept: "application/json;odata=verbose",
        },
      });

      if (!employeesResponse.ok) {
        throw new Error(
          `Error fetching employees: ${employeesResponse.statusText}`
        );
      }

      const employees = await employeesResponse.json();
      console.log("Employees API Response:", employees);

      const employeeOptions: IEmployeeOption[] = employees.d.results.map(
        (emp: any) => {
          const mechDepartment = emp.FieldValuesAsText
            ? emp.FieldValuesAsText.MechDepartment || ""
            : "";
          console.log(
            "MechDepartment Value:",
            emp.FieldValuesAsText.MechDepartment
          ); // Added log
          const departmentGuid = mechDepartment.includes("|")
            ? mechDepartment.split("|")[1]
            : "";
          return {
            key: emp.ID,
            text: `${emp.FirstName} ${emp.Title}`,
            department: mechDepartment,
            departmentGuid: departmentGuid,
          };
        }
      );

      this.setState({ employees: employeeOptions, isLoading: false });
    } catch (error) {
      this.setState({
        errorMessage: `Error loading employees: ${error.message}`,
        isLoading: false,
      });
      console.error("Error loading employees:", error);
    }
  }

  private handleEmployeeChange = (option?: IDropdownOption): void => {
    if (option) {
      let selectedEmployee: IEmployeeOption | undefined = undefined;
      for (let i = 0; i < this.state.employees.length; i++) {
        if (this.state.employees[i].key === option.key) {
          selectedEmployee = this.state.employees[i];
          break;
        }
      }

      const selectedDepartmentGuid = selectedEmployee
        ? selectedEmployee.departmentGuid
        : "";
      console.log("Selected Department Guid:", selectedDepartmentGuid);

      this.setState(
        { selectedEmployee: option.key as string, questions: [] },
        () => {
          this.loadQuestions(selectedDepartmentGuid);
        }
      );
    }
  };

  private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
    try {
      this.setState({ isLoading: true });

      const questionsUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.questionBankListName}')/items?$select=ID,Title,QuestionWeight,MechDepartment`;
      console.log("Questions URL:", questionsUrl);

      const questionsResponse = await fetch(questionsUrl, {
        headers: {
          Accept: "application/json;odata=verbose",
        },
      });

      if (!questionsResponse.ok) {
        throw new Error(
          `Error fetching questions: ${questionsResponse.statusText}`
        );
      }

      const questions = await questionsResponse.json();
      console.log("Questions API Response:", questions);

      // Debugging: Log the retrieved and filtered questions
      const filteredQuestions = questions.d.results.filter(
        (q: any) => q.MechDepartment.TermGuid === selectedDepartmentGuid
      );
      console.log("Filtered Questions:", filteredQuestions);

      this.setState({
        questions: filteredQuestions.map((q: any) => ({
          id: q.ID,
          text: q.Title,
          weight: q.QuestionWeight,
        })),
        scores: {},
        isLoading: false,
      });
    } catch (error) {
      this.setState({
        errorMessage: `Error loading questions: ${error.message}`,
        isLoading: false,
      });
      console.error("Error fetching questions:", error);
    }
  }

  private handleScoreChange = (questionId: number, score: number): void => {
    this.setState((prevState) => ({
      scores: {
        ...prevState.scores,
        [questionId]: score,
      },
    }));
  };

  private handleSubmit: () => Promise<void> = async (): Promise<void> => {
    const { selectedEmployee, scores, questions, evaluationPeriod } =
      this.state;

    if (!selectedEmployee) {
      this.setState({ errorMessage: "Please select an employee." });
      return;
    }

    if (Object.keys(scores).length !== questions.length) {
      this.setState({ errorMessage: "Please rate all questions." });
      return;
    }

    try {
      this.setState({ isLoading: true, errorMessage: null });

      const batchOperations = questions.map((question) => {
        const weightedScore = (scores[question.id] / 5) * question.weight;

        const item = {
          EmployeeIDId: selectedEmployee,
          QuestionDescription: question.text,
          Score: scores[question.id],
          WeightedScore: weightedScore,
          EvaluationPeriod: evaluationPeriod,
        };

        return fetch(
          `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.evaluationResultsListName}')/items`,
          {
            method: "POST",
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
            },
            body: JSON.stringify(item),
          }
        );
      });

      await Promise.all(batchOperations);

      this.setState({ isLoading: false, questions: [] });
      alert("ارزیابی با موفقیت ثبت شد");

      this.loadEmployees();
    } catch (error) {
      this.setState({
        errorMessage: `Error submitting evaluation: ${error.message}`,
        isLoading: false,
      });
      console.error("Error submitting evaluation:", error);
    }
  };

  private closeDialog = (): void => {
    this.setState({ isDialogHidden: true });
  };

  render(): React.ReactElement<any> {
    const {
      employees,
      selectedEmployee,
      questions,
      scores,
      isLoading,
      errorMessage,
      isDialogHidden,
    } = this.state;

    const isRtl = this.props.context.pageContext.cultureInfo.isRightToLeft;

    return (
      <div dir={isRtl ? "rtl" : "ltr"}>
        <h3>ارزیابی عملکرد کارکنان</h3>
        <EvaluationPeriod
          spfxContext={this.props.context}
          onPeriodLoaded={this.handlePeriodLoaded}
        />
        {isLoading && <Spinner size={SpinnerSize.large} label="بارگذاری ..." />}
        {errorMessage && <Label style={{ color: "red" }}>{errorMessage}</Label>}
        <EmployeeDropdown
          employees={employees}
          selectedEmployee={selectedEmployee}
          onChange={this.handleEmployeeChange}
        />
        {questions.length > 0 && (
          <QuestionTable
            questions={questions}
            scores={scores}
            onScoreChange={this.handleScoreChange}
          />
        )}
        <PrimaryButton text="ثبت" onClick={this.handleSubmit} />
        <Dialog
          hidden={isDialogHidden}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Some Title",
            subText: "Some subtitle",
          }}
          modalProps={{
            isBlocking: false,
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this.closeDialog} text="OK" />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
```

<!-- ================================================================================================================================================= -->

```tsx QuestionTable.tsx
import * as React from "react";
import "./PersonnelAppraisal.module.scss"; // Import your styles

export interface IQuestion {
  id: number;
  text: string;
  weight: number;
}

// Exporting the interface separately
export interface IQuestionTableProps {
  questions: IQuestion[];
  scores: { [questionId: number]: number };
  onScoreChange: (questionId: number, score: number) => void;
}

class QuestionTable extends React.Component<IQuestionTableProps, {}> {
  public render() {
    return (
      <table>
        <thead>
          <tr>
            <th>شاخص ارزیابی</th>
            <th>امتیاز</th>
          </tr>
        </thead>
        <tbody>
          {this.props.questions.map((question) => (
            <tr key={question.id}>
              <td>{question.text}</td>
              <td>
                {[1, 2, 3, 4, 5].map((score) => (
                  <label key={score}>
                    <input
                      type="radio"
                      name={`question-${question.id}`}
                      value={score}
                      checked={this.props.scores[question.id] === score}
                      onChange={() =>
                        this.props.onScoreChange(question.id, score)
                      }
                    />
                    {score}
                  </label>
                ))}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    );
  }
}

export default QuestionTable;
```

<!-- ================================================================================================================================================= -->

```tsx EvaluationPeriod
import * as React from "react";
// Define and export the interfaces separately
export interface IEvaluationPeriodState {
  evaluationPeriod: string;
  isLoading: boolean;
  errorMessage: string | null;
}

export interface IEvaluationPeriodProps {
  spfxContext: any;
  onPeriodLoaded: (period: string) => void;
}

// Use the interfaces in the class definition
class EvaluationPeriod extends React.Component<
  IEvaluationPeriodProps,
  IEvaluationPeriodState
> {
  constructor(props: IEvaluationPeriodProps) {
    super(props);
    this.state = {
      evaluationPeriod: "",
      isLoading: false,
      errorMessage: null,
    };
  }

  componentDidMount(): void {
    this.loadLatestPeriod();
  }

  private async loadLatestPeriod(): Promise<void> {
    try {
      this.setState({ isLoading: true });

      const response = await fetch(
        `${this.props.spfxContext.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EvaluationPeriod')/items?$orderby=Created desc&$top=1`,
        {
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }
      );

      const data = await response.json();

      if (data.d.results.length > 0) {
        const evaluationPeriod = data.d.results[0].Title;
        this.setState({ evaluationPeriod, isLoading: false });
        this.props.onPeriodLoaded(evaluationPeriod);
      } else {
        this.setState({
          errorMessage: "دوره ارزیابی یافت نشد.",
          isLoading: false,
        });
      }
    } catch (error) {
      this.setState({
        errorMessage: "خطا در بارگذاری دوره ارزیابی",
        isLoading: false,
      });
      console.error("بارگذاری این دوره با خطا مواجه شد: ", error);
    }
  }

  render(): React.ReactElement<any> {
    const { evaluationPeriod, isLoading, errorMessage } = this.state;

    if (isLoading) {
      return <div>بارگذاری دوره ی ارزیابی ...</div>;
    }

    if (errorMessage) {
      return <div style={{ color: "red" }}>{errorMessage}</div>;
    }

    return <div id="PeriodTitle">دوره ارزیابی جاری: {evaluationPeriod}</div>;
  }
}

export default EvaluationPeriod;
```

<!-- ================================================================================================================================================= -->

```tsx EmployeeDropdown
import * as React from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react";

export interface IEmployeeDropdownProps {
  employees: IDropdownOption[];
  selectedEmployee: string | number | undefined;
  onChange: (option?: IDropdownOption) => void;
}

class EmployeeDropdown extends React.Component<IEmployeeDropdownProps, {}> {
  render() {
    const { employees, selectedEmployee, onChange } = this.props;
    const placeHolderText =
      employees.length === 0
        ? "همه افراد لیست شما برای این دوره ارزیابی شده اند."
        : "ارزیابی شونده را انتخاب کنید";
    return (
      <Dropdown
        placeHolder={placeHolderText}
        options={employees}
        onChanged={onChange}
        selectedKey={selectedEmployee}
      />
    );
  }
}
export default EmployeeDropdown;
```

<!-- ================================================================================================================================================= -->

```ts IPersonnelAppraisalProps
export interface IPersonnelAppraisalProps {
  description: string;
  context: any;
  employeeListName: string;
  evaluationResultsListName: string;
  evaluationPeriodListName: string;
  questionBankListName: string;
}
```

<!-- ================================================================================================================================================= -->

```ts PersonnelAppraisalWebPart
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
    const element: React.ReactElement<IPersonnelAppraisalProps> =
      React.createElement(PersonnelAppraisal, {
        description: this.properties.description,
        context: this.context,
        employeeListName: this.properties.employeeListName,
        questionBankListName: this.properties.questionBankListName,
        evaluationPeriodListName: this.properties.evaluationPeriodListName,
        evaluationResultsListName: this.properties.evaluationResultsListName,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

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
```

<!-- ================================================================================================================================================= -->

```json Package
{
  "name": "personnel-appraisal",
  "version": "0.0.1",
  "private": true,
  "main": "lib/index.js",
  "engines": {
    "node": ">=0.10.0"
  },
  "scripts": {
    "build": "gulp bundle",
    "clean": "gulp clean",
    "test": "gulp test"
  },
  "dependencies": {
    "@microsoft/sp-core-library": "~1.4.0",
    "@microsoft/sp-lodash-subset": "~1.4.0",
    "@microsoft/sp-office-ui-fabric-core": "~1.4.0",
    "@microsoft/sp-webpart-base": "~1.4.0",
    "@pnp/sp": "^2.0.9",
    "@types/es6-promise": "0.0.33",
    "@types/react": "15.6.6",
    "@types/react-dom": "15.5.6",
    "@types/webpack-env": "1.13.1",
    "jspdf": "^1.5.3",
    "jspdf-autotable": "^3.2.4",
    "react": "15.6.2",
    "react-dom": "15.6.2"
  },
  "resolutions": {
    "@types/react": "15.6.6"
  },
  "devDependencies": {
    "@microsoft/sp-build-web": "~1.4.1",
    "@microsoft/sp-module-interfaces": "~1.4.1",
    "@microsoft/sp-webpart-workbench": "~1.4.1",
    "gulp": "~3.9.1",
    "@types/chai": "3.4.34",
    "@types/mocha": "2.2.38",
    "ajv": "~5.2.2"
  }
}
```

<!-- ================================================================================================================================================= -->

```json tsconfig
{
  "compilerOptions": {
    "target": "es5",
    "forceConsistentCasingInFileNames": true,
    "module": "esnext",
    "moduleResolution": "node",
    "jsx": "react",
    "declaration": true,
    "sourceMap": true,
    "experimentalDecorators": true,
    "skipLibCheck": true,
    "typeRoots": ["./node_modules/@types", "./node_modules/@microsoft"],
    "types": ["es6-promise", "webpack-env"],
    "lib": ["es5", "dom", "es2015.collection"]
  }
}
```

<!-- ================================================================================================================================================= -->

```js gulpfile
"use strict";
const build = require("@microsoft/sp-build-web");

build.addSuppression(
  `Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`
);

build.initialize(require("gulp"));
```

<!-- ================================================================================================================================================= -->

```js console.log()
Current User: Objectd: Alerts: __deferred: uri: "https://apps.sarirpey.com/_api/Web/GetUserById(1)/Alerts"[[Prototype]]: Object[[Prototype]]: ObjectEmail: ""Groups: __deferred: uri: "https://apps.sarirpey.com/_api/Web/GetUserById(1)/Groups"[[Prototype]]: Object[[Prototype]]: ObjectId: 1IsEmailAuthenticationGuestUser: falseIsHiddenInUI: falseIsShareByEmailGuestUser: falseIsSiteAdmin: trueLoginName: "i:0#.w|sarirpey\\vradmin"PrincipalType: 1Title: "Farm administrator user account"UserId: NameId: "s-1-5-21-4145149021-2059769573-471417889-1202"NameIdIssuer: "urn:office:idp:activedirectory"__metadata: type: "SP.UserIdInfo"[[Prototype]]: Object[[Prototype]]: Object__metadata: id: "https://apps.sarirpey.com/_api/Web/GetUserById(1)"type: "SP.User"uri: "https://apps.sarirpey.com/_api/Web/GetUserById(1)"[[Prototype]]: Object[[Prototype]]: Object[[Prototype]]: Object
PersonnelAppraisal.tsx:516 Employees URL with corrected filter: https://apps.sarirpey.com/_api/web/lists/getbytitle('%D9%BE%D8%B1%D8%B3%D9%86%D9%84%20%D9%85%D8%B9%D8%A7%D9%88%D9%86%D8%AA%20%D9%85%DA%A9%D8%A7%D9%86%DB%8C%DA%A9')/items?$select=ID,Title,FirstName,FieldValuesAsText/MechDepartment,Evaluator/Name&$expand=FieldValuesAsText,Evaluator&$filter=Evaluator/Name eq 'i%3A0%23.w%7Csarirpey%5Cvradmin'
PersonnelAppraisal.tsx:531 Employees API Response: Objectd: results: Array(4)0: Evaluator: Name: "i:0#.w|sarirpey\\vradmin"__metadata: id: "3f4e73ea-f63e-459a-ba03-de0ef66d49f4"type: "SP.Data.UserInfoItem"[[Prototype]]: Object[[Prototype]]: ObjectFieldValuesAsText: __metadata: id: "https://apps.sarirpey.com/_api/Web/Lists(guid'ba000b34-17ee-4e64-9a0c-2e22dbfd45ef')/Items(14)/FieldValuesAsText"type: "SP.FieldStringValues"uri: "https://apps.sarirpey.com/_api/Web/Lists(guid'ba000b34-17ee-4e64-9a0c-2e22dbfd45ef')/Items(14)/FieldValuesAsText"[[Prototype]]: Object[[Prototype]]: ObjectFirstName: "نوید"ID: 14Id: 14Title: "پاکروان"__metadata: etag: "\"5\""id: "16ac3720-783c-401a-9852-5deb22401e78"type: "SP.Data.MechPersonnelListItem"uri: "https://apps.sarirpey.com/_api/Web/Lists(guid'ba000b34-17ee-4e64-9a0c-2e22dbfd45ef')/Items(14)"[[Prototype]]: Object[[Prototype]]: Object1: {__metadata: {…}, FieldValuesAsText: {…}, Evaluator: {…}, Id: 15, Title: 'دژم', …}2: {__metadata: {…}, FieldValuesAsText: {…}, Evaluator: {…}, Id: 16, Title: 'سامانی پور', …}3: {__metadata: {…}, FieldValuesAsText: {…}, Evaluator: {…}, Id: 17, Title: 'داودی نیا', …}length: 4[[Prototype]]: Array(0)[[Prototype]]: Object[[Prototype]]: Object
4PersonnelAppraisal.tsx:538 MechDepartment Value: undefined
PersonnelAppraisal.tsx:601 Selected Department Guid:
PersonnelAppraisal.tsx:662 Questions URL: https://apps.sarirpey.com/_api/web/lists/getbytitle('QuestionBank')/items?$select=ID,Title,QuestionWeight,MechDepartment
PersonnelAppraisal.tsx:677 Questions API Response: Object
PersonnelAppraisal.tsx:683 Filtered Questions: Array(0)length: 0[[Prototype]]: Array(0)at: ƒ at()concat: ƒ concat()constructor: ƒ Array()copyWithin: ƒ copyWithin()entries: ƒ entries()every: ƒ every()fill: ƒ fill()filter: ƒ filter()find: ƒ find()findIndex: ƒ findIndex()findLast: ƒ findLast()findLastIndex: ƒ findLastIndex()flat: ƒ flat()flatMap: ƒ flatMap()forEach: ƒ forEach()includes: ƒ includes()indexOf: ƒ indexOf()join: ƒ join()keys: ƒ keys()lastIndexOf: ƒ lastIndexOf()length: 0map: ƒ map()pop: ƒ pop()push: ƒ push()reduce: ƒ reduce()reduceRight: ƒ reduceRight()reverse: ƒ reverse()shift: ƒ shift()slice: ƒ slice()some: ƒ some()sort: ƒ sort()splice: ƒ splice()toLocaleString: ƒ toLocaleString()toReversed: ƒ toReversed()toSorted: ƒ toSorted()toSpliced: ƒ toSpliced()toString: ƒ toString()unshift: ƒ unshift()values: ƒ values()with: ƒ with()Symbol(Symbol.iterator): ƒ values()Symbol(Symbol.unscopables): {at: true, copyWithin: true, entries: true, fill: true, find: true, …}[[Prototype]]: Object
SPAndroidAppManifest.aspx:1

```

<https://apps.sarirpey.com/_api/web/lists/getbytitle('QuestionBank')/items?$select=ID,Title,QuestionWeight,MechDepartment&$filter=MechDepartment/Label%20eq%20%274%27>

<m:error xmlns:m="<http://schemas.microsoft.com/ado/2007/08/dataservices/metadata">>

 <div id="in-page-channel-node-id" data-channel-name="in_page_channel_oojcvz"/>
<m:code>-1, Microsoft.SharePoint.SPException</m:code>
<m:message xml:lang="en-US">The field 'MechDepartment' of type 'TaxonomyFieldType' cannot be used in the query filter expression.</m:message>
</m:error>
