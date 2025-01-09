```tsx
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

interface IEmployeeOption extends IDropdownOption {
  department: string;
  departmentGuid: string;
}

interface IAppraisalFormState {
  employees: IEmployeeOption[];
  selectedEmployee: string | number | undefined;
  questions: { id: number; text: string; weight: number }[];
  scores: { [questionId: number]: number };
  isLoading: boolean;
  errorMessage: string | null;
  isDialogHidden: boolean;
  evaluationPeriod: string;
}

import "core-js/es6/array";

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

      // Query using MechDepartment/TermGuid
      const employeesUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=ID,Title,FirstName,MechDepartment/TermGuid,Evaluator/Name&$expand=MechDepartment,Evaluator&$filter=Evaluator/Name eq '${encodedLoginName}'`;
      console.log("Employees URL:", employeesUrl);

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
        (emp: any) => ({
          key: emp.ID,
          text: `${emp.FirstName} ${emp.Title}`,
          department: emp.MechDepartment?.Label || "Unknown Department",
          departmentGuid: emp.MechDepartment?.TermGuid || "Unknown Department",
        })
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
      console.log("Selected Employee:", selectedEmployee);
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
      if (!selectedDepartmentGuid) {
        throw new Error("selectedDepartmentGuid is empty");
      }

      this.setState({ isLoading: true });

      const questionBankListName = encodeURIComponent(
        this.props.questionBankListName
      );
      const evaluationPeriod = this.state.evaluationPeriod;

      const questionsUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${questionBankListName}')/items?$select=ID,Title,QuestionWeight,MechDepartment/TermGuid,MechDepartment/Label&$expand=MechDepartment&$filter=MechDepartment/TermGuid eq '${selectedDepartmentGuid}'`;
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

      this.setState({
        questions: questions.d.results.map((q: any) => ({
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
            className: "some-class",
          }}
          modalProps={{
            isBlocking: false,
            containerClassName: "some-container-class",
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




```js
Selected Employee: {key: 14, text: 'نوید پاکروان', department: undefined, departmentGuid: ''}department: undefineddepartmentGuid: ""key: 14text: "نوید پاکروان"[[Prototype]]: Object
PersonnelAppraisal.tsx:245 Selected Department Guid: 
PersonnelAppraisal.tsx:381  Error fetching questions: Error: selectedDepartmentGuid is empty
    at PersonnelAppraisal.<anonymous> (PersonnelAppraisal.tsx:304:15)
    at step (personnel-appraisal-web-part.js:11486:23)
    at Object.next (personnel-appraisal-web-part.js:11467:53)
    at personnel-appraisal-web-part.js:11461:71
    at new Promise (<anonymous>)
    at __awaiter (personnel-appraisal-web-part.js:11457:12)
    at PersonnelAppraisal.loadQuestions (PersonnelAppraisal.tsx:301:62)
    at PersonnelAppraisal.<anonymous> (PersonnelAppraisal.tsx:250:16)
    at e.notifyAll (sp-webpart-workbench-assembly.js?uniqueId=hB47W:214:32505)
    at o.close (sp-webpart-workbench-assembly.js?uniqueId=hB47W:214:30964)
```

```
[21:03:06] Error - typescript - src\webparts\personnelAppraisal\components\PersonnelAppraisal.tsx(620,41): error TS1109: Expression expected.
[21:03:06] Error - typescript - src\webparts\personnelAppraisal\components\PersonnelAppraisal.tsx(620,71): error TS1005: ':' expected.
[21:03:06] Error - typescript - src\webparts\personnelAppraisal\components\PersonnelAppraisal.tsx(621,45): error TS1109: Expression expected.
[21:03:06] Error - typescript - src\webparts\personnelAppraisal\components\PersonnelAppraisal.tsx(621,78): error TS1005: ':' expected.
[21:03:06] Warning - tslint - src\webparts\personnelAppraisal\components\QuestionTable.tsx(16,3): error member-access: The class method 'render' must be marked either 'private', 'public', or 'protected'
[21:03:06] Warning - tslint - src\webparts\personnelAppraisal\components\QuestionTable.tsx(1,24): error quotemark: " should be '
[21:03:06] Warning - tslint - src\webparts\personnelAppraisal\components\QuestionTable.tsx(2,8): error quotemark: " should be '
[21:03:06] Warning - tslint - src\webparts\personnelAppraisal\components\QuestionTable.tsx(33,28): error quotemark: " should be '
[21:03:06] Warning - tslint - src\webparts\personnelAppraisal\components\QuestionTable.tsx(16,3): error typedef: expected call-signature: 'render' to have a typedef
[21:03:06] Warning - tslint - src\webparts\personnelAppraisal\components\QuestionTable.tsx(32,21): error react-a11y-role-has-required-aria-props: Tag 'input' has implicit role 'radio'. It requires aria-* attributes: aria-checked that are missing in the element. A reference to role definitions can be found at https://www.w3.org/TR/wai-aria/roles#role_definitions.
[21:03:06] Finished subtask 'tslint' after 5.81 s
[21:03:06] Error - 'typescript' sub task errored after 3.36 s 
 TypeScript error(s) occurred.
  Request: '/temp/manifests.js'
  Request: '/temp/manifests.js'
```
<!-- ------------------------------------------------------------ -->


```tsx
...
  componentDidMount(): void {
    this.loadEmployees();
  }

  private handlePeriodLoaded = (period: string): void => {
    this.setState({ evaluationPeriod: period });
  };

  ...
    private async loadEmployees(): Promise<void> {
    try {
      this.setState({ isLoading: true });
      const response = await fetch(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/web/currentUser`,
        {
          headers: { Accept: "application/json;odata=verbose" },
        }
      );
      const currentUser = await response.json();
      console.log("Current User:", currentUser);

      const encodedLoginName = encodeURIComponent(currentUser.d.LoginName);
      const listName = encodeURIComponent(this.props.employeeListName);

      const employeesUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=ID,Title,FirstName,FieldValuesAsText/MechDepartment,Evaluator/Name&$expand=FieldValuesAsText,Evaluator&$filter=Evaluator/Name eq '${encodedLoginName}'`;
      console.log("Employees URL:", employeesUrl);

      const employeesResponse = await fetch(employeesUrl, {
        headers: { Accept: "application/json;odata=verbose" },
      });

      if (!employeesResponse.ok) {
        throw new Error(
          `Error fetching employees: ${employeesResponse.statusText}`
        );
      }

      const employees = await employeesResponse.json();
      console.log("Employees API Response:", employees);

      const employeeOptions: IEmployeeOption[] = employees.d.results.map(
        (emp: any) => ({
          key: emp.ID,
          text: `${emp.FirstName} ${emp.Title}`,
          department:
            emp.MechDepartment && emp.MechDepartment.Label
              ? emp.MechDepartment.Label
              : "Unknown Department",
          departmentGuid:
            emp.MechDepartment && emp.MechDepartment.TermGuid
              ? emp.MechDepartment.TermGuid
              : "",
        })
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
  ...
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
      console.log("Selected Employee:", selectedEmployee);
      console.log("Selected Department Guid:", selectedDepartmentGuid);

      this.setState(
        { selectedEmployee: option.key as string, questions: [] },
        () => {
          this.loadQuestions(selectedEmployee.department);
        }
      );
    }
  };
...
  private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
    try {
      if (!selectedDepartmentGuid) {
        console.error(
          "selectedDepartmentGuid is empty. Ensure the mapping is correct."
        );
        throw new Error("selectedDepartmentGuid is empty");
      }

      this.setState({ isLoading: true });

      const questionBankListName = encodeURIComponent(
        this.props.questionBankListName
      );
      const evaluationPeriod = this.state.evaluationPeriod;

      // Query to fetch questions based on MechDepartment
      const questionsUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${questionBankListName}')/items?$select=ID,Title,QuestionWeight,MechDepartment/TermGuid,MechDepartment/Label&$expand=MechDepartment&$filter=MechDepartment/TermGuid eq '${selectedDepartmentGuid}'`;
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

      // Fetch evaluated items based on evaluation period and PersonnelCode
      const evaluationResultsListName = encodeURIComponent(
        this.props.evaluationResultsListName
      );
      const evaluatedItemsUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${evaluationResultsListName}')/items?$select=ID,EvaluationPeriod,PersonnelCode`;
      console.log("Evaluated Items URL:", evaluatedItemsUrl);

      const evaluatedItemsResponse = await fetch(evaluatedItemsUrl, {
        headers: {
          Accept: "application/json;odata=verbose",
        },
      });

      if (!evaluatedItemsResponse.ok) {
        throw new Error(
          `Error fetching evaluated items: ${evaluatedItemsResponse.statusText}`
        );
      }

      const evaluatedItems = await evaluatedItemsResponse.json();
      console.log("Evaluated Items API Response:", evaluatedItems);

      const evaluatedCombinations = evaluatedItems.d.results.map(
        (item: any) => `${item.EvaluationPeriod}${item.PersonnelCode}`
      );

      // Filter questions to ensure they do not exist in the evaluated combinations
      const filteredQuestions = questions.d.results.filter(
        (q: any) =>
          !evaluatedCombinations.includes(
            `${evaluationPeriod}${q.PersonnelCode}`
          )
      );

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
```
and we have this error:

```js
Employees URL: https://apps.sarirpey.com/_api/web/lists/getbytitle('%D9%BE%D8%B1%D8%B3%D9%86%D9%84%20%D9%85%D8%B9%D8%A7%D9%88%D9%86%D8%AA%20%D9%85%DA%A9%D8%A7%D9%86%DB%8C%DA%A9')/items?$select=ID,Title,FirstName,FieldValuesAsText/MechDepartment,Evaluator/Name&$expand=FieldValuesAsText,Evaluator&$filter=Evaluator/Name eq 'i%3A0%23.w%7Csarirpey%5Cvradmin'
PersonnelAppraisal.tsx:149 Employees API Response: {d: {…}}
PersonnelAppraisal.tsx:243 Selected Employee: {key: 14, text: 'نوید پاکروان', department: 'Unknown Department', departmentGuid: ''}department: "Unknown Department"departmentGuid: ""key: 14text: "نوید پاکروان"[[Prototype]]: Object
PersonnelAppraisal.tsx:244 Selected Department Guid: 
PersonnelAppraisal.tsx:318 Questions URL: https://apps.sarirpey.com/_api/web/lists/getbytitle('QuestionBank')/items?$select=ID,Title,QuestionWeight,MechDepartment/TermGuid,MechDepartment/Label&$expand=MechDepartment&$filter=MechDepartment/TermGuid eq 'Unknown Department'
PersonnelAppraisal.tsx:320 
        
        
        GET https://apps.sarirpey.com/_api/web/lists/getbytitle('QuestionBank')/items?$select=ID,Title,QuestionWeight,MechDepartment/TermGuid,MechDepartment/Label&$expand=MechDepartment&$filter=MechDepartment/TermGuid%20eq%20%27Unknown%20Department%27 400 (Bad Request)
(anonymous) @ PersonnelAppraisal.tsx:320
step @ personnel-appraisal-web-part.js:11486
(anonymous) @ personnel-appraisal-web-part.js:11467
(anonymous) @ personnel-appraisal-web-part.js:11461
__awaiter @ personnel-appraisal-web-part.js:11457
PersonnelAppraisal.loadQuestions @ PersonnelAppraisal.tsx:300
(anonymous) @ PersonnelAppraisal.tsx:249
(anonymous) @ sp-webpart-workbench-assembly.js?uniqueId=hB47W:214
close @ sp-webpart-workbench-assembly.js?uniqueId=hB47W:214
closeAll @ sp-webpart-workbench-assembly.js?uniqueId=hB47W:214
perform @ sp-webpart-workbench-assembly.js?uniqueId=hB47W:214
perform @ sp-webpart-workbench-assembly.js?uniqueId=hB47W:214
k @ sp-webpart-workbench-assembly.js?uniqueId=hB47W:214
closeAll @ sp-webpart-workbench-assembly.js?uniqueId=hB47W:214
perform @ sp-webpart-workbench-assembly.js?uniqueId=hB47W:214
batchedUpdates @ sp-webpart-workbench-assembly.js?uniqueId=hB47W:228
i @ sp-webpart-workbench-assembly.js?uniqueId=hB47W:214
dispatchEvent @ sp-webpart-workbench-assembly.js?uniqueId=hB47W:228
PersonnelAppraisal.tsx:383  Error fetching questions: Error: Error fetching questions: Bad Request
    at PersonnelAppraisal.<anonymous> (PersonnelAppraisal.tsx:327:15)
    at step (personnel-appraisal-web-part.js:11486:23)
    at Object.next (personnel-appraisal-web-part.js:11467:53)
    at fulfilled (personnel-appraisal-web-part.js:11458:58)
```
