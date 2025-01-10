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
  evaluatedEmployees: { employeeId: number; evaluationPeriod: string }[];
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
      evaluatedEmployees: [],
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
    this.loadEvaluationResults(); // Fetch evaluation results first
    this.loadEmployees();
  }

  private async loadEvaluationResults(): Promise<void> {
    try {
      const response = await fetch(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.evaluationResultsListName}')/items?$select=EmployeeIDId,EvaluationPeriod`,
        {
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }
      );

      if (!response.ok) {
        throw new Error(
          `Error fetching evaluation results: ${response.statusText}`
        );
      }

      const data = await response.json();

      const evaluatedEmployees = data.d.results.map((result: any) => ({
        employeeId: result.EmployeeIDId,
        evaluationPeriod: result.EvaluationPeriod,
      }));

      this.setState({ evaluatedEmployees: evaluatedEmployees });
    } catch (error) {
      this.setState({
        errorMessage: `Error loading evaluation results: ${error.message}`,
      });
      console.error("Error loading evaluation results:", error);
    }
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

      const encodedLoginName = encodeURIComponent(currentUser.d.LoginName);
      const listName = encodeURIComponent(this.props.employeeListName);

      const employeesUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=ID,Title,FirstName,MechDepartment,Evaluator/Name&$expand=Evaluator&$filter=Evaluator/Name eq '${encodedLoginName}'`;

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

      const evaluatedEmployees = this.state.evaluatedEmployees;

      const employeeOptions: IEmployeeOption[] = employees.d.results
        .filter((emp: any) => {
          return !evaluatedEmployees.some(
            (evaluated: any) =>
              evaluated.employeeId === emp.ID &&
              evaluated.evaluationPeriod === this.state.evaluationPeriod
          );
        })
        .map((emp: any) => ({
          key: emp.ID,
          text: `${emp.FirstName} ${emp.Title}`,
          department: emp.MechDepartment ? emp.MechDepartment.Label : "",
          departmentGuid: emp.MechDepartment ? emp.MechDepartment.TermGuid : "",
        }));

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

      // Client-Side Filtering based on selectedDepartmentGuid
      const filteredQuestions = questions.d.results.filter((q: any) => {
        const department = q.MechDepartment ? q.MechDepartment.TermGuid : "";
        return department === selectedDepartmentGuid;
      });

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

  // private handleSubmit: () => Promise<void> = async (): Promise<void> => {
  //   const { selectedEmployee, scores, questions, evaluationPeriod } =
  //     this.state;

  //   if (!selectedEmployee) {
  //     this.setState({ errorMessage: "Please select an employee." });
  //     return;
  //   }

  //   if (Object.keys(scores).length !== questions.length) {
  //     this.setState({ errorMessage: "Please rate all questions." });
  //     return;
  //   }

  //   try {
  //     this.setState({ isLoading: true, errorMessage: null });

  //     const digestResponse = await fetch(
  //       `${this.props.context.pageContext.web.absoluteUrl}/_api/contextinfo`,
  //       {
  //         method: "POST",
  //         headers: {
  //           Accept: "application/json;odata=verbose",
  //         },
  //       }
  //     );

  //     const digestData = await digestResponse.json();
  //     const requestDigest =
  //       digestData.d.GetContextWebInformation.FormDigestValue;

  //     const batchOperations = questions.map((question) => {
  //       const weightedScore = (scores[question.id] / 5) * question.weight;

  //       const item = {
  //         __metadata: { type: "SP.Data.EvaluationResultsListItem" },
  //         EmployeeIDId: selectedEmployee,
  //         QuestionDescription: question.text,
  //         Score: scores[question.id],
  //         WeightedScore: weightedScore,
  //         EvaluationPeriod: evaluationPeriod,
  //       };

  //       return fetch(
  //         `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.evaluationResultsListName}')/items`,
  //         {
  //           method: "POST",
  //           headers: {
  //             Accept: "application/json;odata=verbose",
  //             "Content-Type": "application/json;odata=verbose",
  //             "X-RequestDigest": requestDigest,
  //           },
  //           body: JSON.stringify(item),
  //         }
  //       ).then((response) => {
  //         if (!response.ok) {
  //           return response.text().then((text) => {
  //             throw new Error(text);
  //           });
  //         }
  //         return response.json();
  //       });
  //     });

  //     await Promise.all(batchOperations);

  //     this.setState({ isLoading: false, questions: [] });
  //     alert("ارزیابی با موفقیت ثبت شد");

  //     this.loadEmployees();
  //   } catch (error) {
  //     this.setState({
  //       errorMessage: `Error submitting evaluation: ${error.message}`,
  //       isLoading: false,
  //     });
  //     console.error("Error submitting evaluation:", error);
  //   }
  // };

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

      const digestResponse = await fetch(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/contextinfo`,
        {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }
      );

      const digestData = await digestResponse.json();
      const requestDigest =
        digestData.d.GetContextWebInformation.FormDigestValue;

      const batchOperations = questions.map((question) => {
        const weightedScore = (scores[question.id] / 5) * question.weight;

        const item = {
          __metadata: { type: "SP.Data.EvaluationResultsListItem" },
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
              "X-RequestDigest": requestDigest,
            },
            body: JSON.stringify(item),
          }
        ).then((response) => {
          if (!response.ok) {
            return response.text().then((text) => {
              throw new Error(text);
            });
          }
          return response.json();
        });
      });

      await Promise.all(batchOperations);

      this.setState({ isLoading: false, questions: [] });
      alert("ارزیابی با موفقیت ثبت شد");

      await this.loadEvaluationResults(); // Load evaluated employees first
      await this.loadEmployees(); // Then reload the employees list
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
