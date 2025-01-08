import * as React from "react";
import { sp } from "@pnp/sp/presets/all";
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
import EvaluationPeriod from "./EvaluationPeriod"; // Import the new component
import "./PersonnelAppraisal.module.scss"; // Import your styles

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
    sp.setup({
      spfxContext: this.props.context,
    });

    this.loadEmployees();
  }

  private handlePeriodLoaded = (period: string): void => {
    this.setState({ evaluationPeriod: period });
  };

  // Load employees from SharePoint list
  private async loadEmployees(): Promise<void> {
    try {
      this.setState({ isLoading: true });
      const currentUser = await sp.web.currentUser.get();

      const employees = await sp.web.lists
        .getByTitle("پرسنل معاونت مکانیک")
        .items.select(
          "ID",
          "Title",
          "FirstName",
          "Department",
          "Evaluator/Name",
          "MechDepartment"
        )
        .expand("Evaluator")
        .filter(`Evaluator/Name eq '${currentUser.LoginName}'`)
        .get();

      const evaluatedEmployees = await sp.web.lists
        .getByTitle("EvaluationResults")
        .items.select("EmployeeID/ID", "EvaluationPeriod")
        .expand("EmployeeID")
        .filter(`EvaluationPeriod eq '${this.state.evaluationPeriod}'`)
        .get();

      const evaluatedEmployeeIds = evaluatedEmployees.map(
        (evalItem) => evalItem.EmployeeID.ID
      );

      const employeeOptions: IEmployeeOption[] = employees
        .filter((emp) => evaluatedEmployeeIds.indexOf(emp.ID) === -1)
        .map((emp) => {
          let departmentText = "";
          let departmentTermGuid = "";
          if (emp.MechDepartment && emp.MechDepartment.Label) {
            departmentText = emp.MechDepartment.Label;
            departmentTermGuid = emp.MechDepartment.TermGuid;
          }

          return {
            key: emp.ID,
            text: `${emp.FirstName} ${emp.Title}`,
            department: departmentText,
            departmentGuid: departmentTermGuid,
          };
        });

      this.setState({ employees: employeeOptions, isLoading: false });
    } catch (error) {
      this.setState({
        errorMessage: "Error loading employees.",
        isLoading: false,
      });
      console.error(error);
    }
  }

  // Handle employee selection change
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

      this.setState({ selectedEmployee: option.key as string }, () => {
        this.loadQuestions(selectedDepartmentGuid);
      });
    }
  };

  // Load questions from SharePoint list based on selected department
  private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
    try {
      this.setState({ isLoading: true });

      const questions = await sp.web.lists
        .getByTitle("QuestionBank")
        .items.select("ID", "Title", "QuestionWeight", "Department")
        .get();

      const filteredQuestions = questions.filter(
        (q) => q.Department.TermGuid === selectedDepartmentGuid
      );

      this.setState({
        questions: filteredQuestions.map((q) => ({
          id: q.ID,
          text: q.Title,
          weight: q.QuestionWeight,
        })),
        scores: {},
        isLoading: false,
      });
    } catch (error) {
      this.setState({
        errorMessage: "Error loading questions.",
        isLoading: false,
      });
      console.error("Error fetching questions:", error);
    }
  }

  // Handle score change for questions
  private handleScoreChange = (questionId: number, score: number): void => {
    this.setState((prevState) => ({
      scores: {
        ...prevState.scores,
        [questionId]: score,
      },
    }));
  };

  // Handle form submission to save evaluation results
  private handleSubmit: () => Promise<void> = async (): Promise<void> => {
    const { selectedEmployee, scores, questions, evaluationPeriod } =
      this.state;

    if (!selectedEmployee) {
      this.setState({ errorMessage: ".لطفا یک نفر را انتخاب فرمایید" });
      return;
    }

    if (Object.keys(scores).length !== questions.length) {
      this.setState({ errorMessage: ".لطفا به همه ی سوالات پاسخ دهید" });
      return;
    }

    try {
      this.setState({ isLoading: true, errorMessage: null });

      const batch = sp.web.createBatch();

      questions.forEach((question) => {
        const weightedScore = (scores[question.id] / 5) * question.weight;

        const item = {
          EmployeeIDId: selectedEmployee,
          QuestionDescription: question.text,
          Score: scores[question.id],
          WeightedScore: weightedScore,
          EvaluationPeriod: evaluationPeriod,
        };

        sp.web.lists
          .getByTitle("EvaluationResults")
          .items.inBatch(batch)
          .add(item);
      });

      await batch.execute();

      this.setState({ isLoading: false });
      alert("ارزیابی با موفقیت ثبت شد");

      this.loadEmployees();
    } catch (error) {
      this.setState({
        errorMessage: "Error submitting evaluation.",
        isLoading: false,
      });
      console.error("Error submitting evaluation:", error);
    }
  };

  // Close the dialog
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
        <EvaluationPeriod
          spfxContext={this.props.context}
          onPeriodLoaded={this.handlePeriodLoaded}
        />
        <h3>ارزیابی عملکرد کارکنان</h3>

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
