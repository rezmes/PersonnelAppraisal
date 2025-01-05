
// import * as React from "react";
// import { sp } from "@pnp/sp/presets/all";
// import {
//   Dropdown,
//   IDropdownOption,
//   PrimaryButton,
//   Spinner,
//   SpinnerSize,
//   Label,
//   Dialog,
//   DialogType,
//   DialogFooter,
// } from "office-ui-fabric-react";
// import { IPersonnelAppraisalProps } from "./IPersonnelAppraisalProps";

// interface IEmployeeOption extends IDropdownOption {
//   department: string;
//   departmentGuid: string;
// }

// interface IAppraisalFormState {
//   employees: IEmployeeOption[];
//   selectedEmployee: string | number | undefined;
//   questions: { id: number; text: string; weight: number }[];
//   scores: { [questionId: number]: number };
//   isLoading: boolean;
//   errorMessage: string | null;
//   isDialogHidden: boolean;
// }

// import "core-js/es6/array";

// export default class PersonnelAppraisal extends React.Component<
//   IPersonnelAppraisalProps,
//   IAppraisalFormState
// > {
//   constructor(props: IPersonnelAppraisalProps) {
//     super(props);

//     this.state = {
//       employees: [],
//       selectedEmployee: undefined,
//       questions: [],
//       scores: {},
//       isLoading: false,
//       errorMessage: null,
//       isDialogHidden: true,
//     };
//   }

//   componentDidMount(): void {
//     sp.setup({
//       spfxContext: this.props.context,
//     });

//     this.loadEmployees();
//   }

//   private async loadEmployees(): Promise<void> {
//     try {
//       this.setState({ isLoading: true });
//       const currentUser = await sp.web.currentUser.get();
//       console.log("Current user:", currentUser);

//       const employees = await sp.web.lists
//         .getByTitle("پرسنل معاونت مکانیک")
//         .items.select(
//           "ID",
//           "Title",
//           "FirstName",
//           "Department",
//           "Evaluator/Name",
//           "MechDepartment"
//         )
//         .expand("Evaluator")
//         .filter(`Evaluator/Name eq '${currentUser.LoginName}'`)
//         .get();

//       console.log("Filtered employees:", employees);

//       const evaluationPeriod = "Q1-2024";

//       const evaluatedEmployees = await sp.web.lists
//         .getByTitle("EvaluationResults")
//         .items.select("EmployeeID/ID", "EvaluationPeriod")
//         .expand("EmployeeID") // Expand the EmployeeID lookup field
//         .filter(`EvaluationPeriod eq '${evaluationPeriod}'`)
//         .get();

//       console.log("Evaluated employees:", evaluatedEmployees);

//       const evaluatedEmployeeIds = evaluatedEmployees.map(
//         (evalItem) => evalItem.EmployeeID.ID
//       );

//       const employeeOptions: IEmployeeOption[] = employees
//         .filter((emp) => evaluatedEmployeeIds.indexOf(emp.ID) === -1)
//         .map((emp) => {
//           let departmentText = "";
//           let departmentTermGuid = "";
//           if (emp.MechDepartment && emp.MechDepartment.Label) {
//             departmentText = emp.MechDepartment.Label;
//             departmentTermGuid = emp.MechDepartment.TermGuid;
//           }

//           return {
//             key: emp.ID,
//             text: `${emp.FirstName} ${emp.Title}`,
//             department: departmentText,
//             departmentGuid: departmentTermGuid,
//           };
//         });

//       console.log("Employee options:", employeeOptions);
//       this.setState({ employees: employeeOptions, isLoading: false });
//     } catch (error) {
//       this.setState({
//         errorMessage: "Error loading employees.",
//         isLoading: false,
//       });
//       console.error(error);
//     }
//   }

//   private handleEmployeeChange = (
//     option?: IDropdownOption
//   ): void => {
//     if (option) {
//       let selectedEmployee: IEmployeeOption | undefined = undefined;
//       for (let i = 0; i < this.state.employees.length; i++) {
//         if (this.state.employees[i].key === option.key) {
//           selectedEmployee = this.state.employees[i];
//           break;
//         }
//       }

//       const selectedDepartmentGuid = selectedEmployee
//         ? selectedEmployee.departmentGuid
//         : "";

//       this.setState({ selectedEmployee: option.key as string }, () => {
//         this.loadQuestions(selectedDepartmentGuid);
//       });
//     }
//   };

//   private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
//     try {
//       this.setState({ isLoading: true });

//       const questions = await sp.web.lists
//         .getByTitle("QuestionBank")
//         .items.select("ID", "Title", "QuestionWeight", "Department")
//         .get();

//       const filteredQuestions = questions.filter((q) =>
//         q.Department.TermGuid === selectedDepartmentGuid
//       );

//       this.setState({
//         questions: filteredQuestions.map((q) => ({
//           id: q.ID,
//           text: q.Title,
//           weight: q.QuestionWeight,
//         })),
//         scores: {},
//         isLoading: false,
//       });
//     } catch (error) {
//       this.setState({
//         errorMessage: "Error loading questions.",
//         isLoading: false,
//       });
//       console.error("Error fetching questions:", error);
//     }
//   }

//   private handleScoreChange = (questionId: number, score: number): void => {
//     this.setState((prevState) => ({
//       scores: {
//         ...prevState.scores,
//         [questionId]: score,
//       },
//     }));
//   };

//   private handleSubmit: () => Promise<void> = async (): Promise<void> => {
//     const { selectedEmployee, scores, questions } = this.state;

//     if (!selectedEmployee) {
//       this.setState({ errorMessage: "Please select an employee." });
//       return;
//     }

//     if (Object.keys(scores).length !== questions.length) {
//       this.setState({ errorMessage: "Please rate all questions." });
//       return;
//     }

//     try {
//       this.setState({ isLoading: true, errorMessage: null });

//       const batch = sp.web.createBatch();
//       const evaluationPeriod = "Q1-2024";

//       questions.forEach((question) => {
//         const weightedScore = (scores[question.id] / 5) * question.weight;

//         const item = {
//           EmployeeIDId: selectedEmployee,
//           QuestionDescription: question.text,
//           Score: scores[question.id],
//           WeightedScore: weightedScore,
//           EvaluationPeriod: evaluationPeriod,
//         };

//         sp.web.lists.getByTitle("EvaluationResults").items.inBatch(batch).add(item);
//       });

//       await batch.execute();

//       this.setState({ isLoading: false });
//       alert("Evaluation submitted successfully.");

//       this.loadEmployees();
//     } catch (error) {
//       this.setState({
//         errorMessage: "Error submitting evaluation.",
//         isLoading: false,
//       });
//       console.error("Error submitting evaluation:", error);
//     }
//   };

//   private closeDialog = (): void => {
//     this.setState({ isDialogHidden: true });
//   };

//   render(): React.ReactElement<any> {
//     const { employees, selectedEmployee, questions, scores, isLoading, errorMessage, isDialogHidden } = this.state;

//     return (
//       <div>
//         <h3>{this.props.description}</h3>
//         {isLoading && <Spinner size={SpinnerSize.large} label="Loading..." />}
//         {errorMessage && <Label style={{ color: "red" }}>{errorMessage}</Label>}
//         <Dropdown
//           placeHolder="Choose an employee"
//           options={employees}
//           onChanged={this.handleEmployeeChange}
//           selectedKey={selectedEmployee}
//         />
//         {questions.length > 0 && (
//           <div>
//             <table>
//               <thead>
//                 <tr>
//                   <th>Question</th>
//                   <th>Score</th>
//                 </tr>
//               </thead>
//               <tbody>
//                 {questions.map((question) => (
//                   <tr key={question.id}>
//                     <td>{question.text}</td>
//                     <td>
//                       {[1, 2, 3, 4, 5].map((score) => (
//                         <label key={score}>
//                           <input
//                             type="radio"
//                             name={`question-${question.id}`}
//                             value={score}
//                             checked={scores[question.id] === score}
//                             onChange={() =>
//                               this.handleScoreChange(question.id, score)
//                             }
//                           />
//                           {score}
//                         </label>
//                       ))}
//                     </td>
//                   </tr>
//                 ))}
//               </tbody>
//             </table>
//             <PrimaryButton text="Submit" onClick={this.handleSubmit} />
//           </div>
//         )}
//         <Dialog
//           hidden={isDialogHidden}
//           onDismiss={this.closeDialog}
//           dialogContentProps={{
//             type: DialogType.normal,
//             title: 'Some Title',
//             subText: 'Some subtitle',
//             className: 'some-class',
//           }}
//           modalProps={{
//             isBlocking: false,
//             containerClassName: 'some-container-class',
//           }}
//         >
//           <DialogFooter>
//             <PrimaryButton onClick={this.closeDialog} text="OK" />
//           </DialogFooter>
//         </Dialog>
//       </div>
//     );
//   }
// }

// PersonnelAppraisal.tsx

// PersonnelAppraisal.tsx

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
    };
  }

  componentDidMount(): void {
    sp.setup({
      spfxContext: this.props.context,
    });

    this.loadEmployees();
  }

  // Load employees from SharePoint list
  private async loadEmployees(): Promise<void> {
    try {
      this.setState({ isLoading: true });
      const currentUser = await sp.web.currentUser.get();

      const employees = await sp.web.lists
        .getByTitle("پرسنل معاونت مکانیک")
        .items.select("ID", "Title", "FirstName", "Department", "Evaluator/Name", "MechDepartment")
        .expand("Evaluator")
        .filter(`Evaluator/Name eq '${currentUser.LoginName}'`)
        .get();

      const evaluationPeriod = "Q1-2024";

      const evaluatedEmployees = await sp.web.lists
        .getByTitle("EvaluationResults")
        .items.select("EmployeeID/ID", "EvaluationPeriod")
        .expand("EmployeeID")
        .filter(`EvaluationPeriod eq '${evaluationPeriod}'`)
        .get();

      const evaluatedEmployeeIds = evaluatedEmployees.map(evalItem => evalItem.EmployeeID.ID);

      const employeeOptions: IEmployeeOption[] = employees
        .filter(emp => evaluatedEmployeeIds.indexOf(emp.ID) === -1)
        .map(emp => {
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

      const selectedDepartmentGuid = selectedEmployee ? selectedEmployee.departmentGuid : "";

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

      const filteredQuestions = questions.filter(q => q.Department.TermGuid === selectedDepartmentGuid);

      this.setState({
        questions: filteredQuestions.map(q => ({
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
    this.setState(prevState => ({
      scores: {
        ...prevState.scores,
        [questionId]: score,
      },
    }));
  };

  // Handle form submission to save evaluation results
  private handleSubmit: () => Promise<void> = async (): Promise<void> => {
    const { selectedEmployee, scores, questions } = this.state;

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

      const batch = sp.web.createBatch();
      const evaluationPeriod = "Q1-2024";

      questions.forEach(question => {
        const weightedScore = (scores[question.id] / 5) * question.weight;

        const item = {
          EmployeeIDId: selectedEmployee,
          QuestionDescription: question.text,
          Score: scores[question.id],
          WeightedScore: weightedScore,
          EvaluationPeriod: evaluationPeriod,
        };

        sp.web.lists.getByTitle("EvaluationResults").items.inBatch(batch).add(item);
      });

      await batch.execute();

      this.setState({ isLoading: false });
      alert("Evaluation submitted successfully.");

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
    const { employees, selectedEmployee, questions, scores, isLoading, errorMessage, isDialogHidden } = this.state;

    return (
      <div>
        <h3>{this.props.description}</h3>
        {isLoading && <Spinner size={SpinnerSize.large} label="Loading..." />}
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
        <PrimaryButton text="Submit" onClick={this.handleSubmit} />
        <Dialog
          hidden={isDialogHidden}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Some Title',
            subText: 'Some subtitle',
            className: 'some-class',
          }}
          modalProps={{
            isBlocking: false,
            containerClassName: 'some-container-class',
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
