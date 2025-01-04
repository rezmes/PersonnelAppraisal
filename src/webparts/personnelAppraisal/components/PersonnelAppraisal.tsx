// import * as React from "react";
// import { sp } from "@pnp/sp/presets/all";
// import {
//   Dropdown,
//   IDropdownOption,
//   PrimaryButton,
//   Spinner,
//   SpinnerSize,
//   Label,
// } from "office-ui-fabric-react";
// import { IPersonnelAppraisalProps } from "./IPersonnelAppraisalProps";
// interface IEmployeeOption extends IDropdownOption {
//   department: string;
//   departmentGuid: string; // Add the departmentGuid property
// }

// interface IAppraisalFormState {
//   employees: IDropdownOption[];
//   selectedEmployee: string | undefined;
//   questions: { id: number; text: string; weight: number }[];
//   scores: { [questionId: number]: number };
//   isLoading: boolean;
//   errorMessage: string | null;
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
//     };
//   }

//   componentDidMount(): void {
//     sp.setup({
//       spfxContext: this.props.context, // Use the context
//     });
//     this.loadEmployees();
//   }

//   private async loadEmployees(): Promise<void> {
//     try {
//       this.setState({ isLoading: true });
//       const currentUser = await sp.web.currentUser.get();
//       console.log("this is current user:", currentUser);

//       const employees = await sp.web.lists
//         .getByTitle("پرسنل معاونت مکانیک") // Replace with your actual list title
//         .items.select(
//           "ID",
//           "Title",
//           "FirstName",
//           "Department",
//           "Evaluator/Name",
//           "MechDepartment" // Include the managed metadata field
//         )
//         .expand("Evaluator")
//         .filter(`Evaluator/Name eq '${currentUser.LoginName}'`)
//         .get();

//       console.log("Filtered employees:", employees);

//       const employeeOptions: IEmployeeOption[] = employees.map((emp) => {
//         let departmentText = "";
//         let departmentTermGuid = "";
//         if (emp.MechDepartment && emp.MechDepartment.Label) {
//           departmentText = emp.MechDepartment.Label;
//           departmentTermGuid = emp.MechDepartment.TermGuid;
//         }

//         return {
//           key: emp.ID,
//           text: `${emp.FirstName} ${emp.Title}`,
//           department: departmentText,
//           departmentGuid: departmentTermGuid, // Store the GUID term
//         };
//       });

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
//     event: React.FormEvent<HTMLDivElement>,
//     option?: IEmployeeOption // Use the extended type
//   ): void => {
//     if (option) {
//       const filteredEmployees = this.state.employees.filter(
//         (emp) => emp.key === option.key
//       );
//       const selectedEmployee =
//         filteredEmployees.length > 0 ? filteredEmployees[0] : null;

//       let selectedDepartmentGuid = "";
//       if (selectedEmployee && "departmentGuid" in selectedEmployee) {
//         selectedDepartmentGuid = (selectedEmployee as IEmployeeOption)
//           .departmentGuid;
//       }

//       this.setState(
//         { selectedEmployee: option.key as string },
//         () => this.loadQuestions(selectedDepartmentGuid) // Pass the GUID term here
//       );
//     }
//   };

//   private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
//     if (!this.state.selectedEmployee || !selectedDepartmentGuid) return;

//     try {
//       this.setState({ isLoading: true });

//       const questions = await sp.web.lists
//         .getByTitle("QuestionBank")
//         .items.filter(`Department/TermGuid eq '${selectedDepartmentGuid}'`)
//         .get();

//       this.setState({
//         questions: questions.map((q) => ({
//           id: q.ID,
//           text: q.Title,
//           weight: q.Weight,
//         })),
//         scores: {},
//         isLoading: false,
//       });
//     } catch (error) {
//       this.setState({
//         errorMessage: "Error loading questions.",
//         isLoading: false,
//       });
//       console.error(error);
//     }
//   }

//   // private async loadEmployees(): Promise<void> {
//   //   try {
//   //     this.setState({ isLoading: true });
//   //     const currentUser = await sp.web.currentUser.get();
//   //     console.log("this is current user:", currentUser);

//   //     const employees = await sp.web.lists
//   //       .getByTitle("پرسنل معاونت مکانیک") // Replace with your actual list title
//   //       .items.select(
//   //         "ID",
//   //         "Title",
//   //         "FirstName",
//   //         "Department",
//   //         "Evaluator/Name",
//   //         "MechDepartment" // Include the managed metadata field
//   //       )
//   //       .expand("Evaluator")
//   //       .filter(`Evaluator/Name eq '${currentUser.LoginName}'`)
//   //       .get();

//   //     console.log("Filtered employees:", employees);

//   //     const employeeOptions: IEmployeeOption[] = employees.map((emp) => {
//   //       let departmentText = "";
//   //       let departmentTermGuid = "";
//   //       if (emp.MechDepartment && emp.MechDepartment.Label) {
//   //         departmentText = emp.MechDepartment.Label;
//   //         departmentTermGuid = emp.MechDepartment.TermGuid;
//   //       }

//   //       return {
//   //         key: emp.ID,
//   //         text: `${emp.FirstName} ${emp.Title}`,
//   //         department: departmentText,
//   //         departmentGuid: departmentTermGuid, // Store the GUID term
//   //       };
//   //     });

//   //     this.setState({ employees: employeeOptions, isLoading: false });
//   //   } catch (error) {
//   //     this.setState({
//   //       errorMessage: "Error loading employees.",
//   //       isLoading: false,
//   //     });
//   //     console.error(error);
//   //   }
//   // }

//   // private handleEmployeeChange = (
//   //   event: React.FormEvent<HTMLDivElement>,
//   //   option?: IEmployeeOption // Use the extended type
//   // ): void => {
//   //   if (option) {
//   //     const filteredEmployees = this.state.employees.filter(
//   //       (emp) => emp.key === option.key
//   //     );
//   //     const selectedEmployee =
//   //       filteredEmployees.length > 0 ? filteredEmployees[0] : null;

//   //     let selectedDepartmentGuid = "";
//   //     if (selectedEmployee && "departmentGuid" in selectedEmployee) {
//   //       selectedDepartmentGuid = (selectedEmployee as IEmployeeOption)
//   //         .departmentGuid;
//   //     }

//   //     this.setState(
//   //       { selectedEmployee: option.key as string },
//   //       () => this.loadQuestions(selectedDepartmentGuid) // Pass the GUID term here
//   //     );
//   //   }
//   // };

//   // private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
//   //   if (!this.state.selectedEmployee || !selectedDepartmentGuid) return;

//   //   try {
//   //     this.setState({ isLoading: true });

//   //     const questions = await sp.web.lists
//   //       .getByTitle("QuestionBank")
//   //       .items.filter(`Department/TermGuid eq '${selectedDepartmentGuid}'`)
//   //       .get();

//   //     this.setState({
//   //       questions: questions.map((q) => ({
//   //         id: q.ID,
//   //         text: q.Title,
//   //         weight: q.Weight,
//   //       })),
//   //       scores: {},
//   //       isLoading: false,
//   //     });
//   //   } catch (error) {
//   //     this.setState({
//   //       errorMessage: "Error loading questions.",
//   //       isLoading: false,
//   //     });
//   //     console.error(error);
//   //   }
//   // }

//   // private async loadEmployees(): Promise<void> {
//   //   try {
//   //     this.setState({ isLoading: true });
//   //     const currentUser = await sp.web.currentUser.get();
//   //     console.log("this is current user:", currentUser);
//   //     // const currentUserLoginName = "i:0#.w|ipr-co\\mesgari-m"; // Replace with the actual user login
//   //     const employees = await sp.web.lists
//   //       .getByTitle("پرسنل معاونت مکانیک") // Replace with your actual list title
//   //       .items.select(
//   //         "ID",
//   //         "Title",
//   //         "FirstName",
//   //         "Department",
//   //         "Evaluator/Name",
//   //         "MechDepartment"
//   //       )
//   //       .expand("Evaluator")
//   //       .filter(`Evaluator/Name eq '${currentUser.LoginName}'`)
//   //       .get();

//   //     console.log("Filtered employees:", employees);

//   //     const employeeOptions: IEmployeeOption[] = employees.map((emp) => {
//   //       let departmentText = "";
//   //       if (emp.MechDepartment && emp.MechDepartment.Label) {
//   //         departmentText = emp.MechDepartment.Label;
//   //       }
//   //       return {
//   //         key: emp.ID,
//   //         text: `${emp.FirstName} ${emp.Title}`,
//   //         department: departmentText, // Use the extracted text value
//   //       };
//   //     });
//   //     // ({

//   //     //   key: emp.ID,
//   //     //   text: `${emp.FirstName} ${emp.Title}`,
//   //     //   department: emp.Department, // Assuming "Department" is a field in the list
//   //     // }));

//   //     this.setState({ employees: employeeOptions, isLoading: false });
//   //   } catch (error) {
//   //     this.setState({
//   //       errorMessage: "Error loading employees.",
//   //       isLoading: false,
//   //     });
//   //     console.error(error);
//   //   }
//   // }

//   // private handleEmployeeChange = (
//   //   event: React.FormEvent<HTMLDivElement>,
//   //   option?: IEmployeeOption // Use the extended type
//   // ): void => {
//   //   if (option) {
//   //     const filteredEmployees = this.state.employees.filter(
//   //       (emp) => emp.key === option.key
//   //     );
//   //     const selectedEmployee =
//   //       filteredEmployees.length > 0 ? filteredEmployees[0] : null;

//   //     let selectedDepartment = "";
//   //     if (selectedEmployee && "department" in selectedEmployee) {
//   //       selectedDepartment = (selectedEmployee as IEmployeeOption).department;
//   //     }

//   //     this.setState(
//   //       { selectedEmployee: option.key as string },
//   //       () => this.loadQuestions(selectedDepartment) // Pass it here
//   //     );
//   //   }
//   // };

//   // private async loadQuestions(selectedDepartment?: string): Promise<void> {
//   //   if (!this.state.selectedEmployee || !selectedDepartment) return;

//   //   try {
//   //     this.setState({ isLoading: true });

//   //     const questions = await sp.web.lists
//   //       .getByTitle("QuestionBank")
//   //       .items.filter(`Department eq '${selectedDepartment}'`)
//   //       .get();

//   //     this.setState({
//   //       questions: questions.map((q) => ({
//   //         id: q.ID,
//   //         text: q.QuestionText,
//   //         weight: q.QuestionWeight,
//   //       })),
//   //       scores: {},
//   //       isLoading: false,
//   //     });
//   //   } catch (error) {
//   //     this.setState({
//   //       errorMessage: "Error loading questions.",
//   //       isLoading: false,
//   //     });
//   //     console.error(error);
//   //   }
//   // }

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

//         sp.web.lists.getByTitle("EvaluationResults").items.inBatch(batch).add({
//           EmployeeID: selectedEmployee,
//           QuestionDescription: question.text,
//           Score: scores[question.id],
//           WeightedScore: weightedScore,
//           EvaluationPeriod: evaluationPeriod,
//         });
//       });

//       await batch.execute();

//       this.setState({ isLoading: false });
//       alert("Evaluation submitted successfully.");
//     } catch (error) {
//       this.setState({
//         errorMessage: "Error submitting evaluation.",
//         isLoading: false,
//       });
//       console.error(error);
//     }
//   };

//   render(): React.ReactElement<any> {
//     const {
//       employees,
//       selectedEmployee,
//       questions,
//       scores,
//       isLoading,
//       errorMessage,
//     } = this.state;

//     return (
//       <div>
//         <h3>{this.props.description}</h3> {/* Use the description prop */}
//         {isLoading && <Spinner size={SpinnerSize.large} label="Loading..." />}
//         {errorMessage && <Label style={{ color: "red" }}>{errorMessage}</Label>}
//         <Dropdown
//           label="Select Employee"
//           options={employees}
//           selectedKey={selectedEmployee}
//           onChange={this.handleEmployeeChange}
//           placeHolder="Choose an employee"
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
//       </div>
//     );
//   }
// }

// import * as React from "react";
// import { sp } from "@pnp/sp/presets/all";
// import {
//   Dropdown,
//   IDropdownOption,
//   PrimaryButton,
//   Spinner,
//   SpinnerSize,
//   Label,
// } from "office-ui-fabric-react";
// import { IPersonnelAppraisalProps } from "./IPersonnelAppraisalProps";

// interface IEmployeeOption extends IDropdownOption {
//   department: string;
//   departmentGuid: string;
// }

// interface IAppraisalFormState {
//   employees: IEmployeeOption[];
//   selectedEmployee: string | undefined;
//   questions: { id: number; text: string; weight: number }[];
//   scores: { [questionId: number]: number };
//   isLoading: boolean;
//   errorMessage: string | null;
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
//       console.log("this is current user:", currentUser);

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

//       const employeeOptions: IEmployeeOption[] = employees.map((emp) => {
//         let departmentText = "";
//         let departmentTermGuid = "";
//         if (emp.MechDepartment && emp.MechDepartment.Label) {
//           departmentText = emp.MechDepartment.Label;
//           departmentTermGuid = emp.MechDepartment.TermGuid;
//         }

//         return {
//           key: emp.ID,
//           text: `${emp.FirstName} ${emp.Title}`,
//           department: departmentText,
//           departmentGuid: departmentTermGuid,
//         };
//       });

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
//     event: React.FormEvent<HTMLDivElement>,
//     option?: IEmployeeOption
//   ): void => {
//     if (option) {
//       const filteredEmployees = this.state.employees.filter(
//         (emp) => emp.key === option.key
//       );
//       const selectedEmployee =
//         filteredEmployees.length > 0 ? filteredEmployees[0] : null;

//       let selectedDepartmentGuid = "";
//       if (selectedEmployee && "departmentGuid" in selectedEmployee) {
//         selectedDepartmentGuid = (selectedEmployee as IEmployeeOption)
//           .departmentGuid;
//       }

//       this.setState({ selectedEmployee: option.key as string }, () =>
//         this.loadQuestions(selectedDepartmentGuid)
//       );
//     }
//   };

//   private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
//     if (!this.state.selectedEmployee || !selectedDepartmentGuid) return;

//     try {
//       this.setState({ isLoading: true });

//       const questions = await sp.web.lists
//         .getByTitle("QuestionBank")
//         .items.filter(`Department/TermGuid eq '${selectedDepartmentGuid}'`)
//         .get();

//       this.setState({
//         questions: questions.map((q) => ({
//           id: q.ID,
//           text: q.Title,
//           weight: q.Weight,
//         })),
//         scores: {},
//         isLoading: false,
//       });
//     } catch (error) {
//       this.setState({
//         errorMessage: "Error loading questions.",
//         isLoading: false,
//       });
//       console.error(error);
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

//         sp.web.lists.getByTitle("EvaluationResults").items.inBatch(batch).add({
//           EmployeeID: selectedEmployee,
//           QuestionDescription: question.text,
//           Score: scores[question.id],
//           WeightedScore: weightedScore,
//           EvaluationPeriod: evaluationPeriod,
//         });
//       });

//       await batch.execute();

//       this.setState({ isLoading: false });
//       alert("Evaluation submitted successfully.");
//     } catch (error) {
//       this.setState({
//         errorMessage: "Error submitting evaluation.",
//         isLoading: false,
//       });
//       console.error(error);
//     }
//   };

//   render(): React.ReactElement<any> {
//     const {
//       employees,
//       selectedEmployee,
//       questions,
//       scores,
//       isLoading,
//       errorMessage,
//     } = this.state;

//     return (
//       <div>
//         <h3>{this.props.description}</h3>
//         {isLoading && <Spinner size={SpinnerSize.large} label="Loading..." />}
//         {errorMessage && <Label style={{ color: "red" }}>{errorMessage}</Label>}
//         <Dropdown
//           label="Select Employee"
//           options={employees}
//           selectedKey={selectedEmployee}
//           onChange={this.handleEmployeeChange}
//           placeHolder="Choose an employee"
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
//       </div>
//     );
//   }
// }

import * as React from "react";
import { sp } from "@pnp/sp/presets/all";
import {
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Label,
} from "office-ui-fabric-react";
import { IPersonnelAppraisalProps } from "./IPersonnelAppraisalProps";

interface IEmployeeOption extends IDropdownOption {
  department: string;
  departmentGuid: string;
}

interface IAppraisalFormState {
  employees: IEmployeeOption[];
  selectedEmployee: string | undefined;
  questions: { id: number; text: string; weight: number }[];
  scores: { [questionId: number]: number };
  isLoading: boolean;
  errorMessage: string | null;
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
    };
  }

  componentDidMount(): void {
    sp.setup({
      spfxContext: this.props.context,
    });
    this.loadEmployees();
  }

  private async loadEmployees(): Promise<void> {
    try {
      this.setState({ isLoading: true });
      const currentUser = await sp.web.currentUser.get();
      console.log("Current user:", currentUser);

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

      console.log("Filtered employees:", employees);

      const employeeOptions: IEmployeeOption[] = employees.map((emp) => {
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

      console.log("Employee options:", employeeOptions);
      this.setState({ employees: employeeOptions, isLoading: false });
    } catch (error) {
      this.setState({
        errorMessage: "Error loading employees.",
        isLoading: false,
      });
      console.error(error);
    }
  }

  private handleEmployeeChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IEmployeeOption
  ): void => {
    console.log("handleEmployeeChange called");
    if (option) {
      console.log("Selected option:", option);
      const filteredEmployees = this.state.employees.filter(
        (emp) => emp.key === option.key
      );
      const selectedEmployee =
        filteredEmployees.length > 0 ? filteredEmployees[0] : null;

      let selectedDepartmentGuid = "";
      if (selectedEmployee && "departmentGuid" in selectedEmployee) {
        selectedDepartmentGuid = (selectedEmployee as IEmployeeOption)
          .departmentGuid;
      }

      console.log("Selected employee:", selectedEmployee);
      console.log("Selected department GUID:", selectedDepartmentGuid);

      this.setState({ selectedEmployee: option.key as string }, () => {
        console.log(
          "State after setting selectedEmployee:",
          this.state.selectedEmployee
        );
        this.loadQuestions(selectedDepartmentGuid);
      });
    }
  };

  private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
    console.log("loadQuestions called");
    if (!this.state.selectedEmployee || !selectedDepartmentGuid) {
      console.log("No selected employee or department GUID");
      return;
    }

    try {
      this.setState({ isLoading: true });
      console.log(
        "Fetching questions for department GUID:",
        selectedDepartmentGuid
      );

      const questions = await sp.web.lists
        .getByTitle("QuestionBank")
        .items.filter(`Department/TermGuid eq '${selectedDepartmentGuid}'`)
        .get();

      console.log("Fetched questions:", questions);

      this.setState({
        questions: questions.map((q) => ({
          id: q.ID,
          text: q.Title,
          weight: q.Weight,
        })),
        scores: {},
        isLoading: false,
      });

      console.log("State after loading questions:", this.state);
    } catch (error) {
      this.setState({
        errorMessage: "Error loading questions.",
        isLoading: false,
      });
      console.error(error);
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

      questions.forEach((question) => {
        const weightedScore = (scores[question.id] / 5) * question.weight;

        sp.web.lists.getByTitle("EvaluationResults").items.inBatch(batch).add({
          EmployeeID: selectedEmployee,
          QuestionDescription: question.text,
          Score: scores[question.id],
          WeightedScore: weightedScore,
          EvaluationPeriod: evaluationPeriod,
        });
      });

      await batch.execute();

      this.setState({ isLoading: false });
      alert("Evaluation submitted successfully.");
    } catch (error) {
      this.setState({
        errorMessage: "Error submitting evaluation.",
        isLoading: false,
      });
      console.error(error);
    }
  };

  render(): React.ReactElement<any> {
    const {
      employees,
      selectedEmployee,
      questions,
      scores,
      isLoading,
      errorMessage,
    } = this.state;

    console.log("Rendering component with state:", this.state);

    return (
      <div>
        <h3>{this.props.description}</h3>
        {isLoading && <Spinner size={SpinnerSize.large} label="Loading..." />}
        {errorMessage && <Label style={{ color: "red" }}>{errorMessage}</Label>}
        <Dropdown
          label="Select Employee"
          options={employees}
          selectedKey={selectedEmployee}
          onChange={this.handleEmployeeChange}
          placeHolder="Choose an employee"
        />
        {questions.length > 0 && (
          <div>
            <table>
              <thead>
                <tr>
                  <th>Question</th>
                  <th>Score</th>
                </tr>
              </thead>
              <tbody>
                {questions.map((question) => (
                  <tr key={question.id}>
                    <td>{question.text}</td>
                    <td>
                      {[1, 2, 3, 4, 5].map((score) => (
                        <label key={score}>
                          <input
                            type="radio"
                            name={`question-${question.id}`}
                            value={score}
                            checked={scores[question.id] === score}
                            onChange={() =>
                              this.handleScoreChange(question.id, score)
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
            <PrimaryButton text="Submit" onClick={this.handleSubmit} />
          </div>
        )}
      </div>
    );
  }
}
