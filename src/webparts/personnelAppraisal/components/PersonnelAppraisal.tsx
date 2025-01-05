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
//   selectedEmployee: string | number | undefined;
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
//     // Call the standalone test function
//     // this.testFetchQuestions();

//     this.loadEmployees();
//     // Call loadQuestions directly with a sample GUID
//     // this.loadQuestions("fe836f98-a77b-451b-916b-b59d0287ea0d");
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

//       // const employeeOptions: IEmployeeOption[] = employees.map((emp) => {
//       //   let departmentText = "";
//       //   let departmentTermGuid = "";
//       //   if (emp.MechDepartment && emp.MechDepartment.Label) {
//       //     departmentText = emp.MechDepartment.Label;
//       //     departmentTermGuid = emp.MechDepartment.TermGuid;
//       //   }

//       //   return {
//       //     key: emp.ID,
//       //     text: `${emp.FirstName} ${emp.Title}`,
//       //     department: departmentText,
//       //     departmentGuid: departmentTermGuid,
//       //   };
//       // });

//       const employeeOptions: IEmployeeOption[] = employees.map((emp) => {
//         let departmentText = "";
//         let departmentTermGuid = "";

//         return {
//           key: emp.ID.toString(), // Convert ID to string
//           text: `${emp.FirstName} ${emp.Title}`,
//           department: departmentText,
//           departmentGuid: departmentTermGuid,
//         };
//       });

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
//     event: React.FormEvent<HTMLDivElement>,
//     option?: IDropdownOption
//   ): void => {
//     console.log("Dropdown onChange triggered with option:", option);

//     if (option) {
//       console.log("Selected option:", option);
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

//       console.log("Selected employee:", selectedEmployee);
//       console.log("Selected department GUID:", selectedDepartmentGuid);

//       this.setState({ selectedEmployee: option.key as string }, () => {
//         console.log(
//           "State after setting selectedEmployee:",
//           this.state.selectedEmployee
//         );
//         this.loadQuestions(selectedDepartmentGuid);
//       });
//     }
//   };

//   // private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
//   //   console.log("loadQuestions called");
//   //   // if (!this.state.selectedEmployee || !selectedDepartmentGuid) {
//   //   //   console.log("No selected employee or department GUID");
//   //   //   return;
//   //   // }

//   //   try {
//   //     this.setState({ isLoading: true });
//   //     console.log(
//   //       "Fetching questions for department GUID:",
//   //       selectedDepartmentGuid
//   //     );

//   //     // const questions = await sp.web.lists
//   //     //   .getByTitle("QuestionBank")
//   //     //   .items.filter(`Department/TermGuid eq '${selectedDepartmentGuid}'`)
//   //     //   .get();

//   //     const questions = await sp.web.lists
//   //       .getByTitle("QuestionBank")
//   //       .items.get(); // Fetch all questions without filtering

//   //     console.log("Fetched questions:", questions);
//   //     console.log(
//   //       "Fetching questions for department GUID:",
//   //       selectedDepartmentGuid
//   //     );
//   //     this.setState({
//   //       questions: questions.map((q) => ({
//   //         id: q.ID,
//   //         text: q.Title,
//   //         weight: q.Weight,
//   //       })),
//   //       scores: {},
//   //       isLoading: false,
//   //     });

//   //     console.log("State after loading questions:", this.state);
//   //   } catch (error) {
//   //     console.error("Error loading questions:", error);
//   //     this.setState({
//   //       errorMessage: "Error loading questions.",
//   //       isLoading: false,
//   //     });
//   //     console.error(error);
//   //   }
//   // }

//   private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
//     console.log(
//       "loadQuestions called with selectedDepartmentGuid:",
//       selectedDepartmentGuid
//     );

//     try {
//       this.setState({ isLoading: true });

//       // Fetch all items without filtering
//       const questions = await sp.web.lists
//         .getByTitle("QuestionBank")
//         .items.select("ID", "Title", "QuestionWeight", "Department")
//         .get();

//       console.log("Fetched questions:", questions);

//       // Inspect the Department field structure
//       questions.forEach((q) => {
//         console.log("Question ID:", q.ID, "Department Field:", q.Department);
//       });

//       this.setState({
//         questions: questions.map((q) => ({
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

//   // testFetchQuestions ///////////////////////////////////////////////////////////////////
//   private async testFetchQuestions(): Promise<void> {
//     try {
//       console.log("Fetching all questions from the QuestionBank list...");

//       const questions = await sp.web.lists
//         .getByTitle("QuestionBank")
//         .items.select("ID", "Title", "QuestionWeight", "Department")
//         .get();

//       console.log("Fetched questions:", questions);

//       // Log each question's Department field to verify its structure
//       questions.forEach((q) => {
//         console.log(
//           "Question ID:",
//           q.ID,
//           "Title:",
//           q.Title,
//           "Department Field:",
//           q.Department
//         );
//       });
//     } catch (error) {
//       console.error("Error fetching questions:", error);
//     }
//   }

//   render(): React.ReactElement<any> {
//     const {
//       employees,
//       selectedEmployee,
//       questions,
//       scores,
//       isLoading,
//       errorMessage,
//     } = this.state;

//     console.log("Rendering component with state:", this.state);

//     return (
//       <div>
//         <h3>{this.props.description}</h3>
//         {isLoading && <Spinner size={SpinnerSize.large} label="Loading..." />}
//         {errorMessage && <Label style={{ color: "red" }}>{errorMessage}</Label>}
//         {/* <Dropdown
//           label="Select Employee"
//           options={this.state.employees}
//           selectedKey={this.state.selectedEmployee}
//           onChange={this.handleEmployeeChange}
//           placeHolder="Choose an employee"
//         /> */}
//         <select
//           onChange={(e) => {
//             const selectedKey = e.target.value;
//             console.log("Selected key:", selectedKey);
//             this.setState({ selectedEmployee: selectedKey });
//           }}
//         >
//           <option value="" disabled selected>
//             Choose an employee
//           </option>
//           {this.state.employees.map((emp) => (
//             <option key={emp.key} value={emp.key}>
//               {emp.text}
//             </option>
//           ))}
//         </select>

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
//   // render(): React.ReactElement<any> {
//   //   return (
//   //     <div>
//   //       <button onClick={() => this.testFetchQuestions()}>
//   //         Test Fetch Questions
//   //       </button>
//   //       <button onClick={() => this.loadQuestions()}>
//   //         Test Load Questions
//   //       </button>
//   //     </div>
//   //   );
//   // }
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
  selectedEmployee: string | number | undefined;
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

      const evaluationPeriod = "Q1-2024";

      const evaluatedEmployees = await sp.web.lists
        .getByTitle("EvaluationResults")
        .items.select("EmployeeID/ID", "EvaluationPeriod")
        .expand("EmployeeID") // Expand the EmployeeID lookup field
        .filter(`EvaluationPeriod eq '${evaluationPeriod}'`)
        .get();

      console.log("Evaluated employees:", evaluatedEmployees);

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

  // private async loadEmployees(): Promise<void> {
  //   try {
  //     this.setState({ isLoading: true });
  //     const currentUser = await sp.web.currentUser.get();
  //     console.log("Current user:", currentUser);

  //     const employees = await sp.web.lists
  //       .getByTitle("پرسنل معاونت مکانیک")
  //       .items.select(
  //         "ID",
  //         "Title",
  //         "FirstName",
  //         "Department",
  //         "Evaluator/Name",
  //         "MechDepartment"
  //       )
  //       .expand("Evaluator")
  //       .filter(`Evaluator/Name eq '${currentUser.LoginName}'`)
  //       .get();

  //     console.log("Filtered employees:", employees);

  //     const employeeOptions: IEmployeeOption[] = employees.map((emp) => {
  //       return {
  //         key: emp.ID.toString(),
  //         text: `${emp.FirstName} ${emp.Title}`,
  //         department: emp.Department,
  //         departmentGuid: emp.MechDepartment.TermGuid,
  //       };
  //     });

  //     console.log("Employee options:", employeeOptions);
  //     this.setState({ employees: employeeOptions, isLoading: false });
  //   } catch (error) {
  //     this.setState({
  //       errorMessage: "Error loading employees.",
  //       isLoading: false,
  //     });
  //     console.error(error);
  //   }
  // }

  // private handleEmployeeChange = (
  //   event: React.FormEvent<HTMLDivElement>,
  //   option?: IDropdownOption
  // ): void => {
  //   console.log("Dropdown onChange triggered with option:", option);

  //   if (option) {
  //     console.log("Selected option:", option);
  //     const filteredEmployees = this.state.employees.filter(
  //       (emp) => emp.key === option.key
  //     );
  //     const selectedEmployee =
  //       filteredEmployees.length > 0 ? filteredEmployees[0] : null;

  //     let selectedDepartmentGuid = "";
  //     if (selectedEmployee && "departmentGuid" in selectedEmployee) {
  //       selectedDepartmentGuid = (selectedEmployee as IEmployeeOption)
  //         .departmentGuid;
  //     }

  //     console.log("Selected employee:", selectedEmployee);
  //     console.log("Selected department GUID:", selectedDepartmentGuid);

  //     this.setState({ selectedEmployee: option.key as string }, () => {
  //       console.log(
  //         "State after setting selectedEmployee:",
  //         this.state.selectedEmployee
  //       );
  //       this.loadQuestions(selectedDepartmentGuid);
  //     });
  //   }
  // };

  // private async loadQuestions(selectedDepartmentGuid?: string): Promise<void> {
  //   console.log(
  //     "loadQuestions called with selectedDepartmentGuid:",
  //     selectedDepartmentGuid
  //   );

  //   try {
  //     this.setState({ isLoading: true });

  //     const questions = await sp.web.lists
  //       .getByTitle("QuestionBank")
  //       .items.select("ID", "Title", "QuestionWeight", "Department")
  //       .get();

  //     console.log("Fetched questions:", questions);

  //     questions.forEach((q) => {
  //       console.log("Question ID:", q.ID, "Department Field:", q.Department);
  //     });

  //     this.setState({
  //       questions: questions.map((q) => ({
  //         id: q.ID,
  //         text: q.Title,
  //         weight: q.QuestionWeight,
  //       })),
  //       scores: {},
  //       isLoading: false,
  //     });
  //   } catch (error) {
  //     this.setState({
  //       errorMessage: "Error loading questions.",
  //       isLoading: false,
  //     });
  //     console.error("Error fetching questions:", error);
  //   }
  // }

  private handleEmployeeChange = (
    event: React.ChangeEvent<HTMLSelectElement>,
    option?: IDropdownOption
  ): void => {
    console.log("Dropdown onChange triggered with option:", option);

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
    console.log(
      "loadQuestions called with selectedDepartmentGuid:",
      selectedDepartmentGuid
    );

    try {
      this.setState({ isLoading: true });

      const questions = await sp.web.lists
        .getByTitle("QuestionBank")
        .items.select("ID", "Title", "QuestionWeight", "Department")
        .get();

      console.log("Fetched questions:", questions);

      const filteredQuestions = questions.filter((q) => {
        // Assuming Department is stored as a term GUID string in q.Department.TermGuid
        return q.Department.TermGuid === selectedDepartmentGuid;
      });

      console.log("Filtered questions:", filteredQuestions);

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

  private handleScoreChange = (questionId: number, score: number): void => {
    this.setState((prevState) => ({
      scores: {
        ...prevState.scores,
        [questionId]: score,
      },
    }));
  };

  // private handleSubmit: () => Promise<void> = async (): Promise<void> => {
  //   const { selectedEmployee, scores, questions } = this.state;

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

  //     const batch = sp.web.createBatch();
  //     const evaluationPeriod = "Q1-2024";

  //     questions.forEach((question) => {
  //       const weightedScore = (scores[question.id] / 5) * question.weight;

  //       // Assuming EmployeeID is a lookup field, use the proper structure for lookup fields
  //       const item = {
  //         EmployeeIDId: selectedEmployee, // Use the lookup field suffix 'Id'
  //         QuestionDescription: question.text,
  //         Score: scores[question.id],
  //         WeightedScore: weightedScore,
  //         EvaluationPeriod: evaluationPeriod,
  //       };

  //       console.log("Adding item to batch:", item);

  //       sp.web.lists
  //         .getByTitle("EvaluationResults")
  //         .items.inBatch(batch)
  //         .add(item);
  //     });

  //     await batch.execute();

  //     this.setState({ isLoading: false });
  //     alert("Evaluation submitted successfully.");
  //   } catch (error) {
  //     this.setState({
  //       errorMessage: "Error submitting evaluation.",
  //       isLoading: false,
  //     });
  //     console.error("Error submitting evaluation:", error);
  //   }
  // };
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

        const item = {
          EmployeeIDId: selectedEmployee,
          QuestionDescription: question.text,
          Score: scores[question.id],
          WeightedScore: weightedScore,
          EvaluationPeriod: evaluationPeriod,
        };

        console.log("Adding item to batch:", item);

        sp.web.lists
          .getByTitle("EvaluationResults")
          .items.inBatch(batch)
          .add(item);
      });

      await batch.execute();

      this.setState({ isLoading: false });
      alert("Evaluation submitted successfully.");

      // Re-load employees to update the dropdown list
      this.loadEmployees();
    } catch (error) {
      this.setState({
        errorMessage: "Error submitting evaluation.",
        isLoading: false,
      });
      console.error("Error submitting evaluation:", error);
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
        <select
          onChange={(e) => {
            const selectedKey = e.target.value;
            console.log("Selected key:", selectedKey);
            const selectedOption = this.state.employees.filter(
              (emp) => emp.key === selectedKey
            )[0];
            console.log("Selected option:", selectedOption);

            if (selectedOption) {
              this.setState({ selectedEmployee: selectedOption.key }, () => {
                this.loadQuestions(selectedOption.departmentGuid);
              });
            }
          }}
        >
          <option value="" disabled selected>
            Choose an employee
          </option>
          {this.state.employees.map((emp) => (
            <option key={emp.key} value={emp.key}>
              {emp.text}
            </option>
          ))}
        </select>

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
