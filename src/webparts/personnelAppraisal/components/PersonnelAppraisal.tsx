// import * as React from 'react';
// import styles from './PersonnelAppraisal.module.scss';
// import { IPersonnelAppraisalProps } from './IPersonnelAppraisalProps';
// import { escape } from '@microsoft/sp-lodash-subset';

// export default class PersonnelAppraisal extends React.Component < IPersonnelAppraisalProps, {} > {
//   public render(): React.ReactElement<IPersonnelAppraisalProps> {
//     return(
//       <div className = { styles.personnelAppraisal } >
//   <div className={styles.container}>
//     <div className={styles.row}>
//       <div className={styles.column}>
//         <span className={styles.title}>Welcome to SharePoint!</span>
//         <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
//         <p className={styles.description}>{escape(this.props.description)}</p>
//         <a href='https://aka.ms/spfx' className={styles.button}>
//           <span className={styles.label}>Learn more</span>
//         </a>
//       </div>
//     </div>
//   </div>
//       </div >
//     );
//   }
// }
import * as React from "react";
import { Dropdown, IDropdownOption, TextField, PrimaryButton } from "office-ui-fabric-react";

export interface IEmployeeEvaluationProps {}
export interface IEmployeeEvaluationState {
  employees: IDropdownOption[];
  selectedEmployee: string;
  selectedEvaluationPeriod: string;
  evaluator: string;
  scores: { [key: string]: number };
  questions: { question: string; id: string }[];
}

class EmployeeEvaluation extends React.Component<IEmployeeEvaluationProps, IEmployeeEvaluationState> {
  constructor(props: IEmployeeEvaluationProps) {
    super(props);

    this.state = {
      employees: [],
      selectedEmployee: "",
      selectedEvaluationPeriod: "",
      evaluator: "Current User",
      scores: {},
      questions: [],
    };
  }

  componentDidMount() {
    // Simulate loading employees for the dropdown
    this.setState({
      employees: [
        { key: "emp1", text: "John Doe (HR)" },
        { key: "emp2", text: "Jane Smith (IT)" },
      ],
    });
  }

  handleEmployeeChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const value = event.target.value;

    // Sync selected employee and simulate loading questions
    this.setState({ selectedEmployee: value });

    const department = value.indexOf("HR") !== -1 ? "HR" : "IT"; // Use indexOf instead of includes

    // Simulate loading questions for the selected department
    this.setState({
      questions: department === "HR"
        ? [{ id: "q1", question: "Communication Skills" }]
        : [{ id: "q2", question: "Technical Expertise" }],
    });
  };


  handleEvaluationPeriodChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ selectedEvaluationPeriod: event.target.value });
  };

  handleScoreChange = (questionId: string, score: number) => {
    this.setState((prevState) => ({
      scores: { ...prevState.scores, [questionId]: score },
    }));
  };

  handleSave = () => {
    const { selectedEmployee, selectedEvaluationPeriod, evaluator, scores } = this.state;

    console.log("Saving Evaluation:");
    console.log("Employee:", selectedEmployee);
    console.log("Evaluation Period:", selectedEvaluationPeriod);
    console.log("Evaluator:", evaluator);
    console.log("Scores:", scores);

    // Logic to save to SharePoint would go here
    alert("Evaluation Saved!");
  };

  render() {
    const { employees, selectedEmployee, selectedEvaluationPeriod, questions } = this.state;

    return (
      <div>
        <h3>Employee Evaluation Form</h3>

        <TextField
          label="Evaluator"
          value={this.state.evaluator}
          readOnly
        />

        <TextField
          label="Evaluation Period"
          value={selectedEvaluationPeriod}
          onChange={this.handleEvaluationPeriodChange}
          list="period-options"
        />
        <datalist id="period-options">
          <option value="Quarter 1, 2023" />
          <option value="Quarter 2, 2023" />
        </datalist>

        <TextField
          label="Employee"
          value={selectedEmployee}
          onChange={this.handleEmployeeChange}
          list="employee-options"
        />
        <datalist id="employee-options">
          {employees.map((emp, idx) => (
            <option key={idx} value={emp.text} />
          ))}
        </datalist>

        <div>
          <h4>Evaluation Questions</h4>
          {questions.map((q) => (
            <div key={q.id}>
              <p>{q.question}</p>
              {[1, 2, 3, 4, 5].map((score) => (
                <label key={score}>
                  <input
                    type="radio"
                    name={`question-${q.id}`}
                    value={score}
                    onChange={() => this.handleScoreChange(q.id, score)}
                  />
                  {score}
                </label>
              ))}
            </div>
          ))}
        </div>

        <PrimaryButton text="Save Evaluation" onClick={this.handleSave} />
      </div>
    );
  }
}

export default EmployeeEvaluation;
