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

interface IAppraisalFormState {
  employees: IDropdownOption[];
  selectedEmployee: string | undefined;
  questions: { id: number; text: string; weight: number }[];
  scores: { [questionId: number]: number };
  isLoading: boolean;
  errorMessage: string | null;
}

export default class PersonnelAppraisal extends React.Component<
  {},
  IAppraisalFormState
> {
  constructor(props: {}) {
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
      spfxContext: this.context,
    });
    this.loadEmployees();
  }

  private async loadEmployees(): Promise<void> {
    try {
      this.setState({ isLoading: true });
      const employees = await sp.web.lists
        .getByTitle("Personnel")
        .items.select("ID", "FirstName", "LastName")
        .get();

      const employeeOptions = employees.map((emp) => ({
        key: emp.ID,
        text: `${emp.FirstName} ${emp.LastName}`,
      }));

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
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState(
        { selectedEmployee: option.key as string },
        this.loadQuestions
      );
    }
  };

  private async loadQuestions(): Promise<void> {
    if (!this.state.selectedEmployee) return;

    try {
      this.setState({ isLoading: true });

      const questions = await sp.web.lists
        .getByTitle("QuestionBank")
        .items.filter(`Department eq 'HR'`)
        .select("ID", "QuestionText", "QuestionWeight")
        .get();

      this.setState({
        questions: questions.map((q) => ({
          id: q.ID,
          text: q.QuestionText,
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

    return (
      <div>
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


