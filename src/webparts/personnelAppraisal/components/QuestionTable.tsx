// import * as React from "react";
// import "./PersonnelAppraisal.module.scss"; // Import your styles
// interface IQuestion {
//   id: number;
//   text: string;
//   weight: number;
// }

// interface IQuestionTableProps {
//   questions: IQuestion[];
//   scores: { [questionId: number]: number };
//   onScoreChange: (questionId: number, score: number) => void;
// }

// class QuestionTable extends React.Component<IQuestionTableProps, {}> {
//   render() {
//     return (
//       <table>
//         <thead>
//           <tr>
//             <th>شاخص ارزیابی</th>
//             <th>امتیاز</th>
//           </tr>
//         </thead>
//         <tbody>
//           {this.props.questions.map((question) => (
//             <tr key={question.id}>
//               <td>{question.text}</td>
//               <td>
//                 {[1, 2, 3, 4, 5].map((score) => (
//                   <label key={score}>
//                     <input
//                       type="radio"
//                       name={`question-${question.id}`}
//                       value={score}
//                       checked={this.props.scores[question.id] === score}
//                       onChange={() =>
//                         this.props.onScoreChange(question.id, score)
//                       }
//                     />
//                     {score}
//                   </label>
//                 ))}
//               </td>
//             </tr>
//           ))}
//         </tbody>
//       </table>
//     );
//   }
// }

// export default QuestionTable;
import * as React from "react";
import { Label } from "office-ui-fabric-react";

interface IQuestion {
  id: number;
  text: string;
  weight: number;
}

interface IQuestionTableProps {
  questions: IQuestion[];
  scores: { [questionId: number]: number };
  onScoreChange: (questionId: number, score: number) => void;
}

class QuestionTable extends React.Component<IQuestionTableProps, {}> {
  render() {
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
