// import * as React from "react";
// import { sp } from "@pnp/sp/presets/all";
// import "./PersonnelAppraisal.module.scss"; // Import your styles

// interface IEvaluationPeriodState {
//   evaluationPeriod: string;
//   isLoading: boolean;
//   errorMessage: string | null;
// }

// interface IEvaluationPeriodProps {
//   spfxContext: any;
//   onPeriodLoaded: (period: string) => void;
// }

// class EvaluationPeriod extends React.Component<
//   IEvaluationPeriodProps,
//   IEvaluationPeriodState
// > {
//   constructor(props: IEvaluationPeriodProps) {
//     super(props);
//     this.state = {
//       evaluationPeriod: "",
//       isLoading: false,
//       errorMessage: null,
//     };
//   }

//   componentDidMount(): void {
//     sp.setup({
//       spfxContext: this.props.spfxContext,
//     });
//     this.loadLatestPeriod();
//   }

//   private async loadLatestPeriod(): Promise<void> {
//     try {
//       this.setState({ isLoading: true });

//       const items = await sp.web.lists
//         .getByTitle("EvaluationPeriod")
//         .items.orderBy("Created", false)
//         .top(1)
//         .get();

//       if (items.length > 0) {
//         const evaluationPeriod = items[0].Title;
//         this.setState({ evaluationPeriod, isLoading: false });
//         this.props.onPeriodLoaded(evaluationPeriod);
//       } else {
//         this.setState({
//           errorMessage: "دوره ارزیابی یافت نشد.",
//           isLoading: false,
//         });
//       }
//     } catch (error) {
//       this.setState({
//         errorMessage: "خطا در بارگذاری دوره ارزیابی",
//         isLoading: false,
//       });
//       console.error("بارگذاری این دوره با خطا مواجه شد: ", error);
//     }
//   }

//   render(): React.ReactElement<any> {
//     const { evaluationPeriod, isLoading, errorMessage } = this.state;

//     if (isLoading) {
//       return <div>بارگذاری دوره ی ارزیابی ...</div>;
//     }

//     if (errorMessage) {
//       return <div style={{ color: "red" }}>{errorMessage}</div>;
//     }

//     return <div id="PeriodTitle">دوره ارزیابی جاری: {evaluationPeriod}</div>;
//   }
// }

// export default EvaluationPeriod;

import * as React from "react";
import { sp } from "@pnp/sp/presets/all";

interface IEvaluationPeriodState {
  evaluationPeriod: string;
  isLoading: boolean;
  errorMessage: string | null;
}

interface IEvaluationPeriodProps {
  spfxContext: any;
  onPeriodLoaded: (period: string) => void;
}

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
    sp.setup({
      spfxContext: this.props.spfxContext,
    });
    this.loadLatestPeriod();
  }

  private async loadLatestPeriod(): Promise<void> {
    try {
      this.setState({ isLoading: true });

      const items = await sp.web.lists
        .getByTitle("EvaluationPeriod")
        .items.orderBy("Created", false)
        .top(1)
        .get();

      if (items.length > 0) {
        const evaluationPeriod = items[0].Title;
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

    return (
      <div style={{ textAlign: "left", fontFamily: "IRANSansXFaNum" }}>
        دوره ارزیابی جاری: {evaluationPeriod}
      </div>
    );
  }
}

export default EvaluationPeriod;
