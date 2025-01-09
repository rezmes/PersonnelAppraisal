import * as React from "react";

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
    this.loadLatestPeriod();
  }

  private async loadLatestPeriod(): Promise<void> {
    try {
      this.setState({ isLoading: true });

      const response = await fetch(
        `${this.props.spfxContext.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EvaluationPeriod')/items?$orderby=Created desc&$top=1`,
        {
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }
      );

      const data = await response.json();

      if (data.d.results.length > 0) {
        const evaluationPeriod = data.d.results[0].Title;
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

    return <div id="PeriodTitle">دوره ارزیابی جاری: {evaluationPeriod}</div>;
  }
}

export default EvaluationPeriod;
