import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IPersonnelAppraisalProps {
  description: string;
  context: any;
  employeeListName: string;
  evaluationResultsListName: string;
  evaluationPeriodListName: string;
  questionBankListName: string;

}
