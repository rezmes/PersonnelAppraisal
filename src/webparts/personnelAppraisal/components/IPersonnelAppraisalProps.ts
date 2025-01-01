import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IPersonnelAppraisalProps {
  description: string;
  context: WebPartContext; // Add this line
  selectedDepartment: string; // Add this
}
