export interface IRiskInformationState {
  Title?: string;
  RiskId: number;
  ProjectID: string;
  RiskName: string;
  RiskDescription: string;
  RiskCategory: string;
  RiskIdentifiedOn: string;
  RiskClosedOn: string | null;
  RiskStatus: string;
  RiskOwner: string;
  RiskResponse: string;
  RiskImpact: string;
  RiskProbability: string;
  RiskRank: string;
  Remarks: string;
  focusedInput: any;
  FormDigestValue: string;
}