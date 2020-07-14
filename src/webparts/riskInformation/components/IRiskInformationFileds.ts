export interface ISPRiskInformationFields{
  ID?:string;
  ProjectID: string;
  //RiskID: string;
  RiskName: string;
  RiskDescription: string;
  RiskCategory:string;
  RiskIdentifiedOn: string;
  RiskClosedOn: string;
  RiskStatus: string;
  RiskOwner: string;
  RiskResponse: string;
  RiskImpact: string;
  RiskProbability: string;
  Remarks: string;
  RiskRank: string
}