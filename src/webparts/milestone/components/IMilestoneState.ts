export interface IMilestoneState {
  ID?: string;
  ProjectID: string;
  Phase: string;
  PlannedStart: string;
  PlannedEnd: string;
  MilestoneStatus: string;
  Remarks: string;
  MilestoneCreatedOn : string;
  LastUpdatedOn : string;
  ActualStart: string;
  ActualEnd: string;
  focusedInput: any;
  FormDigestValue: string;
}