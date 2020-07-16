export interface ISPMilestoneFields{
    ID?:string;
    ProjectID: string;    
    Phase: string;
    PlannedStart: string;
    PlannedEnd:string;
    MilestoneStatus: string;
    Remarks: string;
    Created?: string;
    Modified?: string;
    ActualStart: string;
    ActualEnd: string;      
  }