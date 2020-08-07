export interface SPProjectListEditForm {
  ProjectID: string;
  Project_x0020_Name: string;
  Client_x0020_Name: string;
  Delivery_x0020_Manager: string;
  Project_x0020_Manager: string;
  Project_x0020_Mode: string;
  Project_x0020_Type: string;
  Project_x0020_Phase: string;
  Project_x0020_Description: string;
  PlannedStart: string;
  Planned_x0020_End: string;
  Region: string;
  Status: string;
  Progress: number;
  Actual_x0020_Start: string;
  Actual_x0020_End: string;
  Revised_x0020_Budget: number;
  Total_x0020_Cost: number;
  Invoiced_x0020_amount: number;
  Scope: string;
  Schedule: string;
  Resource: string;
  Project_x0020_Cost: string;
  PMId:number;
  DMId:number;
  Previous_PM:number;
  Previous_DM:number;
}
