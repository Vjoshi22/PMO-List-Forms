import * as React from 'react';
import styles from './PmoListForms.module.scss';
import { IPmoListFormsProps } from './IPmoListFormsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration  ,SPHttpClientResponse, HttpClientResponse} from "@microsoft/sp-http";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker"; 
import { _getParameterValues } from './getQueryString';
import { Form, FormGroup, Button, FormControl } from "react-bootstrap";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { SPProjectListEditForm } from "../components/IEditFormProps";
import * as $ from "jquery";
import { _getListEntityName, listType } from './getListEntityName';
import { data } from 'jquery';


require('./PmoListForms.module.scss');
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");

var allchoiceColumnsEditForm: any[] = ["Project_x0020_Type", "Project_x0020_Mode", "Project_x0020_Cost" ,"Status", "Scope", "Resource", "Schedule"];

export interface IreactState{
    ProjectID: string,
    ProjectName: string;
    ClientName: string;
    ProjectManager: string;
    ProjectType: string;
    ProjectMode: string;
    PlannedStart: string;
    PlannedCompletion: string;
    ProjectDescription: string;
    ProjectLocation: string;
    // ProjectBudget: string;
    ProjectStatus: string;
    ProjectProgress: string;
    ActualStartDate:string; //edit only
    ActualEndDate:string; //edit only
    RevisedBudget:string; //edit only
    TotalCost:string; //edit only
    InvoicedAmount: string; //edit only
    ProjectScope: string; // Project Scope edit only
    ProjectSchedule: string; //project scheduled edit only
    ProjectResource: string;
    ProjectCost: string; //only in edit
    //peoplepicker
    DeliveryManager: string;
    //date
    startDate: any;
    disable_RMSID: boolean;
    disable_plannedCompletion: boolean;
    endDate: any;  
    focusedInput: any;
    FormDigestValue: string;
}

var listGUID: any = "2c3ffd4e-1b73-4623-898d-8e3a1bb60b91";   //"47272d1e-57d9-447e-9cfd-4cff76241a93"; 
var timerID;
var newitem: boolean;

export default class PmoListEditForm extends React.Component<IPmoListFormsProps, IreactState> {
    constructor(props: IPmoListFormsProps, state: IreactState) {  
        super(props);  
      
        this.state = {  
            ProjectID : '',
            ProjectName: '',
            ClientName: '',
            ProjectManager: '',
            ProjectType: '',
            ProjectMode: '',
            ProjectDescription: '',
            PlannedStart: '',
            PlannedCompletion: '',
            ProjectLocation: '',
            ProjectProgress:'',
            ProjectStatus: '',
            ActualStartDate:'',
            ActualEndDate:'',
            RevisedBudget:'',
            TotalCost:'',
            InvoicedAmount: '',
            ProjectScope:'',
            ProjectSchedule: '',
            ProjectResource: '',
            ProjectCost: '',
            DeliveryManager:'',
            startDate: '',
            endDate: '',
            disable_RMSID: false,
            disable_plannedCompletion: true,
            focusedInput: '',
            FormDigestValue:''
        };  
        this._saveItem=this._saveItem.bind(this);
        this.handleChange=this.handleChange.bind(this);
        this._getProjectManager =this._getProjectManager.bind(this);
        //this.loadItems = this.loadItems.bind(this);
        //this.isOutsideRange = this.isOutsideRange.bind(this);
      }
      public componentDidMount() {
          $('.webPartContainer').hide();
        //calling function to fetch dropdown values form sp choice coluns
        //window.addEventListener('load', this.handleload)
        allchoiceColumnsEditForm.forEach(colName => {
            this._retrieveAllChoicesFromListField(this.props.currentContext.pageContext.web.absoluteUrl, colName);
          });
        _getListEntityName(this.props.currentContext, listGUID);
        $('.pickerText_4fe0caaf').css('border','0px');
        $('.pickerInput_4fe0caaf').addClass('form-control');
        $('.form-row').css('justify-content','center');

        //this._loadItems();
        setTimeout(() =>this._loadItems(), 1000);
        
        this._getAccessToken();
        timerID=setInterval(
          () =>this._getAccessToken(),300000); 
     }
     public componentWillUnmount()
     {
      clearInterval(timerID);
      //window.removeEventListener('load', this.handleload)
      
     } 
    //  private handleload(){
    //     allchoiceColumnsEditForm.forEach(colName => {
    //         this._retrieveAllChoicesFromListField(this.props.currentContext.pageContext.web.absoluteUrl, colName);
    //       });
    //  }
     //public  isOutsideRange = day =>day.isAfter(Moment()) || day.isBefore(Moment().subtract(0, "days"));  
      private handleChange = (e) => {
        let newState = {};
        newState[e.target.name] = e.target.value;
        this.setState(newState);
        //fun to validate date
        this._validateDate(e);
        //func to validate progrerss
        this._validateProgress(e);
         //functin to check the existing Id
        if(e.target.name == "ProjectID" && (e.target.value != 0 || e.target.value =="")){
         this._checkExistingProjectId(this.props.currentContext.pageContext.web.absoluteUrl, e.target.value);
        } else if(e.target.value == 0){
            $('.ProjectID').remove();
            $('#ProjectId').closest('div').append('<span class="ProjectID" style="color:red;font-size:9pt">Project Id cannot be 0</span>');
            }
    }
      private _handleSubmit = (e) =>{
        this._saveItem(e);
      }
      private _getProjectManager = (items: any[]) => {  
        console.log('Items:', items);  
        this.setState({ ProjectManager: items[0].text });
      }
      private _getDeliveryManager = (items: any[]) => {  
        console.log('Items:', items);  
        this.setState({ DeliveryManager: items[0].text });
      }

    public render(): React.ReactElement<IPmoListFormsProps> {
    
        return (
            <div id="newItemDiv" className={styles["_main-div"]} >
        <div id="heading" className={styles.heading}><h3>Project Details</h3></div>
        <Form onSubmit={this._handleSubmit}>
            <Form.Row className="mt-3">
            {/*-----------Project ID------------------- */}
                <FormGroup className="col-2">
                    <Form.Label className={styles.customlabel}>Project Id</Form.Label>
                </FormGroup>
                <FormGroup className={styles.disabledValue + " col-3"}>
                    {/* Please check: --- disable RMS id to be removed */}
                    {/* <Form.Control size="sm" type="text" disabled={this.state.disable_RMSID} id="ProjectId" name="ProjectID" placeholder="Project Id" onChange={this.handleChange} value={this.state.ProjectID}/> */}
                    <Form.Label>{this.state.ProjectID}</Form.Label>
                </FormGroup>
                <FormGroup className="col-1"></FormGroup>
                {/*-----------Project Type------------- */}
                <FormGroup className="col-2">
                    <Form.Label className={styles.customlabel}>Project Type</Form.Label>
                </FormGroup>
                <FormGroup className={styles.disabledValue + " col-3"}>
                    {/* Other options appending from sp list column using retrieveAllChoicesFromListField fun */}
                    {/* <Form.Control size="sm" id="ProjectType" as="select" name="ProjectType" onChange={this.handleChange} value={this.state.ProjectType}>
                        <option value="">Select an Option</option>
                    </Form.Control> */}
                    <Form.Label>{this.state.ProjectType}</Form.Label>
                </FormGroup>
            </Form.Row>

            <Form.Row>
                {/* -----------Client Name------------- */}
                <FormGroup className="col-2">
                    <Form.Label className={styles.customlabel}>Client Name</Form.Label>
                </FormGroup>
                <FormGroup className={styles.disabledValue + " col-3"}>
                    {/* <Form.Control size="sm" type="text" id="ClientName" name="ClientName" placeholder="Client Name" onChange={this.handleChange} value={this.state.ClientName}/> */}
                    <Form.Label>{this.state.ClientName}</Form.Label>
                </FormGroup>
                <FormGroup className="col-1"></FormGroup>
                {/* -----------Project Name---------------- */}
                <FormGroup className="col-2">
                    <Form.Label className={styles.customlabel}>Project Name</Form.Label>
                </FormGroup>
                <FormGroup className={styles.disabledValue + " col-3"}>
                    {/* <Form.Control size="sm" type="text" id="ProjectName" name="ProjectName" placeholder="Ex: John Deer" onChange={this.handleChange} value={this.state.ProjectName} /> */}
                    <Form.Label>{this.state.ProjectName}</Form.Label>
                </FormGroup>
            </Form.Row>

            <Form.Row>
                {/* --------Delivery Manager------------ */}
                <FormGroup className="col-2">
                    <Form.Label className={styles.customlabel + " " + styles.required}>Delivery Manager</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <div id="DeliveryManager">
                        <PeoplePicker
                        context={this.props.currentContext}
                        personSelectionLimit={1}    
                        groupName={""} // Leave this blank in case you want to filter from all users    
                        showtooltip={true}    
                        isRequired={true}    
                        disabled={false}    
                        ensureUser={true}  
                        selectedItems={this._getDeliveryManager}   
                        defaultSelectedUsers={[this.state.DeliveryManager]} 
                        showHiddenInUI={false}    
                        principalTypes={[PrincipalType.User]}    
                        resolveDelay={1000} />
                    </div>
                </FormGroup>
                <FormGroup className="col-1"></FormGroup>
                {/*--------Project Manager-------------  */}
                <FormGroup className="col-2">
                    <Form.Label className={styles.customlabel +" " + styles.required}>Project Manager</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <div id="ProjectManager">
                        <PeoplePicker
                        context={this.props.currentContext}   
                        personSelectionLimit={1}    
                        groupName={""} // Leave this blank in case you want to filter from all users    
                        showtooltip={true}    
                        isRequired={true}    
                        disabled={false}    
                        ensureUser={true}    
                        selectedItems={this._getProjectManager} 
                        defaultSelectedUsers={[this.state.ProjectManager]}   
                        showHiddenInUI={false}    
                        principalTypes={[PrincipalType.User]}    
                        resolveDelay={1000} />
                    </div>
                </FormGroup>
            </Form.Row>
            <Form.Row> 
                <FormGroup className="col-2">
                    <Form.Label className={styles.customlabel}>Project Mode</Form.Label>
                </FormGroup>
                <FormGroup className= {styles.disabledValue + " col-3"}>
                    {/* <Form.Control size="sm" id="ProjectMode" as="select" name="ProjectMode" onChange={this.handleChange} value={this.state.ProjectMode}>
                        <option value="">Select an Option</option>
                    </Form.Control> */}
                    <Form.Label>{this.state.ProjectMode}</Form.Label>
                </FormGroup>
                <FormGroup className="col-1"></FormGroup>
                <FormGroup className="col-2">
                    <Form.Label className={styles.customlabel}>Project Status</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <Form.Control size="sm" id="Status" as="select" name="ProjectStatus"  onChange={this.handleChange} value={this.state.ProjectStatus}>
                        <option value="">Select an Option</option>
                    </Form.Control>
                </FormGroup>
             </Form.Row>
             <Form.Row>
          <FormGroup className="col-2"> 
            <Form.Label className={styles.customlabel}>Tentative Start Date</Form.Label>
          </FormGroup>
          <FormGroup className="col-3">
            {/* <Form.Control size="sm" type="date" id="PlannedStart" name="PlannedStart" placeholder="Planned Start Date" onChange={this.handleChange} value={this.state.PlannedStart}/> */}
            {/* <DatePicker selected={this.state.PlannedStart}  onChange={this.handleChange} />; */}
            <Form.Label>{this.state.PlannedStart}</Form.Label>
          </FormGroup>
          <FormGroup className="col-1"></FormGroup>
          <FormGroup className="col-2"> 
            <Form.Label className={styles.customlabel}>Tentative End Date</Form.Label>
          </FormGroup>
          <FormGroup className="col-3">
            {/* <Form.Control size="sm" type="date" disabled={this.state.disable_plannedCompletion} id="PlannedCompletion" name="PlannedCompletion" placeholder="Planned Completion Date" onChange={this.handleChange} value={this.state.PlannedCompletion}/> */}
            <Form.Label>{this.state.PlannedCompletion}</Form.Label>
          </FormGroup>
        </Form.Row>
            <Form.Row>
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel}>Actual Start Date</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <Form.Control size="sm" type="date" id="ActualStartDate" name="ActualStartDate" placeholder="Actual Start Date" onChange={this.handleChange} value={this.state.ActualStartDate}/>
                    {/* <DatePicker selected={this.state.PlannedStart}  onChange={this.handleChange} />; */}
                </FormGroup>
                <FormGroup className="col-1"></FormGroup>
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel}>Actual End Date</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <Form.Control size="sm" type="date" disabled={this.state.disable_plannedCompletion} id="ActualEndDate" name="ActualEndDate" placeholder="Planned Completion Date" onChange={this.handleChange} value={this.state.ActualEndDate}/>
                </FormGroup>
             </Form.Row>
            {/* Project Description */}
            <Form.Row>
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel + " " + styles.required }>Project Description</Form.Label>
                </FormGroup>
                <FormGroup className="col-9 mb-3">
                    <Form.Control size="sm" as="textarea" rows={4} type="text" id="ProjectDescription" name="ProjectDescription" placeholder="Project Description" onChange={this.handleChange} value={this.state.ProjectDescription}/>
                </FormGroup>
            </Form.Row>
            {/* Next Row */}
            <Form.Row>
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel}>Region</Form.Label>
                </FormGroup>
                <FormGroup className={styles.disabledValue + " col-3"}>
                    {/* <Form.Control size="sm" type="text" id="Region" name="Region" placeholder="Project Location" onChange={this.handleChange} value={this.state.ProjectLocation}/> */}
                    <Form.Label>{this.state.ProjectLocation}</Form.Label>
                </FormGroup>
                <FormGroup className="col-1"></FormGroup>
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel +" " + styles.required}>Revised Budget</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <Form.Control size="sm" type="number" min="1" id="RevisedBudget" name="RevisedBudget" placeholder="Revised Budget" onChange={this.handleChange} value={this.state.RevisedBudget}/>
                </FormGroup>
            </Form.Row>
            <Form.Row>
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel}>Project Progress</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <Form.Control size="sm" type="number" min="1" id="ProjectProgress" name="ProjectProgress" placeholder="Project Progress (%)" onChange={this.handleChange} value={this.state.ProjectProgress}/>
                </FormGroup>
                <FormGroup className="col-1"></FormGroup>
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel + " " + styles.required}>Project Scope</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <Form.Control size="sm" type="text" as="select" id="Scope" name="ProjectScope" placeholder="Project Scope" onChange={this.handleChange} value={this.state.ProjectScope}>
                        <option value="">Select an Option</option>
                    </Form.Control>
                </FormGroup>
            </Form.Row>
            <Form.Row>
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel + " " + styles.required}>Project Schedule</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <Form.Control size="sm" type="text" as="select" id="Schedule" name="ProjectSchedule" placeholder="Project Schedule" onChange={this.handleChange} value={this.state.ProjectSchedule}>
                        <option value="">Select an Option</option>
                    </Form.Control>
                </FormGroup>
                <FormGroup className="col-1"></FormGroup>
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel + " " + styles.required}>Project Resources</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <Form.Control size="sm" type="text" as="select" id="Resource" name="ProjectResource" placeholder="Project Resource" onChange={this.handleChange} value={this.state.ProjectResource}>
                        <option value="">Select an Option</option>
                    </Form.Control>
                </FormGroup>
            </Form.Row>
            <Form.Row>
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel}>Invoiced Amount</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <Form.Control size="sm" type="number" id="InvoicedAmount" name="InvoicedAmount" placeholder="Invoiced Amount" onChange={this.handleChange} value={this.state.InvoicedAmount}/>
                </FormGroup>
                <FormGroup className="col-1"></FormGroup>
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel + " " + styles.required}>Project Cost</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <Form.Control size="sm" type="text" as="select" id="ProjectCost" name="ProjectCost" placeholder="Project Cost" onChange={this.handleChange} value={this.state.ProjectCost}>
                        <option value="">Select an Option</option>
                    </Form.Control>
                </FormGroup>
            </Form.Row>
            <Form.Row className="mb-4">
                <FormGroup className="col-2"> 
                    <Form.Label className={styles.customlabel}>Total Cost</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                    <Form.Control size="sm" type="text" id="TotalCost" name="TotalCost" placeholder="Total Cost" onChange={this.handleChange} value={this.state.TotalCost}/>
                </FormGroup>
                <FormGroup className="col-6"></FormGroup>
            </Form.Row>
            <Form.Row className={styles.buttonCLass}>
                <FormGroup></FormGroup>
                <div>
                <Button id="submit" size="sm" variant="primary" type="submit">
                    Submit
                </Button> 
                </div>  
                <FormGroup className="col-.5"></FormGroup>  
                <div>
                <Button id="cancel" size="sm" variant="primary" onClick={this._closeform}>
                    Cancel
                </Button>
                </div>
                {/* <div>
                <Button id="reset" size="sm" variant="primary" onClick={this.resetform}>
                    Reset
                </Button>
                </div> */}
             </Form.Row>
        </Form> 
    </div>);
    }

    //function to validate the date, not allowing the user to enter end date lesser than start date
    private _validateDate(e){
        let newState = {};
        //validation for date
        if((e.target.name == "ActualStartDate" && e.target.value!="") && (this.state.ProjectProgress=="100")){
            this.setState({
              disable_plannedCompletion: false
            })
            if(this.state.ActualEndDate!=""){
              $('.ActualEndDate').text("");
              var date1 = $('#ActualStartDate').val();
              var date2 = $('#ActualEndDate').val()
              if(date1>=date2){
                $('#ActualEndDate').val("");
                newState[e.target.name] = "";
                this.setState(newState);
                //alert("Planned Completion Cannot be less than Planned Start");
                $('#ActualEndDate').closest('div').append('<span class="ActualEndDate" style="color:red;font-size:9pt">Must be greater than Actual Start date</span>')
              }else{
                $('.ActualEndDate').remove();
                }
            }
            }else if((e.target.name == "ActualStartDate" && e.target.value =="") && (this.state.ProjectProgress=="100") ){
            this.setState({
              ActualEndDate: "",
              //disable_plannedCompletion: true
            })
          }
          if(e.target.name == "ActualEndDate"){
            $('.ActualEndDate').text("");
            var date1 = $('#ActualStartDate').val();
            var date2 = $('#ActualEndDate').val()
            if(date1>=date2){
              $('#ActualEndDate').val("")
              newState[e.target.name] = "";
              this.setState(newState);
              //alert("Planned Completion Cannot be less than Planned Start");
              $('#ActualEndDate').closest('div').append('<span class="ActualEndDate" style="color:red;font-size:9pt">Must be greater than Actual Start date</span>')
            }else{
              $('.ActualEndDate').remove();
            }
          }//validation for date ending
    }

    //function to validate progress
    private _validateProgress(e){
        
        if(e.target.name == "ProjectProgress" && e.target.value!=""){
            e.target.value > 100 ? this.setState({ProjectProgress: "100"}) : e.target.value;
        }
        if(e.target.name == "ProjectProgress" && e.target.value >= 100){
            this.setState({
                disable_plannedCompletion: false,
                ProjectStatus: "Completed"
            })
        }else if(e.target.name == "ProjectProgress" && e.target.value != 100){
            this.setState({
                disable_plannedCompletion: true,
                ActualEndDate: '',
                ProjectStatus: ""
            })
        }

        if(e.target.name == "ProjectStatus" && e.target.value =="Completed"){
            this.setState({
                ProjectProgress: "100",
                disable_plannedCompletion: false
            })
        }else if(e.target.name == "ProjectStatus" && e.target.value !="Completed"){
            this.setState({
                ProjectProgress: (this.state.ProjectProgress == "100" ? "" : this.state.ProjectProgress ),
                ActualEndDate:'',
                disable_plannedCompletion: true
            })
        }
    }

    //function to check if ProjectId already exists or not
    private _checkExistingProjectId(siteColUrl, ProjectIDValue){
        const endPoint: string = `${siteColUrl}/_api/web/lists('`+ listGUID +`')/items?select = ProjectID`;
        let breakCondition = false;
        $('.ProjectID').remove();
        this.props.currentContext.spHttpClient.get(endPoint, SPHttpClient.configurations.v1)
        .then((response: HttpClientResponse) => {
            if (response.ok) {
            response.json()
                .then((jsonResponse) => {
                jsonResponse.value.forEach( item => {
                if(ProjectIDValue == item.ProjectID && !breakCondition){
                    this.setState({
                    ProjectID: ''
                    })
                    $('#ProjectId').closest('div').append('<span class="ProjectID" style="color:red;font-size:9pt">Project Id already Exists</span>');
                    breakCondition = true;
                } 
                // if(ProjectIDValue != item.ProjectID && breakCondition){
                //   $('.ProjectID').remove();
                // }
                
                });
                }, (err: any): void => {
                console.warn(`Failed to fulfill Promise\r\n\t${err}`);
                });
            } else {
            console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
            }
        });
    }
    //fucntion to load items for particular item id on edit form
    private _loadItems(){
    
        var itemId = _getParameterValues('id');
        if(itemId==""){
        alert("Incorrect URL");
        let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
        window.open(winURL,'_self');
        }else{
        const url = this.props.currentContext.pageContext.web.absoluteUrl + `/_api/web/lists('`+ listGUID +`')/items(`+ itemId +`)`;
        return this.props.currentContext.spHttpClient.get(url,SPHttpClient.configurations.v1,  
            {  
                headers: {  
                'Accept': 'application/json;odata=nometadata',  
                'odata-version': ''  
                }  
            }).then((response: SPHttpClientResponse): Promise<SPProjectListEditForm> => {  
                return response.json();  
            })  
            .then((item: SPProjectListEditForm): void => {   
            this.setState({
                ProjectID: item.ProjectID,
                DeliveryManager: item.Delivery_x0020_Manager,
                ProjectName: item.Project_x0020_Name,
                ClientName: item.Client_x0020_Name,
                ProjectManager: item.Project_x0020_Manager,
                ProjectType: item.Project_x0020_Type,
                ProjectMode: item.Project_x0020_Mode,
                PlannedStart: item.PlannedStart,
                PlannedCompletion: item.Planned_x0020_End,
                ActualStartDate: item.Actual_x0020_Start,
                ActualEndDate: item.Actual_x0020_End,
                ProjectDescription: item.Project_x0020_Description,
                ProjectLocation: item.Region,
                RevisedBudget: item.Revised_x0020_Budget,
                ProjectStatus: item.Status,
                TotalCost: item.Total_x0020_Cost,
                InvoicedAmount: item.Invoiced_x0020_amount,
                ProjectScope:item.Scope,
                ProjectSchedule: item.Schedule,
                ProjectResource: item.Resource,
                ProjectCost: item.Project_x0020_Cost,
                ProjectProgress: item.Progress,
                disable_RMSID: true
            })  
            // console.log(this.state.PlannedStart + " " + this.state.PlannedCompletion) ;
            });
        }
        }
    //function to save the edit item
    private _saveItem(e){
        if(this.state.disable_plannedCompletion){
            this.setState({
                ActualEndDate:""
            })
        }
        var itemId = _getParameterValues('id');
        let _validate=0;
        e.preventDefault();
       

        let requestData = {
            __metadata:  
            {  
                type: listType
            },  
            ProjectID : this.state.ProjectID,
            Project_x0020_Name: this.state.ProjectName,
            Client_x0020_Name: this.state.ClientName,
            Delivery_x0020_Manager: this.state.DeliveryManager,
            Project_x0020_Manager: this.state.ProjectManager,
            Project_x0020_Type: this.state.ProjectType,
            Project_x0020_Mode: this.state.ProjectMode,
            Actual_x0020_Start: this.state.ActualStartDate,
            Actual_x0020_End: this.state.ActualEndDate,
            Project_x0020_Description: this.state.ProjectDescription,
            Region: this.state.ProjectLocation,
            Revised_x0020_Budget: this.state.RevisedBudget,
            Status: this.state.ProjectStatus,
            Progress: this.state.ProjectProgress,
            Total_x0020_Cost: this.state.TotalCost,
            Invoiced_x0020_amount: this.state.InvoicedAmount,
            Scope: this.state.ProjectScope,
            Schedule: this.state.ProjectSchedule,
            Resource: this.state.ProjectResource,
            Project_x0020_Cost: this.state.ProjectCost
        
        };
        //validation
        // if (requestData.ProjectID.length < 1){
        //     $('#ProjectId').css('border','1px solid red');
        //     _validate++;
        // }else{
        //     $('#ProjectId').css('border','1px solid #ced4da')
        // }
        //projectId validation
        // if (requestData.ProjectID == null ||  requestData.ProjectID == "" || requestData.ProjectID.length < 1){
        //     $('#ProjectId').css('border','1px solid red');
        //     this._validationMessage("ProjectId", "ProjectID", "Project Id cannot be empty");
        //     _validate++;
        // }else if ((requestData.ProjectID != "" || requestData.ProjectID != null) && requestData.ProjectID == "0"){
        //     //$('.ProjectID').remove();
        //     $('#ProjectId').css('border','1px solid red');
        //     this._validationMessage("ProjectId", "ProjectID", "Project Id cannot be 0");
        //     _validate++;
        // }else{
        //     $('.ProjectID').remove();
        //     $('#ProjectId').css('border','1px solid #ced4da')
        // }
        // if( requestData.Project_x0020_Name.length < 1){
        //     $('#ProjectName').css('border','1px solid red');
        //     _validate++;
        // }else{
        //     $('#ProjectName').css('border','1px solid #ced4da')
        // }
        //delivery manager 
        if(requestData.Delivery_x0020_Manager == null || requestData.Delivery_x0020_Manager == "" ){
            this._validationMessage("DeliveryManager", "DeliveryManager", "Delivery Manager cannot be empty");
            $('#DeliveryManager input').css('border','1px solid red');
            _validate++;
        }else{
            $('.DeliveryManager').remove();
            $('#DeliveryManager input').css('border','1px solid #ced4da');
        }
        //project manager
        if(requestData.Project_x0020_Manager == null || requestData.Project_x0020_Manager == ""){
            this._validationMessage("ProjectManager", "ProjectManager", "Project Manager cannot be empty");
            $('#ProjectManager input').css('border','1px solid red');
            _validate++;
        }else{
            $('.ProjectManager').remove();
            $('#ProjectManager input').css('border','1px solid #ced4da');
        }
        //   //revised project
        // if ((requestData.Revised_x0020_Budget == null || requestData.Revised_x0020_Budget=="")) {
        //     this._validationMessage("RevisedBudget", "RevisedBudget", "Revised Budget cannot be empty");
        //     $('#RevisedBudget').css('border','1px solid red');
        //     _validate++;
        // }else{
        //     $('.RevisedBudget').remove();
        //     $('#RevisedBudget').css('border','1px solid #ced4da')
        // }
        if (requestData.Revised_x0020_Budget == null || requestData.Revised_x0020_Budget.length < 1 || requestData.Revised_x0020_Budget =="") {
            this._validationMessage("RevisedBudget", "RevisedBudget", "Revised Budget cannot be empty");
            $('#RevisedBudget').css('border','1px solid red');
            _validate++;
          }else if ((requestData.Revised_x0020_Budget != "" || requestData.Revised_x0020_Budget != null) && requestData.Revised_x0020_Budget == "0"){
            //$('.ProjectID').remove();
            $('#RevisedBudget').css('border','1px solid red');
            this._validationMessage("RevisedBudget", "RevisedBudget", "Revised Budget cannot be 0");
            _validate++;
          }else{
            $('.RevisedBudget').remove();
            $('#RevisedBudget').css('border','1px solid #ced4da')
        }
        //progress
        if(requestData.Progress.length < 1 || requestData.Progress == null || requestData.Progress =="") {
            this._validationMessage("ProjectProgress", "ProjectProgress", "Project Progress cannot be empty");
            $('#ProjectProgress').css('border','1px solid red');
            _validate++;
          }else if ((requestData.Progress != "" || requestData.Progress != null) && requestData.Progress == "0"){
            //$('.ProjectID').remove();
            $('#ProjectProgress').css('border','1px solid red');
            this._validationMessage("ProjectProgress", "ProjectProgress", "Project Progress cannot be 0");
            _validate++;
          }else{
            $('.ProjectProgress').remove();
            $('#ProjectProgress').css('border','1px solid #ced4da')
        }
        //actual start
        if(requestData.Actual_x0020_Start == null || requestData.Actual_x0020_Start == ""){
            this._validationMessage("ActualStartDate", "ActualStartDate", "Actual Start Date cannot be empty");
            $('#ActualStartDate').css('border','1px solid red');
            _validate++;
        }else{
            $('.ActualStartDate').remove();
            $('#ActualStartDate').css('border','1px solid #ced4da');
        }
        if(requestData.Progress=="100" && (requestData.Actual_x0020_End == null || requestData.Actual_x0020_End == "") ){
            this._validationMessage("ActualEndDate", "ActualEndDate", "Actual End Date cannot be empty");
            $('#ActualEndDate').css('border','1px solid red');
            _validate++;
        }else{
            $('.ActualEndDate').remove();
            $('#ActualEndDate').css('border','1px solid #ced4da');
        }
        if (requestData.Scope == null || requestData.Scope == "") {
            this._validationMessage("Scope", "Scope", "Project Scope cannot be empty");
            $('#Scope').css('border','1px solid red');
            _validate++;
        }else{
            $('.Scope').remove();
            $('#Scope').css('border','1px solid #ced4da')
        }
        if (requestData.Status == null || requestData.Status =="" || requestData.Status.length < 1) {
            this._validationMessage("Status", "Status", "Project Status cannot be empty");
            $('#Status').css('border','1px solid red');
            _validate++;
          }else{
            $('.Status').remove();
            $('#Status').css('border','1px solid #ced4da')
          }
        if (requestData.Schedule == null || requestData.Schedule == "") {
            this._validationMessage("Schedule", "Schedule", "Project Schedule cannot be empty");
            $('#Schedule').css('border','1px solid red');
            _validate++;
        }else{
            $('.Schedule').remove();
            $('#Schedule').css('border','1px solid #ced4da')
        }
        if (requestData.Project_x0020_Cost == null || requestData.Schedule == "") {
            this._validationMessage("ProjectCost", "ProjectCost", "Project Cost cannot be empty");
            $('#ProjectCost').css('border','1px solid red');
            _validate++;
        }else{
            $('.ProjectCost').remove();
            $('#ProjectCost').css('border','1px solid #ced4da')
        }
        if (requestData.Resource == null || requestData.Resource == "") {
            this._validationMessage("Resource", "Resource", "Project Resource cannot be empty");
            $('#Resource').css('border','1px solid red');
            _validate++;
        }else{
            $('.Resource').remove();
            $('#Resource').css('border','1px solid #ced4da')
        }
        if (requestData.Project_x0020_Description == null ||  requestData.Project_x0020_Description =="") {
            this._validationMessage("ProjectDescription", "ProjectDescription", "Project Description cannot be empty");
            $('#ProjectDescription').css('border','1px solid red');
            _validate++;
          }else{
              $('.ProjectDescription').remove();
            $('#ProjectDescription').css('border','1px solid #ced4da')
          }
        if(_validate>0){
            return false;
        }
        
        
        $.ajax({
            url:  this.props.currentContext.pageContext.web.absoluteUrl+ "/_api/web/lists('" + listGUID + "')/items("+ itemId +")",  
            type: "POST",  
            data: JSON.stringify(requestData),  
            headers:  
            {  
                "Accept": "application/json;odata=verbose",  
                "Content-Type": "application/json;odata=verbose",  
                "X-RequestDigest": this.state.FormDigestValue,
                "IF-MATCH": "*",
                'X-HTTP-Method': 'MERGE' 
            },  
            success:(data, status, xhr) => 
            {  
                alert("Submitted successfully");
                let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
                window.open(winURL,'_self');
            },  
            error: (xhr, status, error)=>
            {  
                alert(JSON.stringify(xhr.responseText));
                let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
                window.open(winURL,'_self');
            }  
        });
        
        this.setState({
            ProjectID: '',
            ProjectName: '',
            ClientName: '',
            DeliveryManager:'',
            ProjectManager: '',
            ProjectType: '',
            ProjectMode: '',
            ProjectDescription: '',
            ProjectLocation: '',
            ProjectStatus: '',
            ActualStartDate:'',
            ActualEndDate:'',
            RevisedBudget:'',
            TotalCost:'',
            InvoicedAmount: '',
            ProjectScope:'',
            ProjectSchedule: '',
            ProjectResource: '',
            ProjectCost: '',
            startDate: '',
            disable_plannedCompletion:true,
            endDate: '',
            focusedInput: '',
            FormDigestValue:''
        });
    }

    private _validationMessage(_id, _classname, _message){
        $('.' + _classname).remove();
        $('#' + _id).closest('div').append('<span class="' + _classname + '" style="color:red;font-size:9pt">'+ _message +'</span>');
      }
    //function to keep the request digest token active
    private _getAccessToken(){
    
    $.ajax({  
        url: this.props.currentContext.pageContext.web.absoluteUrl+"/_api/contextinfo",  
        type: "POST",  
        headers:{'Accept': 'application/json; odata=verbose;', "Content-Type": "application/json;odata=verbose",  
        },  
        success: (resultData)=> {  
        
        this.setState({  
            FormDigestValue: resultData.d.GetContextWebInformation.FormDigestValue
        });  
        },  
        error : (jqXHR, textStatus, errorThrown) =>{  
        }  
    });  
    }
    //function to close the form and redirect to the Grid page
    private _closeform(e){
      e.preventDefault();
    let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
    // this.setState({
    //     ProjectID : '',
    //     ProjectName: '',
    //     ClientName: '',
    //     DeliveryManager:'',
    //     ProjectManager: '',
    //     ProjectType: '',
    //     ProjectMode: '',
    //     //   PlannedStart: '',
    //     //   PlannedCompletion: '',
    //     ProjectDescription: '',
    //     ProjectLocation: '',
    //     //   ProjectBudget: '',
    //     ProjectStatus: '',
    //     ProjectProgress:'',
    //     ActualStartDate:'',
    //     ActualEndDate:'',
    //     RevisedBudget:'',
    //     TotalCost:'',
    //     InvoicedAmount: '',
    //     ProjectScope:'',
    //     ProjectSchedule: '',
    //     ProjectResource: '',
    //     ProjectCost: '',
    //     startDate: '',
    //     endDate: '',
    //     focusedInput: '',
    //     FormDigestValue:''
    // });
    window.open(winURL,'_self');
    }
    //function to load choice column values in the dropdown
    private _retrieveAllChoicesFromListField(siteColUrl: string, columnName: string): void {
        const endPoint: string = `${siteColUrl}/_api/web/lists('`+ listGUID +`')/fields?$filter=EntityPropertyName eq '`+ columnName +`'`;
      
        this.props.currentContext.spHttpClient.get(endPoint, SPHttpClient.configurations.v1)
          .then((response: HttpClientResponse) => {
            if (response.ok) {
              response.json()
                .then((jsonResponse) => {
                  console.log(jsonResponse.value[0]);
                  let dropdownId = jsonResponse.value[0].Title.replace(/\s/g, '');
                  jsonResponse.value[0].Choices.forEach(dropdownValue => {
                    $('#' + dropdownId ).append('<option value="'+ dropdownValue +'">'+ dropdownValue +'</option>');
                  });
                }, (err: any): void => {
                  console.warn(`Failed to fulfill Promise\r\n\t${err}`);
                });
            } else {
              console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
            }
          });
      }


}