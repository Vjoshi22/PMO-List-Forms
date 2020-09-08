import * as React from 'react';
import styles from './PmoListForms.module.scss';
import { IPmoListFormsProps } from './IPmoListFormsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient,HttpClient, IHttpClientOptions, HttpClientResponse, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse } from "@microsoft/sp-http";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { _getParameterValues } from './getQueryString';
import { Form, FormGroup, Button, FormControl } from "react-bootstrap";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { SPProjectList } from "../components/IProjectListProps";
import * as $ from "jquery";
import { _getListEntityName, listType } from './getListEntityName';
import { data } from 'jquery';
import { _logExceptionError } from '../../../ExceptionLogging';
//import json_RMSData1 from "../../../../data/data.json";
import { sp } from "@pnp/sp";
import "@pnp/sp/profiles";
import { Field } from 'sp-pnp-js';


export var allchoiceColumns: any[] = ["Project_x0020_Type", "Project_x0020_Mode", "Status", "Project_x0020_Phase", "Region"];
export var inputfieldLength = 50;
var PM_userInfo;
var DM_userInfo;
require('./PmoListForms.module.scss');
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");

export interface IreactState {
  ProjectID: string,
  CRM_Id: string,
  ProjectName: string;
  ClientName: string;
  ProjectManager: string;
  ProjectType: string;
  ProjectMode: string;
  ProjectPhase: string;
  PlannedStart: string;
  PlannedCompletion: string;
  ProjectDescription: string;
  ProjectLocation: string;
  ProjectBudget: number;
  ProjectStatus: string;
  ProjectProgress: number;
  TotalCost: number;
  //peoplepicker
  DeliveryManager: string;
  PM:number;
  DM:number;
  //date
  startDate: any;
  disable_RMSID: boolean;
  disable_plannedCompletion: boolean;
  endDate: any;
  focusedInput: any;
  FormDigestValue: string;
  RMSData:{};
}

var listGUID: any = "2c3ffd4e-1b73-4623-898d-8e3a1bb60b91";   //"47272d1e-57d9-447e-9cfd-4cff76241a93"; 
var timerID;
let breakCondition = false;
//var newitem: boolean;

export default class PmoListForms extends React.Component<IPmoListFormsProps, IreactState> {
  constructor(props: IPmoListFormsProps, state: IreactState) {
    super(props);

    this.state = {
      //status: 'Ready',  
      //items: []
      ProjectID: '',
      CRM_Id: '',
      ProjectName: '',
      ClientName: '',
      ProjectManager: '',
      ProjectType: '',
      ProjectMode: '',
      ProjectPhase: '',
      PlannedStart: '',
      PlannedCompletion: '',
      ProjectDescription: '',
      ProjectLocation: '',
      ProjectBudget: 0,
      ProjectProgress: 0,
      TotalCost:0,
      ProjectStatus: '',
      DeliveryManager: '',
      PM:0,
      DM:0,
      startDate: '',
      endDate: '',
      disable_RMSID: false,
      disable_plannedCompletion: true,
      focusedInput: '',
      FormDigestValue: '',
      RMSData:{}
    };
    this._getdropdownValues = this._getdropdownValues.bind(this);
    this.handleChange = this.handleChange.bind(this);
    this._getProjectManager = this._getProjectManager.bind(this);
    this._getDeliveryManager = this._getDeliveryManager.bind(this);
    //this.loadItems = this.loadItems.bind(this);
    //this.isOutsideRange = this.isOutsideRange.bind(this);
  }
  public componentDidMount() {
    $('.webPartContainer').hide();
    allchoiceColumns.forEach(elem => {
      this.retrieveAllChoicesFromListField(this.props.currentContext.pageContext.web.absoluteUrl, elem);
    })
    _getListEntityName(this.props.currentContext, this.props.listGUID);
    $('.pickerText_4fe0caaf').css('border', '0px');
    $('.pickerInput_4fe0caaf').addClass('form-control');
    $('.form-row').css('justify-content', 'center');

    // if((/edit/.test(window.location.href))){
    //   newitem = false;
    //   this.loadItems();
    // }
    // if((/new/.test(window.location.href))){
    //   newitem = true
    // }
    if (!this.state.PlannedStart) {
      this.setState({
        disable_plannedCompletion: false
      })
    }
    this.getAccessToken();
    timerID = setInterval(
      () => this.getAccessToken(), 300000);
  }
  public componentWillUnmount() {
    clearInterval(timerID);

  }
  //public  isOutsideRange = day =>day.isAfter(Moment()) || day.isBefore(Moment().subtract(0, "days"));  
  private handleChange = (e) => {
    let newState = {};
    newState[e.target.name] = e.target.value;
    this.setState(newState);

    //functin to check the existing Id
    let _value = e.target.value;
    if (e.target.name == "ProjectID" && (e.target.value.trim() != 0 || e.target.value.trim() != "") && _value.trim().match(/^[a-zA-Z1-9][A-Za-z0-9_]*$/) != null) {
      this._checkExistingProjectId(this.props.currentContext.pageContext.web.absoluteUrl, e.target.value.trim().toLowerCase());
    } else if (e.target.name == "ProjectID" && (_value.trim().match(/^[a-zA-Z1-9][A-Za-z0-9_]*$/) == null || _value.trim() == "")) {
      $('.ProjectID').remove();
      $('#ProjectId').closest('div').append('<span class="ProjectID" style="color:red;font-size:9pt">Cannot start with 0 or special charachters</span>');
    }
    this.validateDate(e);
    this._validateProgress(e);
  }
  private handleSubmit = (e) => {

    this.createItem(e);
    // if(newitem){
    //   this.createItem(e);
    // }else{
    //   this.saveItem(e);
    // }
  }
  private _getProjectManager = (items: any[]) => {
    console.log('Items:', items);
    this.setState({ 
      ProjectManager: items[0].text,
      PM: items[0].id
    });
  }
  private _getDeliveryManager = (items: any[]) => {
    console.log('Items:', items);
    this.setState({ 
      DeliveryManager: items[0].text,
      DM: items[0].id
     });
  }
  private _getdropdownValues(e) {
    // this.retrieveAllChoicesFromListField(e);
  }

  public render(): React.ReactElement<IPmoListFormsProps> {

    SPComponentLoader.loadCss("//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css");
    // SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datetimepicker/4.7.14/js/bootstrap-datetimepicker.min.js");
    // SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.4.1/css/bootstrap-datepicker3.css");

    return (
      <div id="newItemDiv" className={styles["_main-div"]} >
        <div id="heading" className={styles.heading}><h3>Project Details</h3></div>
        <Form onSubmit={this.handleSubmit}>
          <Form.Row className="mt-3">
            {/*-----------RMS ID------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Project Id</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" maxLength={inputfieldLength} type="text" disabled={this.state.disable_RMSID} id="ProjectId" name="ProjectID" placeholder="Project ID" onChange={this.handleChange} onBlur={() => {this._getRMSData()}} value={this.state.ProjectID} />
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            {/*-----------Project Type------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Project Type</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" id="ProjectType" as="select" name="ProjectType" onClick={() => this._getdropdownValues} onChange={this.handleChange} value={this.state.ProjectType}>
                <option value="">Select an Option</option>
              </Form.Control>
            </FormGroup>
          </Form.Row>

          <Form.Row>
            {/* -----------Client Name------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Client Name</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              {/* <Form.Control size="sm" maxLength={inputfieldLength} type="text" id="ClientName" name="ClientName" placeholder="Client Name" onChange={this.handleChange} value={this.state.ClientName} /> */}
              <Form.Label>{this.state.ClientName}</Form.Label>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            {/* -----------Project Name---------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Project Name</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              {/* <Form.Control size="sm" maxLength={inputfieldLength} type="text" id="ProjectName" name="ProjectName" placeholder="Ex: John Deer" onChange={this.handleChange} value={this.state.ProjectName} /> */}
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
                  disabled={true}
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
              <Form.Label className={styles.customlabel + " " + styles.required}>Project Manager</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <div id="ProjectManager">
                <PeoplePicker
                  context={this.props.currentContext}
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users    
                  showtooltip={true}
                  isRequired={true}
                  disabled={true}
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
              <Form.Label className={styles.customlabel + " " + styles.required}>Project Mode</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              {/* <Form.Control size="sm" id="ProjectMode" as="select" name="ProjectMode" onChange={this.handleChange} value={this.state.ProjectMode}>
                <option value="">Select an Option</option>
              </Form.Control> */}
              <Form.Label>{this.state.ProjectMode}</Form.Label>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Project Status</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" id="Status" as="select" name="ProjectStatus" onChange={this.handleChange} value={this.state.ProjectStatus}>
                <option value="">Select an Option</option>
                {/* <option value="In progress">In progress</option>
              <option value="Initiated">Initiated</option>
              <option value="Closed">Closed</option>
              <option value="Withdrawn">Withdrawn</option> */}
              </Form.Control>
            </FormGroup>
          </Form.Row>
          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Project Phase</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" id="ProjectPhase" as="select" name="ProjectPhase" onChange={this.handleChange} value={this.state.ProjectPhase}>
                <option value="">Select an Option</option>
              </Form.Control>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel}>Region</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              {/* <Form.Control size="sm" as="select" id="Region" name="ProjectLocation" placeholder="Project Location" onChange={this.handleChange} value={this.state.ProjectLocation}>
                <option value="">Select an Option</option>
              </Form.Control> */}
              <Form.Label>{this.state.ProjectLocation}</Form.Label>
            </FormGroup>
          </Form.Row>
          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Planned Start Date</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              {/* <Form.Control size="sm" type="date" id="PlannedStart" name="PlannedStart" placeholder="Planned Start Date" onChange={this.handleChange} value={this.state.PlannedStart} /> */}
              {/* <DatePicker selected={this.state.PlannedStart}  onChange={this.handleChange} />; */}
            <Form.Label>{this.state.PlannedStart}</Form.Label>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Planned End Date</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              {/* <Form.Control size="sm" type="date" disabled={this.state.disable_plannedCompletion} id="PlannedCompletion" name="PlannedCompletion" placeholder="Planned Completion Date" onChange={this.handleChange} value={this.state.PlannedCompletion} /> */}
              <Form.Label>{this.state.PlannedCompletion}</Form.Label>
            </FormGroup>
          </Form.Row>
          {/* Project Description */}
          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Project Description</Form.Label>
            </FormGroup>
            <FormGroup className="col-9 mb-3">
              <Form.Control size="sm" as="textarea" maxLength={inputfieldLength} rows={4} type="text" id="ProjectDescription" name="ProjectDescription" placeholder="Project Description" onChange={this.handleChange} value={this.state.ProjectDescription} />
            </FormGroup>
          </Form.Row>
          {/* Next Row */}
          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Project Progress</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="number" id="ProjectProgress" name="ProjectProgress" placeholder="Project Progress (%)" onChange={this.handleChange} value={this.state.ProjectProgress} />
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Budget as per SOW</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" maxLength={inputfieldLength} type="number" id="BudgetSOW" name="ProjectBudget" placeholder="Project Budget" onChange={this.handleChange} value={this.state.ProjectBudget} />
            </FormGroup>
          </Form.Row>
          {/* <Form.Row className="mb-4">
          <FormGroup className="col-2"> 
            <Form.Label className={styles.customlabel + " " + styles.required}>Project Progress</Form.Label>
          </FormGroup>
          <FormGroup className="col-3">
            <Form.Control size="sm" type="number" id="ProjectProgress" name="ProjectProgress" placeholder="Project Progress (%)" onChange={this.handleChange} value={this.state.ProjectProgress}/>
          </FormGroup>
          <FormGroup className="col-6">
          </FormGroup>
        </Form.Row> */}
          <Form.Row className={styles.buttonCLass}>
            <FormGroup></FormGroup>
            <div>
              <Button id="submit" size="sm" variant="primary" type="submit">
                Submit
              </Button>
            </div>
            <FormGroup className="col-.5"></FormGroup>
            <div>
              <Button id="cancel" size="sm" variant="primary" onClick={() => { this.closeform() }}>
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
  //load data form RMS
  private async _getRMSData(): Promise<void>{

    var apiURL = "https://rms.yash.com/rms/projects/projectAttributePMO?find=PMOProjectAttribute&projectId=" + this.state.ProjectID;   
    const myOptions: IHttpClientOptions = {
      headers: new Headers({
        'Authorization': 'Basic YWRtaW46YWRtaW4xMjM0NQ=='
      }),
      method: 'GET'
    };  
    return this.props.currentContext.httpClient.get(apiURL, HttpClient.configurations.v1, myOptions)
          .then((apiResponse: HttpClientResponse) => {
            console.log(apiResponse);
            return apiResponse.json();
          }).then(json_RMSData => {


    if(json_RMSData.status && this.state.ProjectID == json_RMSData.data.projectId){
      this.setState({
        ProjectManager: json_RMSData.data.manager,
        DeliveryManager: json_RMSData.data.deliveryManager,
        ClientName: json_RMSData.data.clientName,
        ProjectName: json_RMSData.data.projectName,
        PlannedStart: json_RMSData.data.projectStartDate,
        PlannedCompletion: json_RMSData.data.projectEndDate,
        startDate: json_RMSData.data.projectStartDate,
        endDate: json_RMSData.data.projectEndDate,
        ProjectMode: json_RMSData.data.projectMode,
        ProjectLocation: json_RMSData.data.region,
        TotalCost: json_RMSData.data.totalCost
      });
      this._getProjectManagerProperties(json_RMSData.data.manager);
      this._getDeliveryManagerProperties(json_RMSData.data.deliveryManager);

    }else if((this.state.ProjectID!="" || this.state.ProjectID != json_RMSData.data.projectId) && !json_RMSData.status){
      this.setState({
        ProjectID:'',
        ProjectManager:'',
        DeliveryManager:'',
        ClientName:'',
        ProjectName: '',
        PlannedStart: '',
        PlannedCompletion: '',
        startDate:'',
        endDate:'',
        ProjectMode: '',
        ProjectLocation: '',
        TotalCost:0
      })
      $('#ProjectId').css('border', '1px solid red');
      this._validationMessage("ProjectId", "ProjectID", json_RMSData.message);
    }
    // else if(!json_RMSData.status){
    //   alert("RMS System is down, Please wait for sometime");
    // }
  });
  }
  //get the userProfile Properties
  private _getProjectManagerProperties(userName){
       /// username should be passed as 'domain\username'
        /// change this prefix according to the environment. 
        /// In below sample, windows authentication is considered.
        var prefix = "i:0#.f|membership|";
        /// get the site url
        var siteUrl = this.props.currentContext.pageContext.web.absoluteUrl;
        /// add prefix, this needs to be changed based on scenario
        var accountName = prefix + userName;

        /// make an ajax call to get the site user
        $.ajax({
            url: siteUrl + "/_api/web/siteusers(@v)?@v='" + 
                encodeURIComponent(accountName) + "'",
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" },
            success: function (data) {
              //user id received from the site 
              PM_userInfo = data.d;
            },
            error: function (data) {
                console.log(JSON.stringify(data));
            }
        }).then(p => {
          this.setState({
            PM: PM_userInfo.Id,
            ProjectManager: PM_userInfo.Title
          })
        });
  }
  //get the userProfile Properties
  private _getDeliveryManagerProperties(userName){
    /// username should be passed as 'domain\username'
     /// change this prefix according to the environment. 
     /// In below sample, windows authentication is considered.
     var prefix = "i:0#.f|membership|";
     /// get the site url
     var siteUrl = this.props.currentContext.pageContext.web.absoluteUrl;
     /// add prefix, this needs to be changed based on scenario
     var accountName = prefix + userName;

     /// make an ajax call to get the site user
     $.ajax({
         url: siteUrl + "/_api/web/siteusers(@v)?@v='" + 
             encodeURIComponent(accountName) + "'",
         method: "GET",
         headers: { "Accept": "application/json; odata=verbose" },
         success: function (data) {
           //user id received from the site 
           DM_userInfo = data.d;
         },
         error: function (data) {
             console.log(JSON.stringify(data));
         }
     }).then(p => {
       this.setState({
         DM: DM_userInfo.Id,
         DeliveryManager: DM_userInfo.Title
       })
     });
}
  //function to validate the date, end date should not be less than start date
  private validateDate(e) {
    let newState = {};
    //validation for date
    if (e.target.name == "PlannedStart" && e.target.value != "") {
      this.setState({
        disable_plannedCompletion: false
      })
      if (this.state.PlannedCompletion != "") {
        $('.PlannedCompletion').text("");
        var date1 = $('#PlannedStart').val();
        var date2 = $('#PlannedCompletion').val()
        if (date1 >= date2) {
          $('#PlannedCompletion').val("")
          newState[e.target.name] = "";
          this.setState(newState);
          //alert("Planned Completion Cannot be less than Planned Start");
          $('#PlannedCompletion').closest('div').append('<span class="PlannedCompletion" style="color:red;font-size:9pt">Must be greater than Planned Start date</span>')
        } else {
          $('.PlannedCompletion').remove();
        }
      }
    } else if (e.target.name == "PlannedStart" && e.target.value == "") {
      this.setState({
        PlannedCompletion: "",
        disable_plannedCompletion: true
      })
    }
    if (e.target.name == "PlannedCompletion") {
      $('.PlannedCompletion').text("");
      var date1 = $('#PlannedStart').val();
      var date2 = $('#PlannedCompletion').val()
      if (date1 >= date2) {
        $('#PlannedCompletion').val("")
        newState[e.target.name] = "";
        this.setState(newState);
        //alert("Planned Completion Cannot be less than Planned Start");
        $('#PlannedCompletion').closest('div').append('<span class="PlannedCompletion" style="color:red;font-size:9pt">Must be greater than Planned Start date</span>')
      } else {
        $('.PlannedCompletion').remove();
      }
    }//validation for date ending
  }
  //Validate  Progress
  //function to validate progress
  private _validateProgress(e) {
    if (e.target.name == "ProjectProgress" && e.target.value != "") {
      e.target.value > 100 ? this.setState({ ProjectProgress: 100 }) : e.target.value;
    }
    if (e.target.name == "ProjectProgress" && e.target.value >= 100) {
      this.setState({
        disable_plannedCompletion: false,
        ProjectStatus: "Completed"
      })
    } else if (e.target.name == "ProjectProgress" && e.target.value != 100) {
      this.setState({
        //disable_plannedCompletion: true
        // PlannedCompletion: '',
        //ProjectStatus: ""
      })
    }

    if (e.target.name == "ProjectStatus" && e.target.value == "Completed") {
      this.setState({
        ProjectProgress: 100
        // disable_plannedCompletion: false
      })
    } else if (e.target.name == "ProjectStatus" && e.target.value != "Completed") {
      this.setState({
        ProjectProgress: (this.state.ProjectProgress == 100 ? 0 : this.state.ProjectProgress)
        // PlannedCompletion:'',
        // disable_plannedCompletion: true
      })
    }
  }
  //function to check if ProjectId already exists or not
  private _checkExistingProjectId(siteColUrl, ProjectIDValue) {
    let _formdigest = this.state.FormDigestValue; //variable for errorlog function
    let _projectID = this.state.ProjectID; //variable for errorlog function
    

    //const endPoint: string = `${siteColUrl}/_api/web/lists('` + listGUID + `')/items?select = ProjectID`;
    const endPoint: string = `${siteColUrl}/_api/web/lists('` + this.props.listGUID + `')/items?Select=ID&$filter=ProjectID eq '${ProjectIDValue}'`;
    $('.ProjectID').remove();
    this.props.currentContext.spHttpClient.get(endPoint, SPHttpClient.configurations.v1)
      .then((response: HttpClientResponse) => {
        if (response.ok) {
          response.json()
            .then((jsonResponse) => {
              if (jsonResponse.value.length > 0) {
              jsonResponse.value.forEach(item => {
                if (ProjectIDValue == item.ProjectID.toLowerCase()) {
                  // this.setState({
                  //   ProjectID: ''
                  // })
                  $('#ProjectId').closest('div').append('<span class="ProjectID" style="color:red;font-size:9pt">Project Id already Exists</span>');
                  breakCondition = true;
                }else{
                  breakCondition = false;
                }
                // if(ProjectIDValue != item.ProjectID && breakCondition){
                //   $('.ProjectID').remove();
                // }

              });
            }else{
              breakCondition = false;
            }
            }, (err: any): void => {
              _logExceptionError(this.props.currentContext, this.props.exceptionLogGUID,  _formdigest, "inside _checkExistingProjectId pmonewitemform: errlog", "PMOListForms", "_checkExistingProjectId", err, _projectID);
              console.warn(`Failed to fulfill Promise\r\n\t${err}`);
            });
        } else {
          console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
        }
      });
  }
  //fucntion to save the new entry in the list
  private createItem(e) {

    let _formdigest = this.state.FormDigestValue; //variable for errorlog function
    let _projectID = this.state.ProjectID; //variable for errorlog function

    let _validate = 0;
    e.preventDefault();

    let requestData = {
      __metadata:
      {
        type: listType
      },
      ProjectID: this.state.ProjectID,
      Project_x0020_Name: this.state.ProjectName,
      Client_x0020_Name: this.state.ClientName,
      Delivery_x0020_Manager: this.state.DeliveryManager,
      Project_x0020_Manager: this.state.ProjectManager,
      Project_x0020_Type: this.state.ProjectType,
      Project_x0020_Mode: this.state.ProjectMode,
      Project_x0020_Phase: this.state.ProjectPhase,
      PlannedStart: this.state.PlannedStart,
      Planned_x0020_End: this.state.PlannedCompletion,
      Actual_x0020_Start: this.state.startDate,
      Actual_x0020_End: this.state.endDate,
      Project_x0020_Description: this.state.ProjectDescription,
      Region: this.state.ProjectLocation,
      Total_x0020_Cost: this.state.TotalCost,
      Project_x0020_Budget: this.state.ProjectBudget,
      Status: this.state.ProjectStatus,
      Progress: this.state.ProjectProgress,
      PMId: this.state.PM,
      DMId:this.state.DM
    };

    //validation
    //projectId validation
    if (requestData.ProjectID.length < 1 || requestData.ProjectID == null || requestData.ProjectID == "") {
      $('#ProjectId').css('border', '1px solid red');
      this._validationMessage("ProjectId", "ProjectID", "Project Id cannot be empty");
      _validate++;
    } else if ((requestData.ProjectID != "" || requestData.ProjectID != null) && this.state.ProjectID.match(/^[a-zA-Z1-9][A-Za-z0-9_]*$/) == null) {
      //$('.ProjectID').remove();
      $('#ProjectId').css('border', '1px solid red');
      this._validationMessage("ProjectId", "ProjectID", "Cannot start with 0 or special charachters");
      _validate++;
    } else if(breakCondition){
      $('#ProjectId').css('border', '1px solid red');
      this._validationMessage("ProjectId", "ProjectID", "Project Id already Exists");
      _validate++;
    } else {
      $('.ProjectID').remove();
      $('#ProjectId').css('border', '1px solid #ced4da')
    }
    // if (requestData.Client_x0020_Name.length < 1 || requestData.Client_x0020_Name == null || requestData.Client_x0020_Name == "") {
    //   this._validationMessage("ClientName", "ClientName", "Client Name cannot be empty");
    //   $('#ClientName').css('border', '1px solid red');
    //   _validate++;
    // } else {
    //   $('.ClientName').remove();
    //   $('#ClientName').css('border', '1px solid #ced4da')
    // }
    // if (requestData.Project_x0020_Name.length < 1) {
    //   this._validationMessage("ProjectName", "ProjectName", "Project Name cannot be empty");
    //   $('#ProjectName').css('border', '1px solid red');
    //   _validate++;
    // } else {
    //   $('.ProjectName').remove();
    //   $('#ProjectName').css('border', '1px solid #ced4da');
    // }
    if (requestData.Delivery_x0020_Manager.length < 1) {
      this._validationMessage("DeliveryManager", "DeliveryManager", "Delivery Manager cannot be empty");
      $('#DeliveryManager input').css('border', '1px solid red');
      _validate++;
    } else {
      $('.DeliveryManager').remove();
      $('#DeliveryManager input').css('border', '1px solid #ced4da');
    }
    if (requestData.Project_x0020_Manager.length < 1) {
      this._validationMessage("ProjectManager", "ProjectManager", "Project Manager cannot be empty");
      $('#ProjectManager input').css('border', '1px solid red');
      _validate++;
    } else {
      $('.ProjectManager').remove();
      $('#ProjectManager input').css('border', '1px solid #ced4da');
    }
    if (requestData.Project_x0020_Type.length < 1 || requestData.Project_x0020_Type == null || requestData.Project_x0020_Type == "") {
      this._validationMessage("ProjectType", "ProjectType", "Project Type cannot be empty");
      $('#ProjectType').css('border', '1px solid red');
      _validate++;
    } else {
      $('.ProjectType').remove();
      $('#ProjectType').css('border', '1px solid #ced4da')
    }
    // if (requestData.PlannedStart.length < 1 || requestData.PlannedStart == null || requestData.PlannedStart == "") {
    //   this._validationMessage("PlannedStart", "PlannedStart", "Planned Start Date cannot be empty");
    //   $('#PlannedStart').css('border', '1px solid red');
    //   _validate++;
    // } else {
    //   $('.PlannedStart').remove();
    //   $('#PlannedStart').css('border', '1px solid #ced4da');
    // }
    // if (requestData.Planned_x0020_End.length < 1 || requestData.Planned_x0020_End == null || requestData.Planned_x0020_End == "") {
    //   this._validationMessage("PlannedCompletion", "PlannedCompletion", "Planned End Date cannot be empty");
    //   $('#PlannedCompletion').css('border', '1px solid red');
    //   _validate++;
    // } else {
    //   $('.PlannedCompletion').remove();
    //   $('#PlannedCompletion').css('border', '1px solid #ced4da');
    // }
    // if (requestData.Project_x0020_Mode.length < 1 || requestData.Project_x0020_Mode == null || requestData.Project_x0020_Mode == "") {
    //   this._validationMessage("ProjectMode", "ProjectMode", "Project Mode cannot be empty");
    //   $('#ProjectMode').css('border', '1px solid red');
    //   _validate++;
    // } else {
    //   $('.ProjectMode').remove();
    //   $('#ProjectMode').css('border', '1px solid #ced4da')
    // }
    //Project Status
    if (requestData.Status.length < 1 || requestData.Status == null || requestData.Status == "") {
      this._validationMessage("Status", "Status", "Project Status cannot be empty");
      $('#Status').css('border', '1px solid red');
      _validate++;
    } else if ((requestData.Progress != null) && requestData.Progress < 100 && requestData.Status == "Completed") {
      this._validationMessage("Status", "Status", "Status cannot be Completed, if Project Progress is less than 100");
      _validate++;
    } else {
      $('.Status').remove();
      $('#Status').css('border', '1px solid #ced4da')
    }
    //Project Phase
    if (requestData.Project_x0020_Phase.length < 1 || requestData.Project_x0020_Phase == null || requestData.Project_x0020_Phase == "") {
      this._validationMessage("ProjectPhase", "ProjectPhase", "Project Phase cannot be empty");
      $('#ProjectPhase').css('border', '1px solid red');
      _validate++;
    } else {
      $('.ProjectPhase').remove();
      $('#ProjectPhase').css('border', '1px solid #ced4da')
    }
    //Project Region
    // if (requestData.Region.length < 1 || requestData.Region == null || requestData.Region == "") {
    //   this._validationMessage("Region", "Region", "Project Region cannot be empty");
    //   $('#Region').css('border', '1px solid red');
    //   _validate++;
    // } else {
    //   $('.Region').remove();
    //   $('#Region').css('border', '1px solid #ced4da')
    // }
    // if (requestData.Delivery_x0020_Manager.length < 1 || requestData.Delivery_x0020_Manager == null || requestData.Delivery_x0020_Manager =="") {
    //   $('#DeliveryManager').css('border','2px solid red');
    //   _validate++;
    // }else{
    //   $('#DeliveryManager').css('border','1px solid #ced4da')
    // }
    if (this.state.ProjectBudget.toLocaleString().length == 0) {
      this._validationMessage("BudgetSOW", "BudgetSOW", "Project Budget cannot be empty");
      $('#BudgetSOW').css('border', '1px solid red');
      _validate++;
    } 
    // else if ((requestData.Project_x0020_Budget != null) && requestData.Project_x0020_Budget == 0) {
    //   //$('.ProjectID').remove();
    //   $('#BudgetSOW').css('border', '1px solid red');
    //   this._validationMessage("BudgetSOW", "BudgetSOW", "Budget as per SOW cannot be 0");
    //   _validate++;
    // } 
    else if((requestData.Project_x0020_Budget !=null && requestData.Project_x0020_Budget < 0)) {
      this._validationMessage("BudgetSOW", "BudgetSOW", "Budget as per SOW cannot be less than 0");
      _validate++;
    }else{
      $('.BudgetSOW').remove();
      $('#BudgetSOW').css('border', '1px solid #ced4da');
    }
    
    if (this.state.ProjectProgress.toLocaleString().length == 0) {
      this._validationMessage("ProjectProgress", "ProjectProgress", "Project Progress cannot be empty");
      $('#ProjectProgress').css('border', '1px solid red');
      _validate++;
    } else if ((requestData.Progress != null) && requestData.Progress < 100 && requestData.Status == "Completed") {
      this._validationMessage("Status", "Status", "Status cannot be Completed, if Project Progress is less than 100");
      _validate++;

    } else{
      $('.ProjectProgress').remove();
      $('#ProjectProgress').css('border', '1px solid #ced4da');
    }
    if ((requestData.Progress != null) && requestData.Progress < 0) {
      //$('.ProjectID').remove();
      $('#ProjectProgress').css('border', '1px solid red');
      this._validationMessage("ProjectProgress", "ProjectProgress", "Project Progress cannot be less than 0");
      _validate++;
    }
    else {
      $('.ProjectProgress').remove();
      $('#ProjectProgress').css('border', '1px solid #ced4da');
    }
    if (requestData.Project_x0020_Description.length < 1 || requestData.Project_x0020_Description == null || requestData.Project_x0020_Description == "") {
      this._validationMessage("ProjectDescription", "ProjectDescription", "Project Description cannot be empty");
      $('#ProjectDescription').css('border', '1px solid red');
      _validate++;
    } else {
      $('.ProjectDescription').remove();
      $('#ProjectDescription').css('border', '1px solid #ced4da')
    }
    if (_validate > 0) {
      return false;
    }

    $.ajax({
      url: this.props.currentContext.pageContext.web.absoluteUrl + "/_api/web/lists('" + this.props.listGUID + "')/items",
      type: "POST",
      data: JSON.stringify(requestData),
      headers:
      {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": this.state.FormDigestValue,
        "IF-MATCH": "*",
        'X-HTTP-Method': 'POST'
      },
      success: (data, status, xhr) => {
        alert("Submitted successfully");
        {if(this.props.customGridRequired){
          let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + "/SitePages/Project-Master.aspx";
        window.open(winUrl, '_self');
      }else{
        let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
        window.open(winUrl, '_self');
      }}
        // let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
        // window.open(winURL, '_self');
      },
      error: (xhr, status, error) => {
        _logExceptionError(this.props.currentContext, this.props.exceptionLogGUID, _formdigest, "inside createItem pmonewitemform: errlog", "PMOListForm", "createItems", xhr, _projectID);
        if (xhr.responseText.match('2130575169')) {
          alert("The Project Id you entered already exists, please try with a new Project Id")
        }else if (xhr.responseText.match('2147024891')) {
          alert("You don't have permission to Create a new Project");
        }else{
          alert(JSON.stringify(xhr.responseText));
        }
        {if(this.props.customGridRequired){
          let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + "/SitePages/Project-Master.aspx";
        window.open(winUrl, '_self');
      }else{
        let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
        window.open(winUrl, '_self');
      }}
        //let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
        //window.open(winURL,'_self');
        //location.reload();
      }
    });
  }
  private _validationMessage(_id, _classname, _message) {
    $('.' + _classname).remove();
    $('#' + _id).closest('div').append('<span class="' + _classname + '" style="color:red;font-size:9pt">' + _message + '</span>');
  }
  //function to keep the request digest token active
  private getAccessToken() {
    let _formdigest = this.state.FormDigestValue; //variable for errorlog function
    let _projectID = this.state.ProjectID; //variable for errorlog function

    $.ajax({
      url: this.props.currentContext.pageContext.web.absoluteUrl + "/_api/contextinfo",
      type: "POST",
      headers: {
        'Accept': 'application/json; odata=verbose;', "Content-Type": "application/json;odata=verbose",
      },
      success: (resultData) => {

        this.setState({
          FormDigestValue: resultData.d.GetContextWebInformation.FormDigestValue
        });
      },
      error: (jqXHR, textStatus, errorThrown) => {
        _logExceptionError(this.props.currentContext, this.props.exceptionLogGUID,  _formdigest, "inside getaccessToken pmonewitem form: errlog", "PMOListform", "getaccessToken", jqXHR, _projectID);
      }
    });
  }
  //function to close the form and redirect to the Grid page
  private closeform() {
    //e.preventDefault();
    {if(this.props.customGridRequired){
      let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + "/SitePages/Project-Master.aspx";
    window.open(winUrl, '_self');
  }else{
    let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
    window.open(winUrl, '_self');
  }}
    //let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
    // this.setState({
    //   ProjectID : '',
    //   CRM_Id :'',
    //   ProjectName: '',
    //   ClientName: '',
    //   DeliveryManager:'',
    //   ProjectManager: '',
    //   ProjectType: '',
    //   ProjectMode: '',
    //   PlannedStart: '',
    //   PlannedCompletion: '',
    //   ProjectDescription: '',
    //   ProjectLocation: '',
    //   ProjectBudget: '',
    //   ProjectStatus: '',
    //   ProjectProgress:'',
    //   startDate: '',
    //   endDate: '',
    //   focusedInput: '',
    //   FormDigestValue:''
    // });
   // window.open(winURL, '_self');
  }
  //function to reset the form. Currently disabled
  private resetform(e) {

    this.setState({
      ProjectID: '',
      CRM_Id: '',
      ProjectName: '',
      ClientName: '',
      DeliveryManager: '',
      ProjectManager: '',
      ProjectType: '',
      ProjectMode: '',
      PlannedStart: '',
      PlannedCompletion: '',
      ProjectDescription: '',
      ProjectLocation: '',
      ProjectBudget: 0,
      ProjectStatus: '',
      ProjectProgress: 0,
      startDate: '',
      endDate: '',
      focusedInput: '',
      FormDigestValue: '',
      TotalCost:0
    });
    console.log(this.state.ProjectID);
  }
  //function to get the choice column values
  private retrieveAllChoicesFromListField(siteColUrl: string, columnName: string): void {
    let _formdigest = this.state.FormDigestValue; //variable for errorlog function
    let _projectID = this.state.ProjectID; //variable for errorlog function

    const endPoint: string = `${siteColUrl}/_api/web/lists('` + this.props.listGUID + `')/fields?$filter=EntityPropertyName eq '` + columnName + `'`;

    this.props.currentContext.spHttpClient.get(endPoint, SPHttpClient.configurations.v1)
      .then((response: HttpClientResponse) => {
        if (response.ok) {
          response.json()
            .then((jsonResponse) => {
              console.log(jsonResponse.value[0]);
              let dropdownId = jsonResponse.value[0].Title.replace(/\s/g, '');
              jsonResponse.value[0].Choices.forEach(dropdownValue => {
                $('#' + dropdownId).append('<option value="' + dropdownValue + '">' + dropdownValue + '</option>');
              });
            }, (err: any): void => {
              _logExceptionError(this.props.currentContext, this.props.exceptionLogGUID,  _formdigest, "inside retrieveAllChoicesFromListField pmonewitemform: errlog", "PMOListForm", "retrieveAllChoicesFromListField", err, _projectID);
              console.warn(`Failed to fulfill Promise\r\n\t${err}`);
            });
        } else {
          console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
        }
      });
  }
}
