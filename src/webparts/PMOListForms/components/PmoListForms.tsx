import * as React from 'react';
import styles from './PmoListForms.module.scss';
import { IPmoListFormsProps } from './IPmoListFormsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration  ,SPHttpClientResponse, HttpClientResponse} from "@microsoft/sp-http";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker"; 
import { GetParameterValues } from './getQueryString';
import { Form, FormGroup, Button, FormControl } from "react-bootstrap";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { SPProjectList } from "../components/IProjectListProps";
import * as $ from "jquery";
import { getListEntityName, listType } from './getListEntityName';
import { allchoiceColumns } from "../PmoListFormsWebPart";
import { data } from 'jquery';


require('./PmoListForms.module.scss');
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");

export interface IreactState{
  ProjectID: string,
  CRM_Id: string,
  ProjectName: string;
  ClientName: string;
  ProjectManager: string;
  ProjectType: string;
  ProjectMode: string;
  PlannedStart: string;
  PlannedCompletion: string;
  ProjectDescription: string;
  ProjectLocation: string;
  ProjectBudget: string;
  ProjectStatus: string;
  ProjectProgress: string;
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

export default class PmoListForms extends React.Component<IPmoListFormsProps, IreactState> {
  constructor(props: IPmoListFormsProps, state: IreactState) {  
    super(props);  
  
    this.state = {  
      //status: 'Ready',  
      //items: []
      ProjectID : '',
      CRM_Id :'',
      ProjectName: '',
      ClientName: '',
      ProjectManager: '',
      ProjectType: '',
      ProjectMode: '',
      PlannedStart: '',
      PlannedCompletion: '',
      ProjectDescription: '',
      ProjectLocation: '',
      ProjectBudget: '',
      ProjectProgress:'',
      ProjectStatus: '',
      DeliveryManager:'',
      startDate: '',
      endDate: '',
      disable_RMSID: false,
      disable_plannedCompletion: true,
      focusedInput: '',
      FormDigestValue:''
    };  
    this._getdropdownValues = this._getdropdownValues.bind(this);
    this.handleChange=this.handleChange.bind(this);
    this._getProjectManager =this._getProjectManager.bind(this);
    //this.loadItems = this.loadItems.bind(this);
    //this.isOutsideRange = this.isOutsideRange.bind(this);
  }
  public componentDidMount() {
    allchoiceColumns.forEach(elem => {
      this.retrieveAllChoicesFromListField(this.props.currentContext.pageContext.web.absoluteUrl, elem);
    })
    getListEntityName(this.props.currentContext, listGUID);
    $('.pickerText_4fe0caaf').css('border','0px');
    $('.pickerInput_4fe0caaf').addClass('form-control');
    $('.form-row').css('justify-content','center');
  
    // if((/edit/.test(window.location.href))){
    //   newitem = false;
    //   this.loadItems();
    // }
    // if((/new/.test(window.location.href))){
    //   newitem = true
    // }
    if(!this.state.PlannedStart){
      this.setState({
        disable_plannedCompletion: false
      })
    }
   this.getAccessToken();
    timerID=setInterval(
      () =>this.getAccessToken(),300000); 
 }
 public componentWillUnmount()
 {
  clearInterval(timerID);
  
 } 
 //public  isOutsideRange = day =>day.isAfter(Moment()) || day.isBefore(Moment().subtract(0, "days"));  
  private handleChange = (e) => {
    let newState = {};
    newState[e.target.name] = e.target.value;
    this.setState(newState);
    this.validateDate(e);
  }
  private handleSubmit = (e) =>{
    
    this.createItem(e);
    // if(newitem){
    //   this.createItem(e);
    // }else{
    //   this.saveItem(e);
    // }
  }
  private _getProjectManager = (items: any[]) => {  
    console.log('Items:', items);  
    this.setState({ ProjectManager: items[0].text });
  }
  private _getDeliveryManager = (items: any[]) => {  
    console.log('Items:', items);  
    this.setState({ DeliveryManager: items[0].text });
  }
  private _getdropdownValues(e){
     // this.retrieveAllChoicesFromListField(e);
  }

  public render(): React.ReactElement<IPmoListFormsProps> {

    SPComponentLoader.loadCss("//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css");
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datetimepicker/4.7.14/js/bootstrap-datetimepicker.min.js");
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.4.1/css/bootstrap-datepicker3.css");
    
    return (
      <div id="newItemDiv" className={styles["_main-div"]} >
        <div id="heading" className={styles.heading}><h3>Project Details</h3></div>
      <Form onSubmit={this.handleSubmit}>
        <Form.Row className="mt-3">
          {/*-----------RMS ID------------------- */}
          <FormGroup className="col-2">
            <Form.Label className={styles.customlabel +" " + styles.required}>RMS ID</Form.Label>
          </FormGroup>
          <FormGroup className="col-3">
            <Form.Control size="sm" type="text" disabled={this.state.disable_RMSID} id="ProjectId" name="ProjectID" placeholder="RMS ID" onChange={this.handleChange} value={this.state.ProjectID}/>
          </FormGroup>
          <FormGroup className="col-1"></FormGroup>
          {/*-----------Project Type------------- */}
          <FormGroup className="col-2">
              <Form.Label className={styles.customlabel}>Project Type</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" id="ProjectType" as="select" name="ProjectType" onClick={() =>this._getdropdownValues} onChange={this.handleChange} value={this.state.ProjectType}>
                <option value="">Select an Option</option>
              </Form.Control>
            </FormGroup>
        </Form.Row>

        <Form.Row>
            {/* -----------Client Name------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel}>Client Name</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="text" id="ClientName" name="ClientName" placeholder="Client Name" onChange={this.handleChange} value={this.state.ClientName}/>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            {/* -----------Project Name---------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel +" " + styles.required}>Project Name</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="text" id="ProjectName" name="ProjectName" placeholder="Ex: John Deer" onChange={this.handleChange} value={this.state.ProjectName} />
            </FormGroup>
          </Form.Row>

          <Form.Row>
            {/* --------Delivery Manager------------ */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel}>Delivery Manager</Form.Label>
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
            <FormGroup className="col-3">
              <Form.Control size="sm" id="ProjectMode" as="select" name="ProjectMode" onChange={this.handleChange} value={this.state.ProjectMode}>
                <option value="">Select an Option</option>
              </Form.Control>
            </FormGroup>
          <FormGroup className="col-1"></FormGroup>
          <FormGroup className="col-2">
            <Form.Label className={styles.customlabel}>Project Status</Form.Label>
          </FormGroup>
          <FormGroup className="col-3">
            <Form.Control size="sm" id="Status" as="select" name="ProjectStatus"  onChange={this.handleChange} value={this.state.ProjectStatus}>
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
            <Form.Label className={styles.customlabel}>Tentative Start Date</Form.Label>
          </FormGroup>
          <FormGroup className="col-3">
            <Form.Control size="sm" type="date" id="PlannedStart" name="PlannedStart" placeholder="Planned Start Date" onChange={this.handleChange} value={this.state.PlannedStart}/>
            {/* <DatePicker selected={this.state.PlannedStart}  onChange={this.handleChange} />; */}
          </FormGroup>
          <FormGroup className="col-1"></FormGroup>
          <FormGroup className="col-2"> 
            <Form.Label className={styles.customlabel}>Tentative End Date</Form.Label>
          </FormGroup>
          <FormGroup className="col-3">
            <Form.Control size="sm" type="date" disabled={this.state.disable_plannedCompletion} id="PlannedCompletion" name="PlannedCompletion" placeholder="Planned Completion Date" onChange={this.handleChange} value={this.state.PlannedCompletion}/>
          </FormGroup>
        </Form.Row>
        {/* Project Description */}
        <Form.Row>
          <FormGroup className="col-2"> 
            <Form.Label className={styles.customlabel}>Project Description</Form.Label>
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
          <FormGroup className="col-3">
            <Form.Control size="sm" type="text" id="Region" name="ProjectLocation" placeholder="Project Location" onChange={this.handleChange} value={this.state.ProjectLocation}/>
          </FormGroup>
          <FormGroup className="col-1"></FormGroup>
          <FormGroup className="col-2"> 
            <Form.Label className={styles.customlabel +" " + styles.required}>Budget as per SOW</Form.Label>
          </FormGroup>
          <FormGroup className="col-3">
            <Form.Control size="sm" type="text" id="BudgetSOW" name="ProjectBudget" placeholder="Project Budget" onChange={this.handleChange} value={this.state.ProjectBudget}/>
          </FormGroup>
        </Form.Row>
        <Form.Row className="mb-4">
          <FormGroup className="col-2"> 
            <Form.Label className={styles.customlabel}>Project Progress</Form.Label>
          </FormGroup>
          <FormGroup className="col-3">
            <Form.Control size="sm" type="number" id="ProjectProgress" name="ProjectProgress" placeholder="Project Progress (%)" onChange={this.handleChange} value={this.state.ProjectProgress}/>
          </FormGroup>
          <FormGroup className="col-6">
          </FormGroup>
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
              <Button id="cancel" size="sm" variant="primary" onClick={this.closeform}>
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
  //function to validate the date, end date should not be less than start date
  private validateDate(e){
    let newState = {};
    //validation for date
    if(e.target.name == "PlannedStart" && e.target.value!=""){
      this.setState({
        disable_plannedCompletion: false
      })
      if(this.state.PlannedCompletion!=""){
        $('.errorMessage').text("");
        var date1 = $('#PlannedStart').val();
        var date2 = $('#PlannedCompletion').val()
        if(date1>=date2){
          $('#PlannedCompletion').val("")
          newState[e.target.name] = "";
          this.setState(newState);
          //alert("Planned Completion Cannot be less than Planned Start");
          $('#PlannedCompletion').closest('div').append('<span class="errorMessage" style="color:red;font-size:9pt">Must be greater than Planned Start date</span>')
        }else{
          $('.errorMessage').remove();
        }
    }
    }else if(e.target.name == "PlannedStart" && e.target.value ==""){
      this.setState({
        PlannedCompletion: "",
        disable_plannedCompletion: true
      })
    }
    if(e.target.name == "PlannedCompletion"){
      $('.errorMessage').text("");
      var date1 = $('#PlannedStart').val();
      var date2 = $('#PlannedCompletion').val()
      if(date1>=date2){
        $('#PlannedCompletion').val("")
        newState[e.target.name] = "";
        this.setState(newState);
        //alert("Planned Completion Cannot be less than Planned Start");
        $('#PlannedCompletion').closest('div').append('<span class="errorMessage" style="color:red;font-size:9pt">Must be greater than Planned Start date</span>')
      }else{
        $('.errorMessage').remove();
      }
    }//validation for date ending
  }
  //fucntion to save the new entry in the list
  private createItem(e){
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
      PlannedStart: this.state.PlannedStart,
      Planned_x0020_End: this.state.PlannedCompletion,
      Project_x0020_Description: this.state.ProjectDescription,
      Region: this.state.ProjectLocation,
      Project_x0020_Budget: this.state.ProjectBudget,
      Status: this.state.ProjectStatus,
      Progress: this.state.ProjectProgress

    };
    
    //validation
    if (requestData.ProjectID.length < 1){
      $('#ProjectId').css('border','2px solid red');
      _validate++;
    }else{
      $('#ProjectId').css('border','1px solid #ced4da')
    }
    if( requestData.Project_x0020_Name.length < 1){
      $('#ProjectName').css('border','2px solid red');
      _validate++;
    }else{
      $('#ProjectName').css('border','1px solid #ced4da')
    }
    if(requestData.PlannedStart.length <1){
      $('#PlannedStart').css('border','2px solid red');
      _validate++;
    }else{
      $('#PlannedStart').css('border','1px solid #ced4da');
    }
    if(requestData.Planned_x0020_End.length < 1){
      $('#PlannedCompletion').css('border','2px solid red');
      _validate++;
    }else{
      $('#PlannedCompletion').css('border','1px solid #ced4da');
    }
    if (requestData.Project_x0020_Budget.length < 1) {
      $('#BudgetSOW').css('border','2px solid red');
      _validate++;
    }else{
      $('#BudgetSOW').css('border','1px solid #ced4da')
    }
    if(_validate>0){
      return false;
    }
  
    $.ajax({
      url:this.props.currentContext.pageContext.web.absoluteUrl+ "/_api/web/lists('" + listGUID + "')/items",  
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
        success:(data, status, xhr) => 
        {  
          alert("Submitted successfully");
          let winURL = 'https://ytpl.sharepoint.com/sites/yashpmo/SitePages/Projects.aspx';
          window.open(winURL,'_self');
        },  
        error: (xhr, status, error)=>
        {  
          alert(JSON.stringify(xhr.responseText));
          let winURL = 'https://ytpl.sharepoint.com/sites/yashpmo/SitePages/Projects.aspx';
          window.open(winURL,'_self');
        }  
    });
    
    this.setState({
      ProjectID : '',
      CRM_Id :'',
      ProjectName: '',
      ClientName: '',
      DeliveryManager:'',
      ProjectManager: '',
      ProjectType: '',
      ProjectMode: '',
      PlannedStart: '',
      PlannedCompletion: '',
      ProjectDescription: '',
      ProjectLocation: '',
      ProjectBudget: '',
      ProjectStatus: '',
      ProjectProgress:'',
      startDate: '',
      endDate: '',
      focusedInput: '',
      FormDigestValue:''
    });

  }
    //   //function to keep the request digest token active
  private getAccessToken(){
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
  private closeform(e){
    e.preventDefault();
  let winURL = 'https://ytpl.sharepoint.com/sites/yashpmo/SitePages/Projects.aspx';
  window.open(winURL,'_self');
  this.setState({
    ProjectID : '',
    CRM_Id :'',
    ProjectName: '',
    ClientName: '',
    DeliveryManager:'',
    ProjectManager: '',
    ProjectType: '',
    ProjectMode: '',
    PlannedStart: '',
    PlannedCompletion: '',
    ProjectDescription: '',
    ProjectLocation: '',
    ProjectBudget: '',
    ProjectStatus: '',
    ProjectProgress:'',
    startDate: '',
    endDate: '',
    focusedInput: '',
    FormDigestValue:''
  });
  }
  //function to reset the form. Currently disabled
  private resetform(e){
  
  this.setState({
    ProjectID : '',
    CRM_Id :'',
    ProjectName: '',
    ClientName: '',
    DeliveryManager:'',
    ProjectManager: '',
    ProjectType: '',
    ProjectMode: '',
    PlannedStart: '',
    PlannedCompletion: '',
    ProjectDescription: '',
    ProjectLocation: '',
    ProjectBudget: '',
    ProjectStatus: '',
    ProjectProgress:'',
    startDate: '',
    endDate: '',
    focusedInput: '',
    FormDigestValue:''
  });
  console.log(this.state.ProjectID);
  }
  //function to get the choice column values
  private retrieveAllChoicesFromListField(siteColUrl: string, columnName: string): void {
    const endPoint: string = `${siteColUrl}/_api/web/lists('`+ listGUID +`')/fields?$filter=EntityPropertyName eq '`+ columnName +`'`;
  
    this.props.currentContext.spHttpClient.get(endPoint, SPHttpClient.configurations.v1)
      .then((response: HttpClientResponse) => {
        if (response.ok) {
          response.json()
            .then((jsonResponse) => {
              console.log(jsonResponse.value[0]);
              let dropdownId = jsonResponse.value[0].StaticName.replace(/\s/g, '');
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
    
    // //function to save the edit item
    // private saveItem(e){
    //   var itemId = GetParameterValues('id');
    //   let _validate=0;
    //   e.preventDefault();


    //   let requestData = {
    //     __metadata:  
    //     {  
    //         type: listType
    //     },  
    //     ProjectID : this.state.ProjectID,
    //     Project_x0020_Name: this.state.ProjectName,
    //     Client_x0020_Name: this.state.ClientName,
    //     Delivery_x0020_Manager: this.state.DeliveryManager,
    //     Project_x0020_Manager: this.state.ProjectManager,
    //     Project_x0020_Type: this.state.ProjectType,
    //     Project_x0020_Mode: this.state.ProjectMode,
    //     PlannedStart: this.state.PlannedStart,
    //     Planned_x0020_End: this.state.PlannedCompletion,
    //     Project_x0020_Description: this.state.ProjectDescription,
    //     Region: this.state.ProjectLocation,
    //     Project_x0020_Budget: this.state.ProjectBudget,
    //     Status: this.state.ProjectStatus
    
    //   };
    //   //validation
    //   if (requestData.ProjectID.length < 1){
    //     $('input[name="RMS_Id"]').css('border','2px solid red');
    //     _validate++;
    //   }else{
    //     $('input[name="RMS_Id"]').css('border','1px solid #ced4da')
    //   }
    //   if( requestData.Project_x0020_Name.length < 1){
    //     $('#_projectName').css('border','2px solid red');
    //     _validate++;
    //   }else{
    //     $('#_projectName').css('border','1px solid #ced4da')
    //   }
    //   if (requestData.Project_x0020_Budget.length < 1) {
    //     $('#_budget').css('border','2px solid red');
    //     _validate++;
    //   }else{
    //     $('#_budget').css('border','1px solid #ced4da')
    //   }
    //   if(requestData.PlannedStart.length <1){
    //     $('#inpt_plannedStart').css('border','2px solid red');
    //     _validate++;
    //   }else{
    //     $('#inpt_plannedStart').css('border','1px solid #ced4da');
    //   }
    //   if(requestData.Planned_x0020_End.length < 1){
    //     $('#inpt_plannedCompletion').css('border','2px solid red');
    //     _validate++;
    //   }else{
    //     $('#inpt_plannedCompletion').css('border','1px solid #ced4da');
    //   }
    //   if(_validate>0){
    //     return false;
    //   }
     
    
    //   $.ajax({
    //       url:  this.props.currentContext.pageContext.web.absoluteUrl+ "/_api/web/lists('" + listGUID + "')/items("+ itemId +")",  
    //       type: "POST",  
    //       data: JSON.stringify(requestData),  
    //       headers:  
    //       {  
    //           "Accept": "application/json;odata=verbose",  
    //           "Content-Type": "application/json;odata=verbose",  
    //           "X-RequestDigest": this.state.FormDigestValue,
    //           "IF-MATCH": "*",
    //           'X-HTTP-Method': 'MERGE' 
    //       },  
    //       success:(data, status, xhr) => 
    //       {  
    //         alert("Submitted successfully");
    //         let winURL = 'https://ytpl.sharepoint.com/sites/yashpmo/SitePages/Projects.aspx';
    //         window.open(winURL,'_self');
    //       },  
    //       error: (xhr, status, error)=>
    //       {  
    //         alert(JSON.stringify(xhr.responseText));
    //         let winURL = 'https://ytpl.sharepoint.com/sites/yashpmo/SitePages/Projects.aspx';
    //         window.open(winURL,'_self');
    //       }  
    //   });
      
    //   this.setState({
    //     ProjectID: '',
    //     CRM_Id :'',
    //     ProjectName: '',
    //     ClientName: '',
    //     DeliveryManager:'',
    //     ProjectManager: '',
    //     ProjectType: '',
    //     ProjectMode: '',
    //     PlannedStart: '',
    //     PlannedCompletion: '',
    //     ProjectDescription: '',
    //     ProjectLocation: '',
    //     ProjectBudget: '',
    //     ProjectStatus: '',
    //     startDate: '',
    //     disable_plannedCompletion:true,
    //     endDate: '',
    //     focusedInput: '',
    //     FormDigestValue:''
    //   });
    // }
   
    // //fucntion to load items for particular item id on edit form
    // private loadItems(){
    
    // var itemId = GetParameterValues('id');
    // if(itemId==""){
    //   alert("Incorrect URL");
    //   let winURL = 'https://ytpl.sharepoint.com/sites/yashpmo/SitePages/Projects.aspx';
    //   window.open(winURL,'_self');
    // }else{
    // const url = this.props.currentContext.pageContext.web.absoluteUrl + `/_api/web/lists('`+ listGUID +`')/items(`+ itemId +`)`;
    // return this.props.currentContext.spHttpClient.get(url,SPHttpClient.configurations.v1,  
    //     {  
    //         headers: {  
    //           'Accept': 'application/json;odata=nometadata',  
    //           'odata-version': ''  
    //         }  
    //     }).then((response: SPHttpClientResponse): Promise<SPProjectList> => {  
    //         return response.json();  
    //       })  
    //     .then((item: SPProjectList): void => {   
    //       this.setState({
    //         ProjectID: item.ProjectID,
    //         DeliveryManager: item.Delivery_x0020_Manager,
    //         ProjectName: item.Project_x0020_Name,
    //         ClientName: item.Client_x0020_Name,
    //         ProjectManager: item.Project_x0020_Manager,
    //         ProjectType: item.Project_x0020_Type,
    //         ProjectMode: item.Project_x0020_Mode,
    //         PlannedStart: item.PlannedStart.slice(0,10),
    //         PlannedCompletion: item.Planned_x0020_End.slice(0,10),
    //         ProjectDescription: item.Project_x0020_Description,
    //         ProjectLocation: item.Region,
    //         ProjectBudget: item.Project_x0020_Budget,
    //         ProjectStatus: item.Status,
    //         disable_RMSID: true
    //       })  
    //       console.log(this.state.PlannedStart + " " + this.state.PlannedCompletion) ;
    //     });
    //   }
    // }
}
