import * as React from 'react';
//import styles from './IssueInformation.module.scss';
import { IIssueInformationProps } from './IIssueInformationProps';
import { escape } from '@microsoft/sp-lodash-subset';
//extra imports
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration  ,SPHttpClientResponse, HttpClientResponse} from "@microsoft/sp-http";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker"; 
import { _getParameterValues } from '../../PMOListForms/components/getQueryString';
import styles from '../../PMOListForms/components/PmoListForms.module.scss';
import { Form, FormGroup, Button, FormControl } from "react-bootstrap";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { SPCreateIssueForm } from "./ICreateIssueColumnFields";
import * as $ from "jquery";
import { _getListEntityName, listType } from '../../PMOListForms/components/getListEntityName';

//declaring state
export interface ICreateIssueState{
  ProjectID: string;
  IssueCategory: string;
  IssueDescription: string;
  NextStepsOrResolution: string;
  IssueStatus: string;
  IssuePriority: string;
  Assignedteam: string;
  Assginedperson: string;
  IssueReportedOn: string;
  IssueClosedOn: string;
  RequiredDate: string;
  FormDigestValue: string;
}
//declaring variables
var listGUID: any = "A373C7C3-3379-49C9-B3B1-AC87C2166DC0";   //"47272d1e-57d9-447e-9cfd-4cff76241a93"; 
var ProjectMasterListGuid: any = "2c3ffd4e-1b73-4623-898d-8e3a1bb60b91";
var timerID;
var allchoiceColumns: any[] = ["IssueCategory", "IssueStatus", "IssuePriority"]

export default class CreateIssue extends React.Component<IIssueInformationProps, ICreateIssueState> {
  constructor(props: IIssueInformationProps, state: ICreateIssueState) {  
    super(props);  
  
    this.state = {
      ProjectID: '',
      IssueCategory: '',
      IssueDescription: '',
      NextStepsOrResolution: '',
      IssueStatus: '',
      IssuePriority: '',
      Assignedteam: '',
      Assginedperson: '',
      IssueReportedOn: '',
      IssueClosedOn: '',
      RequiredDate: '',
      FormDigestValue:''
    }
    this.handleChange=this.handleChange.bind(this);
  }
  //loading function when page gets loaded
  public componentDidMount() {

    $('.webPartContainer').hide();
    allchoiceColumns.forEach(elem => {
      this.retrieveAllChoicesFromListField(this.props.currentContext.pageContext.web.absoluteUrl, elem);
    });

    _getListEntityName(this.props.currentContext, listGUID);
    // $('.pickerText_4fe0caaf').css('border','0px');
    // $('.pickerInput_4fe0caaf').addClass('form-control');
    $('.form-row').css('justify-content','center');
    
    this.getAccessToken();
    timerID=setInterval(
      () =>this.getAccessToken(),300000); 
  }

  public componentWillUnmount()
  {
  clearInterval(timerID);
  
  } 
  private handleChange = (e) =>{
    let newState = {};
    newState[e.target.name] = e.target.value;
    this.setState(newState);

    //functin to check the existing Id
    if(e.target.name == "ProjectID" && (e.target.value != 0 || e.target.value =="")){
      this._checkExistingProjectId(this.props.currentContext.pageContext.web.absoluteUrl, e.target.value);
    } else if(e.target.value == 0){
      $('.ProjectID').remove();
      $('#ProjectId').closest('div').append('<span class="ProjectID" style="color:red;font-size:9pt">Project Id cannot be 0</span>');
    }
  }
  private handleSubmit = (e) =>{
    this.saveIssue(e);
  }
  public render(): React.ReactElement<IIssueInformationProps> {
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");
    return (
      <div id="newItemDiv" className={styles["_main-div"]} >
        <div id="heading" className={styles.heading}><h3>Register an Issue</h3></div>
        <Form onSubmit={this.handleSubmit}>
          <Form.Row className="mt-3">
            {/*-----------RMS ID------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel +" " + styles.required}>Project Id</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="number" id="ProjectId" name="ProjectID" placeholder="Project ID" onChange={this.handleChange} value={this.state.ProjectID}/>
            </FormGroup>
            <FormGroup className="col-6"></FormGroup>
          </Form.Row>
          {/* --------ROW 2----------------- */}
          <Form.Row className="mt-3">
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel +" " + styles.required}>Issue Category</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" as="select" id="IssueCategory" name="IssueCategory" placeholder="Issue Category" onChange={this.handleChange} value={this.state.IssueCategory}>
              <option value="">Select an Option</option>
              </Form.Control>
            </FormGroup>
            <FormGroup className="col-6"></FormGroup>
          </Form.Row>
          {/* ---------ROW 3---------------- */}
          <Form.Row className="mt-3">
            {/*-----------Issue Status------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel +" " + styles.required}>Issue Status</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" as="select" id="IssueStatus" name="IssueStatus" placeholder="Issue Status" onChange={this.handleChange} value={this.state.IssueStatus}/>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            {/*-----------Issue Priority------------- */}
            <FormGroup className="col-2">
                <Form.Label className={styles.customlabel + " " + styles.required}>IssuePriority</Form.Label>
              </FormGroup>
              <FormGroup className="col-3">
                <Form.Control size="sm" id="IssuePriority" as="select" name="IssuePriority" onChange={this.handleChange} value={this.state.IssuePriority}>
                  <option value="">Select an Option</option>
                </Form.Control>
              </FormGroup>
          </Form.Row>
          {/* ---------ROW 4---------------- */}
          <Form.Row className="mt-3">
            {/*-----------Issue Status------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel +" " + styles.required}>Assigned Team</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="text" id="Assignedteam" name="Assignedteam" placeholder="Assigned Team" onChange={this.handleChange} value={this.state.Assignedteam}/>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            {/*-----------Issue Priority------------- */}
            <FormGroup className="col-2">
                <Form.Label className={styles.customlabel + " " + styles.required}>Assgined Person</Form.Label>
              </FormGroup>
              <FormGroup className="col-3">
                <Form.Control size="sm" id="Assginedperson" type="text" name="Assginedperson" onChange={this.handleChange} value={this.state.Assginedperson}/>
              </FormGroup>
          </Form.Row>
           {/* ---------ROW 4---------------- */}
           <Form.Row className="mt-3">
            {/*-----------Issue Status------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel +" " + styles.required}>Issue Reported On</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="date" id="IssueReportedOn" name="IssueReportedOn" onChange={this.handleChange} value={this.state.IssueReportedOn}/>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            {/*-----------Issue Priority------------- */}
            <FormGroup className="col-2">
                <Form.Label className={styles.customlabel + " " + styles.required}>Issue Closed On</Form.Label>
              </FormGroup>
              <FormGroup className="col-3">
                <Form.Control size="sm" id="IssueClosedOn" type="date" name="IssueClosedOn" onChange={this.handleChange} value={this.state.IssueClosedOn}/>
              </FormGroup>
          </Form.Row>
          {/* ---------ROW 5---------------- */}
          <Form.Row className="mt-3">
            {/*-----------Issue Status------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel +" " + styles.required}>Required Date</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="date" id="RequiredDate" name="RequiredDate" onChange={this.handleChange} value={this.state.RequiredDate}/>
            </FormGroup>
            <FormGroup className="col-6"></FormGroup>
          </Form.Row>
          {/* ---------ROW 6---------------- */}
          <Form.Row className="mt-3">
            {/*-----------Issue Status------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel +" " + styles.required}>Issue Description</Form.Label>
            </FormGroup>
            <FormGroup className="col-9">
              <Form.Control size="sm" as="textarea" rows={4} id="IssueDescription" name="IssueDescription" placeholder="Description about the Issue" onChange={this.handleChange} value={this.state.IssueDescription}/>
            </FormGroup>
            </Form.Row>
            {/*-----------Issue Priority------------- */}
            <Form.Row>
            <FormGroup className="col-2">
                <Form.Label className={styles.customlabel + " " + styles.required}>Next Steps Or Resolutions</Form.Label>
              </FormGroup>
              <FormGroup className="col-9">
                <Form.Control size="sm" id="NextStepsOrResolution" as="textarea" rows={4} name="NextStepsOrResolution" placeholder="Next Steps and Resolutions for the Issue" onChange={this.handleChange} value={this.state.NextStepsOrResolution}/>
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
                <Button id="cancel" size="sm" variant="primary" onClick={this.closeForm}>
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
      </div>
    );
  }
  //save issue to the list
  private saveIssue(e){
    e.preventDefault();
    let _validate = 0;
    
    let requestData = {
      __metadata:  
      {  
          type: listType
      },  
      ProjectID: this.state.ProjectID,
      IssueCategory: this.state.IssueCategory,
      IssueDescription: this.state.IssueDescription,
      NextStepsOrResolution: this.state.NextStepsOrResolution,
      IssueStatus: this.state.IssueStatus,
      IssuePriority: this.state.IssuePriority,
      Assignedteam: this.state.Assignedteam,
      Assginedperson: this.state.Assginedperson,
      IssueReportedOn: this.state.IssueReportedOn,
      IssueClosedOn: this.state.IssueClosedOn,
      RequiredDate: this.state.RequiredDate

    };

    this._checkExistingProjectId(this.props.currentContext.pageContext.web.absoluteUrl, this.state.ProjectID);


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
          let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
          window.open(winURL,'_self');
        },  
        error: (xhr, status, error)=>
        {  
          if(xhr.responseText.match('2130575169')){
            alert("The Project Id you entered already exists, please try with a new Project Id")
          }
          //alert(JSON.stringify(xhr.responseText));
          let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
          //window.open(winURL,'_self');
        }  
    });
    //clearing the fields
    this.setState({
      ProjectID: '',
      IssueCategory: '',
      IssueDescription: '',
      NextStepsOrResolution: '',
      IssueStatus: '',
      IssuePriority: '',
      Assignedteam: '',
      Assginedperson: '',
      IssueReportedOn: '',
      IssueClosedOn: '',
      RequiredDate: '',
      FormDigestValue:''
    })

  }
  //close the form on cancel button click
  private closeForm(){

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

  //function to check if ProjectId already exists or not
  private _checkExistingProjectId(siteColUrl, ProjectIDValue){
    const endPoint: string = `${siteColUrl}/_api/web/lists('`+ ProjectMasterListGuid +`')/items?Select=ID&$filter=ProjectID eq '${ProjectIDValue}'`;
    let breakCondition = false;
    $('.ProjectID').remove();
    this.props.currentContext.spHttpClient.get(endPoint, SPHttpClient.configurations.v1)
      .then((response: HttpClientResponse) => {
        if (response.ok) {
          response.json()
            .then((jsonResponse) => {
              if(jsonResponse.value.length > 0){
              jsonResponse.value.forEach( item => {
              if(ProjectIDValue == item.ProjectID){
                breakCondition = true;
                
              } 
              // if(ProjectIDValue != item.ProjectID && breakCondition){
              //   $('.ProjectID').remove();
              // }
              
              });
            }else{
              breakCondition = false;
              $('#ProjectId').closest('div').append('<span class="ProjectID" style="color:red;font-size:9pt">Project Id does not Exists</span>');
            }
            }, (err: any): void => {
              console.warn(`Failed to fulfill Promise\r\n\t${err}`);
            });
        } else {
          console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
        }
      });
  }
  //function to keep the request digest token active
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
}
