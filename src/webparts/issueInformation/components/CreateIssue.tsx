import * as React from 'react';
//import styles from './IssueInformation.module.scss';
import { IIssueInformationProps } from './IIssueInformationProps';
import { escape } from '@microsoft/sp-lodash-subset';
//extra imports
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse, HttpClientResponse } from "@microsoft/sp-http";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { _getParameterValues } from '../../PMOListForms/components/getQueryString';
import styles from '../../PMOListForms/components/PmoListForms.module.scss';
import { Form, FormGroup, Button, FormControl } from "react-bootstrap";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { SPCreateIssueForm } from "./ICreateIssueColumnFields";
import * as $ from "jquery";
import { _getListEntityName, listType } from '../../PMOListForms/components/getListEntityName';
import { _logExceptionError } from "../../../ExceptionLogging";
import { inputfieldLength } from '../../PMOListForms/components/PmoListForms';

//declaring state
export interface ICreateIssueState {
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
      FormDigestValue: ''
    }
    this.handleChange = this.handleChange.bind(this);
  }
  //loading function when page gets loaded
  public componentDidMount() {
    //Retrive Project ID
    let itemId = _getParameterValues('ProjectID');
    //let isNumber = parseInt(itemId);

    //if (itemId == null || itemId == "" || isNaN(isNumber)) {
    if (itemId == null || itemId == "") {
      alert("Incorrect URL.Redirecting...");
      window.history.back();
    }
    else {
      this._checkExistingProjectId(this.props.currentContext.pageContext.web.absoluteUrl, itemId);
      this.setState({
        ProjectID: itemId
      });
      $('.webPartContainer').hide();
      $('.form-row').css('justify-content', 'center');
      //Get all choice filed values
      allchoiceColumns.forEach(elem => {
        this.retrieveAllChoicesFromListField(this.props.currentContext.pageContext.web.absoluteUrl, elem);
      });
      _getListEntityName(this.props.currentContext, this.props.listGUID);
      // $('.pickerText_4fe0caaf').css('border','0px');
      // $('.pickerInput_4fe0caaf').addClass('form-control');
      $('.form-row').css('justify-content', 'center');

      this.getAccessToken();
      timerID = setInterval(
        () => this.getAccessToken(), 300000);
    }
  }

  public componentWillUnmount() {
    clearInterval(timerID);

  }
  private handleChange = (e) => {
    let newState = {};
    newState[e.target.name] = e.target.value;
    this.setState(newState);

    this.validateDate(e);

    //on change of issueClosedOn change the status
    if(e.target.name == "IssueClosedOn" && (e.target.value != "" || e.target.value != null)){
      this.setState({
        IssueStatus: "Resolved"
      })
    }
    //on change of issue status need to change the IssueclosedOn
    if(e.target.name == "IssueStatus" && (e.target.value != "Resolved" && e.target.value != null)){
      $('.IssueStatus').remove();
      $('#IssueClosedOn').css('border', '1px solid #ced4da');
      this.setState({
        IssueClosedOn:''
      })
    }else if(e.target.name == "IssueStatus" && (e.target.value == "Resolved")){
      $('#IssueClosedOn').css('border', '1px solid red');
      this._validationMessage("IssueStatus", "IssueStatus", "Please fill the Issue Closed Date if issue is Resolved");
    }

    //functin to check the existing Id
    if (e.target.name == "ProjectID" && (e.target.value != 0 || e.target.value == "")) {
      this._checkExistingProjectId(this.props.currentContext.pageContext.web.absoluteUrl, e.target.value);
    } else if (e.target.value == 0) {
      $('.ProjectID').remove();
      $('#ProjectID').closest('div').append('<span class="ProjectID" style="color:red;font-size:9pt">Project Id cannot be 0</span>');
    }
  }
  private handleSubmit = (e) => {
    this.saveIssue(e);
  }
  public render(): React.ReactElement<IIssueInformationProps> {
    //SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");
    return (
      <div id="newItemDiv" className={styles["_main-div"]} >
        <div id="heading" className={styles.heading}><h3>Register an Issue</h3></div>
        <Form onSubmit={this.handleSubmit}>
          <Form.Row className="mt-4">
            {/*-----------RMS ID------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Project Id</Form.Label>
            </FormGroup>
            <FormGroup className={styles.disabledValue + " col-3"}>
              <Form.Label>{this.state.ProjectID}</Form.Label>
              {/* <Form.Control size="sm" type="number" id="ProjectId" name="ProjectID" placeholder="Project ID" onChange={this.handleChange} value={this.state.ProjectID}/> */}
            </FormGroup>
            <FormGroup className="col-6"></FormGroup>
          </Form.Row>
          {/* --------ROW 2----------------- */}
          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Issue Category</Form.Label>
            </FormGroup>
            <FormGroup className="col-9">
              <Form.Control size="sm" as="select" id="IssueCategory" name="IssueCategory" placeholder="Issue Category" onChange={this.handleChange} value={this.state.IssueCategory}>
                <option value="">Select an Option</option>
              </Form.Control>
            </FormGroup>
            {/* <FormGroup className="col-6"></FormGroup> */}
          </Form.Row>
          {/* ---------ROW 3---------------- */}
          <Form.Row>
            {/*-----------Issue Status------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Issue Description</Form.Label>
            </FormGroup>
            <FormGroup className="col-9">
              <Form.Control size="sm" as="textarea" maxLength={inputfieldLength} rows={4} id="IssueDescription" name="IssueDescription" placeholder="Description about the Issue" onChange={this.handleChange} value={this.state.IssueDescription} />
            </FormGroup>
          </Form.Row>
          {/* ---------ROW 4---------------- */}
          <Form.Row>
            {/*-----------Issue Status------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Issue Status</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" as="select" id="IssueStatus" name="IssueStatus" placeholder="Issue Status" onChange={this.handleChange} value={this.state.IssueStatus}>
                <option value="">Select an Option</option>
              </Form.Control>
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
          {/* ---------ROW 5---------------- */}
          <Form.Row>
            {/*-----------Issue Status------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Assigned Team</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="text" id="Assignedteam" name="Assignedteam" placeholder="Assigned Team" onChange={this.handleChange} value={this.state.Assignedteam} />
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            {/*-----------Issue Priority------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Assgined Person</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" id="Assginedperson" type="text" name="Assginedperson" onChange={this.handleChange} value={this.state.Assginedperson} />
            </FormGroup>
          </Form.Row>
          {/* ---------ROW 6---------------- */}
          <Form.Row>
            {/*-----------Issue Status------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Issue Reported On</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="date" id="IssueReportedOn" name="IssueReportedOn" onChange={this.handleChange} value={this.state.IssueReportedOn} />
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            {/*-----------Issue Priority------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel}>Issue Closed On</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" id="IssueClosedOn" type="date" name="IssueClosedOn" onChange={this.handleChange} value={this.state.IssueClosedOn} />
            </FormGroup>
          </Form.Row>
          {/* ---------ROW 7---------------- */}
          <Form.Row>
            {/*-----------Issue Status------------------- */}
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Required Date</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="date" id="RequiredDate" name="RequiredDate" onChange={this.handleChange} value={this.state.RequiredDate} />
            </FormGroup>
            <FormGroup className="col-6"></FormGroup>
          </Form.Row>
          {/*-----------Issue Priority------------- */}
          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Next Steps Or Resolutions</Form.Label>
            </FormGroup>
            <FormGroup className="col-9">
              <Form.Control size="sm" maxLength={inputfieldLength} id="NextStepsOrResolution" as="textarea" rows={4} name="NextStepsOrResolution" placeholder="Next Steps and Resolutions for the Issue" onChange={this.handleChange} value={this.state.NextStepsOrResolution} />
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
              <Button id="cancel" size="sm" variant="primary" onClick={() => { this.closeForm() }}>
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

  //function to validate the date, end date should not be less than start date
  private validateDate(e) {
    let newState = {};
    //validation for date
    if (e.target.name == "IssueReportedOn" && e.target.value != "") {
      this.setState({
        //disable_plannedCompletion: false
      })
      if (this.state.IssueClosedOn != "") {
        $('.IssueClosedOn').text("");
        var date1 = $('#IssueReportedOn').val();
        var date2 = $('#IssueClosedOn').val()
        if (date1 >= date2) {
          $('#IssueClosedOn').val("")
          newState[e.target.name] = "";
          this.setState(newState);
          //alert("Planned Completion Cannot be less than Planned Start");
          $('#IssueClosedOn').closest('div').append('<span class="IssueClosedOn" style="color:red;font-size:9pt">Must be greater than Issue Reported On</span>')
        } else {
          $('.IssueClosedOn').remove();
        }
      }
    } else if (e.target.name == "IssueReportedOn" && e.target.value == "") {
      this.setState({

        IssueClosedOn: ""
        //disable_plannedCompletion: true
      })
    }
    if (e.target.name == "IssueClosedOn") {
      $('.IssueClosedOn').text("");
      var date1 = $('#IssueReportedOn').val();
      var date2 = $('#IssueClosedOn').val()
      if (date1 >= date2) {
        $('#IssueClosedOn').val("")
        newState[e.target.name] = "";
        this.setState(newState);
        //alert("Planned Completion Cannot be less than Planned Start");
        $('#IssueClosedOn').closest('div').append('<span class="IssueClosedOn" style="color:red;font-size:9pt">Must be greater than Issue Reported On</span>')
      } else {
        $('.IssueClosedOn').remove();
      }
    }//validation for date ending
    //-------------same validation for Required Date Field--------------------
    //validation for date
    if (e.target.name == "IssueReportedOn" && e.target.value != "") {
      this.setState({
        //disable_plannedCompletion: false
      })
      if (this.state.RequiredDate != "") {
        $('.RequiredDate').text("");
        var date1 = $('#IssueReportedOn').val();
        var date2 = $('#RequiredDate').val()
        if (date1 > date2) {
          $('#RequiredDate').val("")
          newState[e.target.name] = "";
          this.setState(newState);
          //alert("Planned Completion Cannot be less than Planned Start");
          $('#RequiredDate').closest('div').append('<span class="RequiredDate" style="color:red;font-size:9pt">Must be greater than Issue Reported On</span>')
        } else {
          $('.RequiredDate').remove();
        }
      }
    } else if (e.target.name == "IssueReportedOn" && e.target.value == "") {
      this.setState({

        RequiredDate: ""
        //disable_plannedCompletion: true
      })
    }
    if (e.target.name == "RequiredDate") {
      $('.RequiredDate').text("");
      var date1 = $('#IssueReportedOn').val();
      var date2 = $('#RequiredDate').val()
      if (date1 > date2) {
        $('#RequiredDate').val("")
        newState[e.target.name] = "";
        this.setState(newState);
        //alert("Planned Completion Cannot be less than Planned Start");
        $('#RequiredDate').closest('div').append('<span class="RequiredDate" style="color:red;font-size:9pt">Must be greater than Issue Reported On</span>')
      } else {
        $('.RequiredDate').remove();
      }
    }//validation for date ending
  }
  //save issue to the list
  private saveIssue(e) {
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
    //issueCategory
    if (requestData.IssueCategory == null || requestData.IssueCategory == "" || requestData.IssueCategory.length < 1) {
      this._validationMessage("IssueCategory", "IssueCategory", "Issue Category cannot be empty");
      $('#IssueCategory').css('border', '1px solid red');
      _validate++;
    } else {
      $('.IssueCategory').remove();
      $('#IssueCategory').css('border', '1px solid #ced4da')
    }
    // Risk Closed On mandatory is status is closed
    // Risk Status mandatory 
    if (requestData.IssueStatus == null || requestData.IssueStatus.length < 1 || requestData.IssueStatus == "") {
      this._validationMessage("IssueStatus", "IssueStatus", "Issue Status cannot be empty");
      $('#IssueStatus').css('border', '1px solid red');
      _validate++;
    }
    else {
      $('.IssueStatus').remove();
      $('#IssueStatus').css('border', '1px solid #ced4da');
      if (requestData.IssueStatus == "Resolved") {
        if (requestData.IssueClosedOn == null || requestData.IssueClosedOn.length < 1 || requestData.IssueClosedOn == "") {
          this._validationMessage("IssueClosedOn", "IssueClosedOn", "Issue Closed On cannot be empty if status is Resolved");
          $('#IssueClosedOn').css('border', '1px solid red');
          _validate++;
        }
        else {
          $('.IssueClosedOn').remove();
          $('#IssueClosedOn').css('border', '1px solid #ced4da');
        }
      } else {
        $('.IssueClosedOn').remove();
        $('#IssueClosedOn').css('border', '1px solid #ced4da');
      }
    }
    //issuePriority
    if (requestData.IssuePriority == null || requestData.IssuePriority == "" || requestData.IssuePriority.length < 1) {
      this._validationMessage("IssuePriority", "IssuePriority", "Issue Priority cannot be empty");
      $('#IssuePriority').css('border', '1px solid red');
      _validate++;
    } else {
      $('.IssuePriority').remove();
      $('#IssuePriority').css('border', '1px solid #ced4da');
    }
    //assignedTeam
    if (requestData.Assignedteam == null || requestData.Assignedteam == "" || requestData.Assignedteam.length < 1) {
      this._validationMessage("Assignedteam", "Assignedteam", "Assigned Team cannot be empty");
      $('#Assignedteam').css('border', '1px solid red');
      _validate++;
    } else {
      $('.Assignedteam').remove();
      $('#Assignedteam').css('border', '1px solid #ced4da');
    }
    //assignedPerson
    if (requestData.Assginedperson == null || requestData.Assginedperson == "" || requestData.Assginedperson.length < 1) {
      this._validationMessage("Assginedperson", "Assginedperson", "Assgined Person cannot be empty");
      $('#Assginedperson').css('border', '1px solid red');
      _validate++;
    } else {
      $('.Assginedperson').remove();
      $('#Assginedperson').css('border', '1px solid #ced4da');
    }
    //IssueReportedOn
    if (requestData.IssueReportedOn == null || requestData.IssueReportedOn == "" || requestData.IssueReportedOn.length < 1) {
      this._validationMessage("IssueReportedOn", "IssueReportedOn", "Issue Reported On cannot be empty");
      $('#IssueReportedOn').css('border', '1px solid red');
      _validate++;
    } else {
      $('.IssueReportedOn').remove();
      $('#IssueReportedOn').css('border', '1px solid #ced4da');
    }
    // //IssueClosedOn
    // if(requestData.IssueClosedOn == null || requestData.IssueClosedOn == "" || requestData.IssueClosedOn.length < 1 ){
    //   this._validationMessage("IssueClosedOn", "IssueClosedOn", "Issue Closed On cannot be empty");
    //   $('#IssueClosedOn').css('border','1px solid red');
    //   _validate++;
    // }else{
    //   $('.IssueClosedOn').remove();
    //   $('#IssueClosedOn').css('border','1px solid #ced4da');
    // }
    //requiredDate
    if (requestData.RequiredDate == null || requestData.RequiredDate == "" || requestData.RequiredDate.length < 1) {
      this._validationMessage("RequiredDate", "RequiredDate", "Required Date cannot be empty");
      $('#RequiredDate').css('border', '1px solid red');
      _validate++;
    } else {
      $('.RequiredDate').remove();
      $('#RequiredDate').css('border', '1px solid #ced4da');
    }
    //issueDescription
    if (requestData.IssueDescription == null || requestData.IssueDescription == "" || requestData.IssueDescription.length < 1) {
      this._validationMessage("IssueDescription", "IssueDescription", "Issue Description cannot be empty");
      $('#IssueDescription').css('border', '1px solid red');
      _validate++;
    } else {
      $('.IssueDescription').remove();
      $('#IssueDescription').css('border', '1px solid #ced4da');
    }
    //nextSteps&Resolution
    if (requestData.NextStepsOrResolution == null || requestData.NextStepsOrResolution == "" || requestData.NextStepsOrResolution.length < 1) {
      this._validationMessage("NextStepsOrResolution", "NextStepsOrResolution", "Next Steps or Resolution cannot be empty");
      $('#NextStepsOrResolution').css('border', '1px solid red');
      _validate++;
    } else {
      $('.NextStepsOrResolution').remove();
      $('#NextStepsOrResolution').css('border', '1px solid #ced4da');
    }

    if (_validate > 0) {
      return false;
    }

    let _formdigest = this.state.FormDigestValue; //variable for errorlog function
    let _projectID = this.state.ProjectID; //variable for errorlog function

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
       let winURL = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
        window.open(winURL, '_self');
        
      },
      error: (xhr, status, error) => {
        //function to log error
        _logExceptionError(this.props.currentContext, this.props.exceptionLogGUID ,_formdigest, "inside saveIssue: errlog", "IssueInformation", "saveIssue", xhr, _projectID );
        if (xhr.responseText.match('2130575169')) {
          alert("The Project Id you entered already exists, please try with a new Project Id")
        }
        //alert(JSON.stringify(xhr.responseText));
        let winURL = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
        window.open(winURL,'_self');
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
      FormDigestValue: ''
    })

  }
  //close the form on cancel button click
  private closeForm() {
    
    let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
    window.open(winUrl, '_self');
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
      FormDigestValue: ''
    })
    
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
              _logExceptionError(this.props.currentContext, this.props.exceptionLogGUID ,_formdigest, "inside reterive choice fields from SP List: errlog", "IssueInformation", "retrieveAllChoicesFromListField", err, _projectID );
              console.warn(`Failed to fulfill Promise\r\n\t${err}`);
            });
        } else {
          console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
        }
      });
  }

  //function to check if ProjectId already exists or not
  private _checkExistingProjectId(siteColUrl, ProjectIDValue) {
    let _formdigest = this.state.FormDigestValue; //variable for errorlog function
    let _projectID = this.state.ProjectID; //variable for errorlog function

    const endPoint: string = `${siteColUrl}/_api/web/lists('` + this.props.ProjectMasterGUID + `')/items?Select=ID&$filter=ProjectID eq '${ProjectIDValue}'`;
    let breakCondition = false;
    $('.ProjectID').remove();
    this.props.currentContext.spHttpClient.get(endPoint, SPHttpClient.configurations.v1)
      .then((response: HttpClientResponse) => {
        if (response.ok) {
          response.json()
            .then((jsonResponse) => {
              if (jsonResponse.value.length > 0) {
                jsonResponse.value.forEach(item => {
                  if (ProjectIDValue == item.ProjectID) {
                    breakCondition = true;
                    return true;
                  } else {
                    alert("Invalid Project ID. Please make sure there is no change in URL. Redirecting...");
                    let winURL = this.props.currentContext.pageContext.web.absoluteUrl + "/SitePages/Project-Master.aspx";
                    window.open(winURL, '_self');
                  }
                  // if(ProjectIDValue != item.ProjectID && breakCondition){
                  //   $('.ProjectID').remove();
                  // }

                });
              } else {
                breakCondition = false;
                alert("Invalid Project ID. Please make sure there is no change in URL. Redirecting...");
                let winURL = this.props.currentContext.pageContext.web.absoluteUrl + "/SitePages/Project-Master.aspx";
                window.open(winURL, '_self');
                return false;
              }
            }, (err: any): void => {
              _logExceptionError(this.props.currentContext, this.props.exceptionLogGUID, _formdigest, "inside saveIssue: errlog", "IssueInformation", "saveIssue", err, _projectID );
              console.warn(`Failed to fulfill Promise\r\n\t${err}`);
              alert("Invalid Project ID. Please make sure there is no change in URL. Redirecting...");
              let winURL = this.props.currentContext.pageContext.web.absoluteUrl + "/SitePages/Project-Master.aspx";
              window.open(winURL, '_self');
            });
        } else {
          //_logExceptionError(this.props.currentContext, _formdigest, "inside saveIssue: errlog", "IssueInformation", "saveIssue", err, _projectID );
          console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
          alert("Invalid Project ID. Please make sure there is no change in URL. Redirecting...");
          let winURL = this.props.currentContext.pageContext.web.absoluteUrl + "/SitePages/Project-Master.aspx";
          window.open(winURL, '_self');
        }
      });
  }
  //validaton message for empty fields
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
        _logExceptionError(this.props.currentContext,this.props.exceptionLogGUID, _formdigest, "inside get access token: errlog", "IssueInformation", "getAccessToken", jqXHR, _projectID );
      }
    });
  }
}
