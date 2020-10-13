import * as React from 'react';
import styles from './Milestone.module.scss';
import { IMilestoneProps } from './IMilestoneProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse, HttpClientResponse } from "@microsoft/sp-http";
import { GetParameterValues } from './getQueryString';
import { Form, FormGroup, Button, FormControl } from "react-bootstrap";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { IMilestoneWebPartProps, allchoiceColumns } from "../MilestoneWebPart";
import * as $ from "jquery";
import { getListEntityName, listType } from './getListEntityName';
import { ISPMilestoneFields } from './IMilestoneFields';
import { IMilestoneState } from './IMilestoneState';
import { _logExceptionError } from '../../../ExceptionLogging';
import { inputfieldLength, multiLineFieldLength } from '../../PMOListForms/components/PmoListForms';

require('./Milestone.module.scss');
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");

let timerID;
let newitem: boolean;
let listGUID: any = "e163f102-1cc9-4cc5-97b6-c5296811b444"; // Milestones
let projectListGUID: any = "2c3ffd4e-1b73-4623-898d-8e3a1bb60b91"; // Milestones


export default class MilestoneNew extends React.Component<IMilestoneProps, IMilestoneState> {
  constructor(props: IMilestoneProps, state: IMilestoneState) {
    super(props);
    this.state = {
      ID: "",
      ProjectID: "",
      //Phase: "",
      Milestone:"", //adding milestone instead of Phase
      PlannedStart: "",
      PlannedEnd: "",
      MilestoneStatus: "",
      Remarks: "",
      MilestoneCreatedOn: "",
      LastUpdatedOn: "",
      ActualStart: "",
      ActualEnd: "",
      focusedInput: "",
      FormDigestValue: ""
    };
    this.handleChange = this.handleChange.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
  }

  public componentDidMount() {
    //Retrive Project ID
    let itemId = GetParameterValues('ProjectID');
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
      })
      $('.webPartContainer').hide();
      $('.form-row').css('justify-content', 'center');
      //Get all choice filed values
      allchoiceColumns.forEach(elem => {
        this.retrieveAllChoicesFromListField(this.props.currentContext.pageContext.web.absoluteUrl, elem);
      });

      getListEntityName(this.props.currentContext, this.props.listGUID);
      this.setFormDigest();
      timerID = setInterval(
        () => this.setFormDigest(), 300000);
    }
  }
  public componentWillUnmount() {
    clearInterval(timerID);
  }

  //For React form controls
  private handleChange = (e) => {
    let newState = {};
    newState[e.target.name] = e.target.value;
    this.setState(newState);

    //checking planned start and planned end less than today's date or not
    if (e.target.name == "PlannedStart" || e.target.name == "PlannedEnd") {
      //Should not be future date
      let todaysdate = new Date();
      let date1 = new Date($('#PlannedStart').val().toString());
      let date2 = new Date($('#PlannedEnd').val().toString());
      if (e.target.name == "PlannedStart") {
        $('.PlannedStart').remove();
        if (todaysdate < date1) {
          this.setState({
            ActualStart:''
          })//$('#PlannedStart').closest('div').append(`<span class="PlannedStart" style="color:red;font-size:9pt">Can't be greater than today's date</span>`)
        } else {
          this.setState({
            ActualStart: $('#PlannedStart').val().toString()
          })
          //$('.errRiskIdentifiedOn').remove();
        }
      }
      if (e.target.name == "PlannedEnd") {
        $('.PlannedEnd').remove();
        if (todaysdate < date2) {
          this.setState({
            ActualEnd: ''
          })//$('#RiskClosedOn').closest('div').append(`<span class="errRiskClosedOn" style="color:red;font-size:9pt">Can't be greater than today's date</span>`)
        } else {
          this.setState({
            ActualEnd: $('#PlannedEnd').val().toString()
          })
          $('.PlannedEnd').remove();
          if (date1 > date2) {
            this.setState({
              PlannedEnd:'',
              ActualEnd:''
            })
            $('#PlannedEnd').closest('div').append(`<span class="PlannedEnd" style="color:red;font-size:9pt">Must be greater than Planned Start</span>`)
          } else {
            $('.PlannedEnd').remove();
          }
        }
      }
    }
    
  }

  private handleSubmit = (e) => {
    this.createItem(e);
  }
  public render(): React.ReactElement<IMilestoneProps> {

    SPComponentLoader.loadCss("//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css");
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datetimepicker/4.7.14/js/bootstrap-datetimepicker.min.js");
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.4.1/css/bootstrap-datepicker3.css");

    return (
      <div id="newItemDiv" className={styles["_main-div"]} >
        <div id="heading" className={styles.heading}><h5>Milestone</h5></div>
        <Form onSubmit={this.handleSubmit}>
          <Form.Row className="mt-3">
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Project ID</Form.Label>
            </FormGroup>
            <FormGroup className={styles.disabledValue + " col-9"}>
              {/* <Form.Control size="sm" type="number" id="ProjectID" name="ProjectID" placeholder="Project ID" onChange={this.handleChange} value={this.state.ProjectID} /> */}
              <Form.Label>{this.state.ProjectID}</Form.Label>
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              {/* ##CR:Renamed Phase to Milestone and changign the dropdown to text */}
              <Form.Label className={styles.customlabel + " " + styles.required}>Milestone</Form.Label>
            </FormGroup>
            <FormGroup className="col-9">
              {/* <Form.Control size="sm" id="Phase" as="select" name="Phase" onChange={this.handleChange} value={this.state.Phase}>
                <option value="">Select an Option</option>
              </Form.Control> */}
              <Form.Control size="sm" maxLength={inputfieldLength} type="text" id="Milestone" name="Milestone" placeholder="Milestone" onChange={this.handleChange} value={this.state.Milestone} />
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Planned Start</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="date" id="PlannedStart" name="PlannedStart" placeholder="Planned Start" onChange={this.handleChange} value={this.state.PlannedStart} />
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Planned End</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="date" id="PlannedEnd" name="PlannedEnd" placeholder="Planned End" onChange={this.handleChange} value={this.state.PlannedEnd} />
            </FormGroup>
          </Form.Row>
          
          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Actual Start</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="date" id="ActualStart" name="ActualStart" placeholder="Actual Start" onChange={this.handleChange} value={this.state.ActualStart} />
              {/* <DatePicker selected={this.state.PlannedStart}  onChange={this.handleChange} />; */}
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel}>Actual End</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="date" id="ActualEnd" name="ActualEnd" placeholder="Actual End" onChange={this.handleChange} value={this.state.ActualEnd} />
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Milestone Status</Form.Label>
            </FormGroup>
            <FormGroup className="col-9">
              <Form.Control size="sm" id="MilestoneStatus" as="select" name="MilestoneStatus" onChange={this.handleChange} value={this.state.MilestoneStatus}>
                <option value="">Select an Option</option>
              </Form.Control>
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel}>Remarks</Form.Label>
            </FormGroup>
            <FormGroup className="col-9">
              <Form.Control size="sm" as="textarea" maxLength={multiLineFieldLength} rows={3} type="text" id="Remarks" name="Remarks" placeholder="Remarks" onChange={this.handleChange} value={this.state.Remarks} />
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
              <Button id="cancel" size="sm" variant="primary" onClick={() => { this.closeform() }}>
                Cancel
              </Button>
            </div>
          </Form.Row>
        </Form>
      </div >
    );
  }

  private _checkExistingProjectId(siteColUrl, ProjectIDValue) {
    let _formdigest = this.state.FormDigestValue; //variable for errorlog function
    let _projectID = this.state.ProjectID; //variable for errorlog function

    if(this.props.ProjectMasterGUID){
    const endPoint: string = `${siteColUrl}/_api/web/lists('` + this.props.ProjectMasterGUID + `')/items?Select=ID&$filter=ProjectID eq '${ProjectIDValue}'`;
    let breakCondition = false;
    $('.ProjectID').remove();
    this.props.currentContext.spHttpClient.get(endPoint, SPHttpClient.configurations.v1)
      .then((response: HttpClientResponse) => {
        if (response.status == 200) {
          response.json()
            .then((jsonResponse) => {
              if (jsonResponse.value.length > 0) {
                jsonResponse.value.forEach(item => {
                  if (ProjectIDValue == item.ProjectID) {
                    breakCondition = true;
                  }
                  else {
                    alert("Invalid Project ID. Please make sure there is no change in URL. Redirecting...");
                    let winURL = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
                    window.open(winURL, '_self');
                  }
                  // if(ProjectIDValue != item.ProjectID && breakCondition){
                  //   $('.ProjectID').remove();
                  // }

                });
              }
              else {
                alert("Invalid Project ID. Please make sure there is no change in URL. Redirecting...");
                let winURL = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
                window.open(winURL, '_self');
              }
            }, (err: any): void => {
              _logExceptionError(this.props.currentContext, this.props.exceptionLogGUID, _formdigest, "inside _checkExistingProjectId MilestoneNew: errlog", "Milestone", "_checkExistingProjectId", err, _projectID );
              console.warn(`Failed to fulfill Promise\r\n\t${err}`);
              alert("Something went wrong. Please try after sometime Redirecting...");
              let winURL = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
              window.open(winURL, '_self');
            });
        } else {
          console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
            alert("Something went wrong. Please try after sometime Redirecting...");
            let winURL = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
            window.open(winURL, '_self');
        }
      });
    }
  }
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
      // Phase: this.state.Phase,
      Milestone: this.state.Milestone,
      PlannedStart: this.state.PlannedStart,
      PlannedEnd: this.state.PlannedEnd,
      MilestoneStatus: this.state.MilestoneStatus,
      Remarks: this.state.Remarks,
      ActualStart: this.state.ActualStart == null ? "" : this.state.ActualStart,
      ActualEnd: this.state.ActualEnd == null ? "" : this.state.ActualEnd
    } as ISPMilestoneFields;

    //validation
    // ProjectID Number only and mandatory
    // if (requestData.ProjectID.length < 1) {
    //   this._validationMessage("ProjectID", "ProjectID", "Project ID cannot be empty");
    //   $('#ProjectID').css('border', '1px solid red');
    //   _validate++;
    // }
    // else {
    //   $('#ProjectID').css('border', '1px solid #ced4da')
    // }

    //----replaced with Milestone----
    // Phase mandatory 
    // if (requestData.Phase == null || requestData.Phase.length < 1 || requestData.Phase == "") {
    //   this._validationMessage("Phase", "Phase", "Phase cannot be empty");
    //   $('#Phase').css('border', '1px solid red');
    //   _validate++;
    // } else {
    //   $('.Phase').remove();
    //   $('#Phase').css('border', '1px solid #ced4da')
    // }

    // Milestone mandatory 
    if (requestData.Milestone == null || requestData.Milestone.length < 1 || requestData.Milestone == "") {
      this._validationMessage("Milestone", "Milestone", "Milestone cannot be empty");
      $('#Milestone').css('border', '1px solid red');
      _validate++;
    } else {
      $('.Milestone').remove();
      $('#Milestone').css('border', '1px solid #ced4da')
    }

    // Planned Start mandatory
    if (requestData.PlannedStart == null || requestData.PlannedStart.length < 1 || requestData.PlannedStart == "") {
      this._validationMessage("PlannedStart", "PlannedStart", "Planned Start cannot be empty");
      $('#PlannedStart').css('border', '1px solid red');
      _validate++;
    }
    else {
      $('.PlannedStart').remove();
      $('#PlannedStart').css('border', '1px solid #ced4da')
    }

    // Planned End mandatory
    if (requestData.PlannedEnd == null || requestData.PlannedEnd.length < 1 || requestData.PlannedEnd == "") {
      this._validationMessage("PlannedEnd", "PlannedEnd", "Planned End cannot be empty");
      $('#PlannedEnd').css('border', '1px solid red');
      _validate++;
    }
    else {
      $('.PlannedEnd').remove();
      $('#PlannedEnd').css('border', '1px solid #ced4da')
    }

    // // Milestone Status mandatory 
    // if (requestData.MilestoneStatus == null || requestData.MilestoneStatus.length < 1 || requestData.MilestoneStatus == "") {
    //   this._validationMessage("MilestoneStatus", "MilestoneStatus", "Milestone Status cannot be empty");
    //   $('#MilestoneStatus').css('border', '1px solid red');
    //   _validate++;
    // } else {
    //   $('.MilestoneStatus').remove();
    //   $('#MilestoneStatus').css('border', '1px solid #ced4da')
    // } //commented  VJ

    // Milestone Status mandatory & Actual End
   if (requestData.MilestoneStatus == null || requestData.MilestoneStatus.length < 1 || requestData.MilestoneStatus == "") {
    this._validationMessage("MilestoneStatus", "MilestoneStatus", "Milestone Status cannot be empty");
    $('#MilestoneStatus').css('border', '1px solid red');
    _validate++;
  } else {
    $('.MilestoneStatus').remove();
    $('#MilestoneStatus').css('border', '1px solid #ced4da')
    if (requestData.MilestoneStatus == "Completed") {
      if (requestData.ActualEnd == null || requestData.ActualEnd.length < 1 || requestData.ActualEnd == "") {
        this._validationMessage("ActualEnd", "ActualEnd", "Actual End cannot be empty if status is Completed");
        $('#ActualEnd').css('border', '1px solid red');
        _validate++;
      }
      else {
        $('.ActualEnd').remove();
        $('#ActualEnd').css('border', '1px solid #ced4da')
      }
    }
  }

   // Actual Start mandatory
   if (requestData.ActualStart == null || requestData.ActualStart.length < 1 || requestData.ActualStart == "") {
    this._validationMessage("ActualStart", "ActualStart", "Actual Start cannot be empty");
    $('#ActualStart').css('border', '1px solid red');
    _validate++;
  }
  else {
    $('.ActualStart').remove();
    $('#ActualStart').css('border', '1px solid #ced4da')
  }
    // // Remarks mandatory
    // if (requestData.Remarks.length < 1) {
    //   this._validationMessage("Remarks", "Remarks", "Remarks cannot be empty");
    //   $('#Remarks').css('border', '1px solid red');
    //   _validate++;
    // }
    // else {
    //   $('#Remarks').css('border', '1px solid #ced4da')
    // }

    if (_validate > 0) {
      return false;
    }

    $.ajax({
      url: `${this.props.currentContext.pageContext.web.absoluteUrl}/_api/web/lists('${this.props.listGUID}')/items`,
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
        console.log("Submitted successfully");
        alert("Submitted successfully");
       
          let winURL = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
          window.open(winURL, '_self');
        
        
      },
      error: (xhr, status, error) => {
        _logExceptionError(this.props.currentContext, this.props.exceptionLogGUID, _formdigest, "inside createItem Milestone New: errlog", "Milestone", "createItem", xhr, _projectID );
        if (xhr.responseText.match('2147024891')) {
          alert("You don't have permission to create a new Milestone");
        }else{
          alert("Something went wrong, please try after sometime");
        }
        console.log(xhr.responseText + " | " + error);
        let winURL = this.props.currentContext.pageContext.web.absoluteUrl +  '/SitePages/Project-Master.aspx';
        window.open(winURL, '_self');
      }
    });

    // this.state = {
    //   Title: "",
    //   RiskId: -1,
    //   ProjectID: "",
    //   RiskName: "",
    //   RiskDescription: "",
    //   RiskCategory: "",
    //   RiskIdentifiedOn: "",
    //   RiskClosedOn: null,
    //   RiskStatus: "",
    //   RiskOwner: "",
    //   RiskResponse: "",
    //   RiskImpact: "",
    //   RiskProbability: "",
    //   RiskRank: "",
    //   Remarks: "",
    //   focusedInput: "",
    //   FormDigestValue: ""
    // };
  }
  private _validationMessage(_id, _classname, _message) {
    $('.' + _classname).remove();
    $('#' + _id).closest('div').append('<span class="' + _classname + '" style="color:red;font-size:9pt">' + _message + '</span>');
  }
  private setFormDigest() {
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
        _logExceptionError(this.props.currentContext,this.props.exceptionLogGUID,  _formdigest, "inside setFormDigest Milestone New: errlog", "Milestone", "setFormDigest", jqXHR, _projectID );
      }
    });
  }

  private closeform() {
      let winURL = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
      window.open(winURL, '_self');
    
    this.state = {
      ID: "",
      ProjectID: "",
      //Phase: "",
      Milestone: "",
      PlannedStart: "",
      PlannedEnd: "",
      MilestoneStatus: "",
      Remarks: "",
      MilestoneCreatedOn: "",
      LastUpdatedOn: "",
      ActualStart: "",
      ActualEnd: "",
      focusedInput: "",
      FormDigestValue: ""
    };
  }
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
              _logExceptionError(this.props.currentContext, this.props.exceptionLogGUID, _formdigest, "inside retrieveAllChoicesFromListField Milestone New: errlog", "Milestone", "retrieveAllChoicesFromListField", err, _projectID );
              console.warn(`Failed to fulfill Promise\r\n\t${err}`);
            });
        } else {
          console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
        }
      });
  }
}
