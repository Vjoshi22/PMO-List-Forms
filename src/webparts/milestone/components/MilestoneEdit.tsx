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

require('./Milestone.module.scss');
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");

let timerID;
let newitem: boolean;
let listGUID: any = "e163f102-1cc9-4cc5-97b6-c5296811b444"; // Milestones

export default class MilestoneEdit extends React.Component<IMilestoneProps, IMilestoneState> {
  constructor(props: IMilestoneProps, state: IMilestoneState) {
    super(props);
    this.state = {
      ID: "",
      ProjectID: "",
      Phase: "",
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
    this.saveItem = this.saveItem.bind(this);
  }

  public componentDidMount() {
    $('.webPartContainer').hide();
    $('.form-row').css('justify-content', 'center');
    //Get all choice filed values
    allchoiceColumns.forEach(elem => {
      this.retrieveAllChoicesFromListField(this.props.currentContext.pageContext.web.absoluteUrl, elem);
    });

    getListEntityName(this.props.currentContext, listGUID);
    // this.loadItems();
    setTimeout(() => this.loadItems(), 1000);

    this.setFormDigest();
    timerID = setInterval(
      () => this.setFormDigest(), 300000);
  }
  public componentWillUnmount() {
    clearInterval(timerID);
  }

  //For React form controls
  private handleChange = (e) => {
    let newState = {};
    newState[e.target.name] = e.target.value;
    this.setState(newState);
  }

  private handleSubmit = (e) => {
    this.saveItem(e);
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
            <FormGroup className={styles.disabledValue + " col-3"}>
              {/* <Form.Control size="sm" type="number" id="ProjectID" name="ProjectID" placeholder="Project ID" onChange={this.handleChange} value={this.state.ProjectID} /> */}
              <Form.Label>{this.state.ProjectID}</Form.Label>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Milestone ID</Form.Label>
            </FormGroup>
            <FormGroup className={styles.disabledValue + " col-3"}>
              <Form.Label>{this.state.ID}</Form.Label>
            </FormGroup>
          </Form.Row>
          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Phase</Form.Label>
            </FormGroup>
            <FormGroup className="col-9">
              {/* <Form.Control size="sm" id="Phase" as="select" name="Phase" onChange={this.handleChange} value={this.state.Phase}>
                <option value="">Select an Option</option>
              </Form.Control> */}
              <Form.Label>{this.state.Phase}</Form.Label>
            </FormGroup>
          </Form.Row>
          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Tentative Start</Form.Label>
            </FormGroup>
            <FormGroup className={styles.disabledValue + " col-3"}>
              {/* <Form.Control size="sm" type="date" id="PlannedStart" name="PlannedStart" placeholder="Tentative Start" onChange={this.handleChange} value={this.state.PlannedStart} /> */}
              <Form.Label>{this.state.PlannedStart}</Form.Label>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Tentative End</Form.Label>
            </FormGroup>
            <FormGroup className={styles.disabledValue + " col-3"}>
              {/* <Form.Control size="sm" type="date" id="PlannedEnd" name="PlannedEnd" placeholder="Tentative End" onChange={this.handleChange} value={this.state.PlannedEnd} /> */}
              <Form.Label>{this.state.PlannedEnd}</Form.Label>
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
            <FormGroup className="col-9 mb-3">
              <Form.Control size="sm" as="textarea" rows={3} type="text" id="Remarks" name="Remarks" placeholder="Remarks" onChange={this.handleChange} value={this.state.Remarks} />
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Milestone Created On</Form.Label>
            </FormGroup>
            <FormGroup className={styles.disabledValue + " col-3"}>
              <Form.Label>{this.state.MilestoneCreatedOn}</Form.Label>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Last Updated On</Form.Label>
            </FormGroup>
            <FormGroup className={styles.disabledValue + " col-3"}>
              <Form.Label>{this.state.LastUpdatedOn}</Form.Label>
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

  private loadItems() {

    var itemId = GetParameterValues('id');
    if (itemId == "") {
      alert("Incorrect URL");
      let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
      window.open(winURL, '_self');
    } else {
      const url = `${this.props.currentContext.pageContext.web.absoluteUrl}/_api/web/lists('${listGUID}')/items(${itemId})`;
      return this.props.currentContext.spHttpClient.get(url, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }).then((response: SPHttpClientResponse): Promise<ISPMilestoneFields> => {
          return response.json();
        })
        .then((item: ISPMilestoneFields): void => {
          this.setState({
            ID: item.ID,
            ProjectID: item.ProjectID,
            Phase: item.Phase,
            PlannedStart: item.PlannedStart,
            PlannedEnd: item.PlannedEnd,
            MilestoneStatus: item.MilestoneStatus,
            Remarks: item.Remarks,
            MilestoneCreatedOn: item.Created.slice(0, 10),
            LastUpdatedOn: item.Modified.slice(0, 10),
            ActualStart: item.ActualStart,
            ActualEnd: item.ActualEnd
          })
        })
    }
  }
  private saveItem(e) {
    let _formdigest = this.state.FormDigestValue; //variable for errorlog function
    let _projectID = this.state.ProjectID; //variable for errorlog function

    var itemId = GetParameterValues('id');
    let _validate = 0;
    e.preventDefault();

    let requestData = {
      __metadata:
      {
        type: listType
      },
      ProjectID: this.state.ProjectID,
      Phase: this.state.Phase,
      PlannedStart: this.state.PlannedStart,
      PlannedEnd: this.state.PlannedEnd,
      MilestoneStatus: this.state.MilestoneStatus,
      Remarks: this.state.Remarks,
      ActualStart: this.state.ActualStart == null ? "" : this.state.ActualStart,
      ActualEnd: this.state.ActualEnd == null ? "" : this.state.ActualEnd
    } as ISPMilestoneFields;

    //validation
    // // ProjectID Number only and mandatory
    // if (requestData.ProjectID.length < 1) {
    //   this._validationMessage("ProjectID", "ProjectID", "Project ID cannot be empty");
    //   $('#ProjectID').css('border', '1px solid red');
    //   _validate++;
    // }
    // else {
    //   $('#ProjectID').css('border', '1px solid #ced4da')
    // }

    // Phase mandatory 
    if (requestData.MilestoneStatus == null || requestData.MilestoneStatus.length < 1 || requestData.MilestoneStatus == "") {
      this._validationMessage("Phase", "Phase", "Phase cannot be empty");
      $('#Phase').css('border', '1px solid red');
      _validate++;
    } else {
      $('.Phase').remove();
      $('#Phase').css('border', '1px solid #ced4da')
    }

    // // Tentative Start mandatory
    // if (requestData.PlannedStart == null || requestData.PlannedStart.length < 1 || requestData.PlannedStart == "") {
    //   this._validationMessage("PlannedStart", "PlannedStart", "Tentative Start cannot be empty");
    //   $('#PlannedStart').css('border', '1px solid red');
    //   _validate++;
    // }
    // else {
    //   $('#PlannedStart').css('border', '1px solid #ced4da')
    // }

    // // Tentative End mandatory
    // if (requestData.PlannedEnd == null || requestData.PlannedEnd.length < 1 || requestData.PlannedEnd == "") {
    //   this._validationMessage("PlannedEnd", "PlannedEnd", "Tentative End cannot be empty");
    //   $('#PlannedEnd').css('border', '1px solid red');
    //   _validate++;
    // }
    // else {
    //   $('#PlannedEnd').css('border', '1px solid #ced4da')
    // }

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

    // // Remarks mandatory
    // if (requestData.Remarks.length < 1) {
    //   this._validationMessage("Remarks", "Remarks", "Remarks cannot be empty");
    //   $('#Remarks').css('border', '1px solid red');
    //   _validate++;
    // }
    // else {
    //   $('#Remarks').css('border', '1px solid #ced4da')
    // }

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

    if (_validate > 0) {
      return false;
    }

    $.ajax({
      url: `${this.props.currentContext.pageContext.web.absoluteUrl}/_api/web/lists('${listGUID}')/items(${itemId})`,
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
      success: (data, status, xhr) => {
        console.log("Submitted successfully");
        alert("Submitted successfully");
        let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/Lists/Milestones/AllItems.aspx?FilterField1=ProjectID&FilterValue1=' + this.state.ProjectID + '&FilterType1=Number&viewid=81200a51-c410-419a-bc04-a8bdebf24ae0';
        window.open(winURL, '_self');
      },
      error: (xhr, status, error) => {
        _logExceptionError(this.props.currentContext, _formdigest, "inside saveitem Milestone Edit: errlog", "Milestone", "saveitem", xhr, _projectID );
        alert(JSON.stringify(xhr.responseText));
        console.log(xhr.responseText + " | " + error);
        let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/Lists/Milestones/AllItems.aspx?FilterField1=ProjectID&FilterValue1=' + this.state.ProjectID + '&FilterType1=Number&viewid=81200a51-c410-419a-bc04-a8bdebf24ae0';
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
        _logExceptionError(this.props.currentContext, _formdigest, "inside setFormDigest Milestone Edit: errlog", "Milestone", "setFormDigest", jqXHR, _projectID );
      }
    });
  }

  private closeform() {
    let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/Lists/Milestones/AllItems.aspx?FilterField1=ProjectID&FilterValue1=' + this.state.ProjectID + '&FilterType1=Number&viewid=81200a51-c410-419a-bc04-a8bdebf24ae0';
    this.state = {
      ID: "",
      ProjectID: "",
      Phase: "",
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
    window.open(winURL, '_self');
  }
  private retrieveAllChoicesFromListField(siteColUrl: string, columnName: string): void {
    let _formdigest = this.state.FormDigestValue; //variable for errorlog function
    let _projectID = this.state.ProjectID; //variable for errorlog function

    const endPoint: string = `${siteColUrl}/_api/web/lists('` + listGUID + `')/fields?$filter=EntityPropertyName eq '` + columnName + `'`;

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
              _logExceptionError(this.props.currentContext, _formdigest, "inside retrieveAllChoicesFromListField Milestone Edit: errlog", "Milestone", "retrieveAllChoicesFromListField", err, _projectID );
              console.warn(`Failed to fulfill Promise\r\n\t${err}`);
            });
        } else {
          console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
        }
      });
  }
}
