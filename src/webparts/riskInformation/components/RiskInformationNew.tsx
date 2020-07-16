import * as React from 'react';
import styles from './RiskInformation.module.scss';
import { IRiskInformationProps } from './IRiskInformationProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse, HttpClientResponse } from "@microsoft/sp-http";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { GetParameterValues } from './getQueryString';
import { Form, FormGroup, Button, FormControl } from "react-bootstrap";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { IRiskInformationWebPartProps } from "../RiskInformationWebPart";
import * as $ from "jquery";
import { getListEntityName, listType } from './getListEntityName';
import { ISPRiskInformationFields } from './IRiskInformationFileds';
import { IRiskInformationState } from './IRiskInformationState';
import { allchoiceColumns } from "../RiskInformationWebPart";

require('./RiskInformation.module.scss');
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");

let timerID;
let newitem: boolean;
let listGUID: any = "b94d8766-9e5a-41ae-afc6-b00a0bbe0149"; // Risk Information;
let ProjectlistGUID: any = "2c3ffd4e-1b73-4623-898d-8e3a1bb60b91" //Project Master List

export default class RiskInformationNew extends React.Component<IRiskInformationProps, IRiskInformationState> {
  constructor(props: IRiskInformationProps, state: IRiskInformationState) {
    super(props);
    this.state = {
      Title: "",
      RiskID: "",
      ProjectID: "",
      RiskName: "",
      RiskDescription: "",
      RiskCategory: "",
      RiskIdentifiedOn: "",
      RiskClosedOn: "",
      RiskStatus: "",
      RiskOwner: "",
      RiskResponse: "",
      RiskImpact: "",
      RiskProbability: "",
      RiskRank: "",
      Remarks: "",
      focusedInput: "",
      FormDigestValue: ""
    };
    this.handleChange = this.handleChange.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
  }

  public componentDidMount() {
    //Retrive Project ID
    let itemId = GetParameterValues('ProjectID');
    let isNumber = parseInt(itemId);
    if (itemId == null || itemId == "" || isNaN(isNumber)) {
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

      getListEntityName(this.props.currentContext, listGUID);
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

    //Calculate Risk Rank
    if (e.target.name == "RiskImpact") {
      if (e.target.value != "" && this.state.RiskProbability != "") {
        try {
          this.setState({
            RiskRank: (Number(e.target.value) * Number(this.state.RiskProbability)).toString()
          })
        }
        catch (ex) {
          console.log("Error in Calculating Risk Rank");
        }
      }
      else {
        this.setState({
          RiskRank: ""
        })
      }
    }
    if (e.target.name == "RiskProbability") {
      if (e.target.value != "" && this.state.RiskImpact != "") {
        try {
          this.setState({
            RiskRank: (Number(e.target.value) * Number(this.state.RiskImpact)).toString()
          })
        }
        catch (ex) {
          console.log("Error in Calculating Risk Rank");
        }
      }
      else {
        this.setState({
          RiskRank: ""
        })
      }
      //console.log("Rank : " + this.state.RiskRank); //this.state.RiskRank Doesn't relect correct value until onchange finish execution
    }
    if (e.target.name == "RiskIdentifiedOn" || e.target.name == "RiskClosedOn") {
      //Should not be future date
      let todaysdate = new Date();
      let date1 = new Date($('#RiskIdentifiedOn').val().toString());
      let date2 = new Date($('#RiskClosedOn').val().toString());
      if (e.target.name == "RiskIdentifiedOn") {
        $('.errRiskClosedOn').remove();
        if (todaysdate < date1) {
          $('#RiskIdentifiedOn').closest('div').append(`<span class="errRiskIdentifiedOn" style="color:red;font-size:9pt">Can't be greater than today's date</span>`)
        } else {
          $('.errRiskIdentifiedOn').remove();
        }
      }
      if (e.target.name == "RiskClosedOn") {
        $('.errRiskClosedOn').remove();
        if (todaysdate < date2) {
          $('#RiskClosedOn').closest('div').append(`<span class="errRiskClosedOn" style="color:red;font-size:9pt">Can't be greater than today's date</span>`)
        } else {
          $('.errRiskClosedOn').remove();
          if (date1 > date2) {
            $('#RiskClosedOn').closest('div').append(`<span class="errRiskClosedOn" style="color:red;font-size:9pt">Must be greater than Risk Identified On</span>`)
          } else {
            $('.errRiskClosedOn').remove();
            this.setState({
              RiskStatus: "Closed"
            })
          }
        }
      }
    }
    if (e.target.name == "RiskStatus") {
      if (e.target.value == "Closed" && this.state.RiskClosedOn == "") {
        $('.RiskStatus').remove();
        $('#RiskClosedOn').css('border', '1px solid red');
        $('#RiskStatus').closest('div').append(`<span class="errRiskClosedOn" style="color:red;font-size:9pt">Please enter closing date</span>`)
      }
    }
  }

  private handleSubmit = (e) => {
    this.createItem(e);
  }
  public render(): React.ReactElement<IRiskInformationProps> {

    SPComponentLoader.loadCss("//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css");
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datetimepicker/4.7.14/js/bootstrap-datetimepicker.min.js");
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.4.1/css/bootstrap-datepicker3.css");

    return (
      <div id="newItemDiv" className={styles["_main-div"]} >
        <div id="heading" className={styles.heading}><h5>Risk Details</h5></div>
        <Form onSubmit={this.handleSubmit}>
          <Form.Row className="mt-3">
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Project ID</Form.Label>
            </FormGroup>
            <FormGroup className="col-9 mb-3">
              {/* <Form.Control size="sm" type="text" disabled={this.state.disable_RMSID} id="_RMSID" name="RMS_Id" placeholder="RMS ID" onChange={this.handleChange} value={this.state.RMS_Id} /> */}
              {/* <Form.Control size="sm" type="number" id="ProjectID" name="ProjectID" placeholder="Project ID" onChange={this.handleChange} value={this.state.ProjectID} /> */}
              <Form.Label>{this.state.ProjectID}</Form.Label>
            </FormGroup>
          </Form.Row>
          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Risk Name</Form.Label>
            </FormGroup>
            <FormGroup className="col-9 mb-3">
              <Form.Control size="sm" type="text" id="RiskName" name="RiskName" placeholder="Risk Name" onChange={this.handleChange} value={this.state.RiskName} />
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Risk Description</Form.Label>
            </FormGroup>
            <FormGroup className="col-9 mb-3">
              <Form.Control size="sm" as="textarea" rows={3} type="text" id="RiskDescription" name="RiskDescription" placeholder="Risk Description" onChange={this.handleChange} value={this.state.RiskDescription} />
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Risk Category</Form.Label>
            </FormGroup>
            <FormGroup className="col-9 mb-3">
              <Form.Control size="sm" id="RiskCategory" as="select" name="RiskCategory" onChange={this.handleChange} value={this.state.RiskCategory}>
                <option >Select an Option</option>
              </Form.Control>
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Risk Identified On</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="date" id="RiskIdentifiedOn" name="RiskIdentifiedOn" placeholder="Risk Identified On" onChange={this.handleChange} value={this.state.RiskIdentifiedOn} />
              {/* <DatePicker selected={this.state.PlannedStart}  onChange={this.handleChange} />; */}
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel}>Risk Closed On</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" type="date" id="RiskClosedOn" name="RiskClosedOn" placeholder="Risk Closed On" onChange={this.handleChange} value={this.state.RiskClosedOn} />
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Risk Response</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" id="RiskResponse" as="select" name="RiskResponse" onChange={this.handleChange} value={this.state.RiskResponse}>
                <option value="">Select an Option</option>
              </Form.Control>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Risk Impact</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" id="RiskImpact" as="select" name="RiskImpact" onChange={this.handleChange} value={this.state.RiskImpact}>
                <option value="">Select an Option</option>
              </Form.Control>
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Risk Status</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" id="RiskStatus" as="select" name="RiskStatus" onChange={this.handleChange} value={this.state.RiskStatus}>
                <option value="">Select an Option</option>
              </Form.Control>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Risk Owner</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              {/* <Form.Control size="sm" type="text" disabled={this.state.disable_RMSID} id="_RMSID" name="RMS_Id" placeholder="RMS ID" onChange={this.handleChange} value={this.state.RMS_Id} /> */}
              <Form.Control size="sm" type="text" id="RiskOwner" name="RiskOwner" placeholder="Risk Owner" onChange={this.handleChange} value={this.state.RiskOwner} />
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Risk Probability</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" id="RiskProbability" as="select" name="RiskProbability" onChange={this.handleChange} value={this.state.RiskProbability}>
                <option value="">Select an Option</option>
              </Form.Control>
            </FormGroup>
            <FormGroup className="col-1"></FormGroup>
            <FormGroup className="col-5">
              {/* <Form.Label className={styles.customlabel}>Risk Rank</Form.Label> */}
            </FormGroup>
            <FormGroup className="col-3">
              {/* <Form.Control size="sm" type="text" id="RiskRank" name="RiskRank" placeholder="Risk Rank" onChange={this.handleChange} value={this.state.RiskRank} /> */}
            </FormGroup>
          </Form.Row>

          <Form.Row>
            <FormGroup className="col-2">
              <Form.Label className={styles.customlabel + " " + styles.required}>Remarks</Form.Label>
            </FormGroup>
            <FormGroup className="col-9 mb-3">
              <Form.Control size="sm" as="textarea" rows={3} type="text" id="Remarks" name="Remarks" placeholder="Remarks" onChange={this.handleChange} value={this.state.Remarks} />
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
            {/* <div>
              <Button id="reset" size="sm" variant="primary" onClick={this.resetform}>
                Reset
              </Button>
            </div> */}
          </Form.Row>
        </Form>
      </div >
    );
  }

  private _checkExistingProjectId(siteColUrl, ProjectIDValue) {
    const endPoint: string = `${siteColUrl}/_api/web/lists('` + ProjectlistGUID + `')/items?Select=ID&$filter=ProjectID eq '${ProjectIDValue}'`;
    let breakCondition = false;
    $('.ProjectID').remove();
    this.props.currentContext.spHttpClient.get(endPoint, SPHttpClient.configurations.v1)
      .then((response: HttpClientResponse) => {
        if (response.ok) {
          response.json()
            .then((jsonResponse) => {
              if(jsonResponse.value.length > 0){
              jsonResponse.value.forEach(item => {
                if (ProjectIDValue == item.ProjectID) {                                   
                  breakCondition = true;
                }
                else{
                  alert("Invalid Project ID. Please make sure there is no change in URL. Redirecting...");
                  //window.history.back();
                }
                // if(ProjectIDValue != item.ProjectID && breakCondition){
                //   $('.ProjectID').remove();
                // }

              });
            }else{
              alert("Invalid Project ID. Please make sure there is no change in URL. Redirecting...");
              window.history.back();
            }
            }, (err: any): void => {
              console.warn(`Failed to fulfill Promise\r\n\t${err}`);
            });
        } else {
          console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
        }
      });
  }
  private createItem(e) {
    let _validate = 0;
    e.preventDefault();

    let requestData = {
      __metadata:
      {
        type: listType
      },
      ProjectID: this.state.ProjectID,
      //RiskID: this.state.RiskID,
      RiskName: this.state.RiskName,
      RiskDescription: this.state.RiskDescription,
      RiskCategory: this.state.RiskCategory,
      RiskIdentifiedOn: this.state.RiskIdentifiedOn,
      RiskClosedOn: this.state.RiskClosedOn == null ? "" : this.state.RiskClosedOn,
      RiskStatus: this.state.RiskStatus,
      RiskOwner: this.state.RiskOwner,
      RiskResponse: this.state.RiskResponse,
      RiskImpact: this.state.RiskImpact,
      RiskProbability: this.state.RiskProbability,
      Remarks: this.state.Remarks,
      RiskRank: this.state.RiskRank
    } as ISPRiskInformationFields;

    //validation
    // ProjectID Number only and mandatory
    if (requestData.ProjectID.length < 1) {
      this._validationMessage("ProjectID", "ProjectID", "Project ID cannot be empty");
      $('#ProjectID').css('border', '1px solid red');
      _validate++;
    }
    else {
      $('#ProjectID').css('border', '1px solid #ced4da')
    }
    // Risk Name mandatory
    if (requestData.RiskName.length < 1) {
      this._validationMessage("RiskName", "RiskName", "Risk Name cannot be empty");
      $('#RiskName').css('border', '1px solid red');
      _validate++;
    }
    else {
      $('#RiskName').css('border', '1px solid #ced4da')
    }

    // Risk Description mandatory
    if (requestData.RiskDescription.length < 1) {
      this._validationMessage("RiskDescription", "RiskDescription", "Risk Description cannot be empty");
      $('#RiskDescription').css('border', '1px solid red');
      _validate++;
    }
    else {
      $('#RiskDescription').css('border', '1px solid #ced4da')
    }
    // Risk category mandatory 
    if (requestData.RiskCategory == null || requestData.RiskCategory.length < 1 || requestData.RiskCategory == "") {
      this._validationMessage("RiskCategory", "RiskCategory", "Risk Category cannot be empty");
      $('#RiskCategory').css('border', '1px solid red');
      _validate++;
    } else {
      $('#RiskCategory').css('border', '1px solid #ced4da')
    }

    // Risk identified On mandatory
    if (requestData.RiskIdentifiedOn == null || requestData.RiskIdentifiedOn.length < 1 || requestData.RiskIdentifiedOn == "") {
      this._validationMessage("RiskIdentifiedOn", "RiskIdentifiedOn", "Risk Identified On cannot be empty");
      $('#RiskIdentifiedOn').css('border', '1px solid red');
      _validate++;
    }
    else {
      $('#RiskIdentifiedOn').css('border', '1px solid #ced4da')
    }
    // Risk Closed On mandatory is status is closed
    // Risk Status mandatory 
    if (requestData.RiskStatus == null || requestData.RiskStatus.length < 1 || requestData.RiskStatus == "") {
      this._validationMessage("RiskStatus", "RiskStatus", "Risk Status cannot be empty");
      $('#RiskStatus').css('border', '1px solid red');
      _validate++;
    }
    else {
      $('#RiskStatus').css('border', '1px solid #ced4da');
      if (requestData.RiskStatus == "Closed") {
        if (requestData.RiskClosedOn == null || requestData.RiskClosedOn.length < 1 || requestData.RiskClosedOn == "") {
          this._validationMessage("RiskClosedOn", "RiskClosedOn", "Risk Closed On cannot be empty if status is closed");
          $('#RiskClosedOn').css('border', '1px solid red');
          _validate++;
        }
        else {
          $('#RiskClosedOn').css('border', '1px solid #ced4da')
        }
      }
    }
    // Risk Onwer mandatory
    if (requestData.RiskOwner.length < 1) {
      this._validationMessage("RiskOwner", "RiskOwner", "Risk Owner cannot be empty");
      $('#RiskOwner').css('border', '1px solid red');
      _validate++;
    }
    else {
      $('#RiskOwner').css('border', '1px solid #ced4da')
    }

    // Risk Response mandatory 
    if (requestData.RiskResponse == null || requestData.RiskResponse.length < 1 || requestData.RiskResponse == "") {
      this._validationMessage("RiskResponse", "RiskResponse", "Risk Response cannot be empty");
      $('#RiskResponse').css('border', '1px solid red');
      _validate++;
    } else {
      $('#RiskResponse').css('border', '1px solid #ced4da')
    }

    // Risk Impact mandatory 
    if (requestData.RiskImpact == null || requestData.RiskImpact.length < 1 || requestData.RiskImpact == "") {
      this._validationMessage("RiskImpact", "RiskImpact", "Risk Impact cannot be empty");
      $('#RiskImpact').css('border', '1px solid red');
      _validate++;
    } else {
      $('#RiskImpact').css('border', '1px solid #ced4da')
    }

    // Risk Probability mandatory 
    if (requestData.RiskProbability == null || requestData.RiskProbability.length < 1 || requestData.RiskProbability == "") {
      this._validationMessage("RiskProbability", "RiskProbability", "Risk Probability cannot be empty");
      $('#RiskProbability').css('border', '1px solid red');
      _validate++;
    } else {
      $('#RiskProbability').css('border', '1px solid #ced4da')
    }

    // Remarks mandatory
    if (requestData.Remarks.length < 1) {
      this._validationMessage("Remarks", "Remarks", "Remarks cannot be empty");
      $('#Remarks').css('border', '1px solid red');
      _validate++;
    }
    else {
      $('#Remarks').css('border', '1px solid #ced4da')
    }

    if (_validate > 0) {
      return false;
    }

    $.ajax({
      url: `${this.props.currentContext.pageContext.web.absoluteUrl}/_api/web/lists('${listGUID}')/items`,
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
        let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
        window.open(winURL, '_self');
      },
      error: (xhr, status, error) => {
        alert("Something went wrong, please try after sometime");
        console.log(xhr.responseText + " | " + error);
        // let winURL = 'https://ytpl.sharepoint.com/sites/yashpmo/SitePages/Projects.aspx';
        // window.open(winURL, '_self');
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
      }
    });
  }

  private closeform() {
    let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
    this.state = {
      Title: "",
      RiskID: "",
      ProjectID: "",
      RiskName: "",
      RiskDescription: "",
      RiskCategory: "",
      RiskIdentifiedOn: "",
      RiskClosedOn: "",
      RiskStatus: "",
      RiskOwner: "",
      RiskResponse: "",
      RiskImpact: "",
      RiskProbability: "",
      RiskRank: "",
      Remarks: "",
      focusedInput: "",
      FormDigestValue: ""
    };
    window.open(winURL, '_self');
  }
  private retrieveAllChoicesFromListField(siteColUrl: string, columnName: string): void {
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
              console.warn(`Failed to fulfill Promise\r\n\t${err}`);
            });
        } else {
          console.warn(`List Field interrogation failed; likely to do with interrogation of the incorrect listdata.svc end-point.`);
        }
      });
  }
}
