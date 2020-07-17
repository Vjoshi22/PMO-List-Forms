import * as $ from 'jquery';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse, HttpClientResponse } from "@microsoft/sp-http";
import { _getListEntityName, listType } from "./webparts/PMOListForms/components/getListEntityName";
import { SPExceptionLoggingProps } from "./IExceptionLoggingProps";

var exceptionLoggingListGUID: any = "c3f39dde-1797-4f26-a8ce-f8e0429c26e5"

export function _logExceptionError(_title, _fnName, err){
    let requestData = {
        __metadata:
        {
          type: listType
        },
        Title: '',
        WebPartName: '',
        FunctionName: '',
        ResponseCode: '',
        ResponseText: '',
        DetailedError: '',
        ItemID: '',
        ProjectID: ''
  
    } as SPExceptionLoggingProps
    $.ajax({
        url: this.props.currentContext.pageContext.web.absoluteUrl + "/_api/web/lists('" + exceptionLoggingListGUID + "')/items",
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
          let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
          window.open(winURL, '_self');
        },
        error: (xhr, status, error) => {
          if (xhr.responseText.match('2130575169')) {
            alert("The Project Id you entered already exists, please try with a new Project Id")
          }
          //alert(JSON.stringify(xhr.responseText));
          let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
          //window.open(winURL,'_self');
        }
      });
}