import * as $ from 'jquery';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientConfiguration, SPHttpClientResponse, HttpClientResponse } from "@microsoft/sp-http";
import { _getListEntityName, listType } from "./webparts/PMOListForms/components/getListEntityName";
import { SPExceptionLoggingProps } from "./IExceptionLoggingProps";
import { _getParameterValues } from './webparts/PMOListForms/components/getQueryString';
import { IPmoListFormsProps } from "./webparts/PMOListForms/components/IPmoListFormsProps";

//var exceptionLoggingListGUID: any = "c3f39dde-1797-4f26-a8ce-f8e0429c26e5";

export function _logExceptionError(_currentContext, exceptionLoggingListGUID, _formdigest, _title, _webpartName, _fnName, err, _projectId){
    let requestData = {
        __metadata:
        {
          type: "SP.Data.Exception_x0020_LoggingListItem"
        },
        Title: _title,
        WebPartName: _webpartName,
        FunctionName: _fnName,
        ResponseCode: err.responseJSON.error.code,
        ResponseText: err.responseText,
        DetailedError: '',
        ItemID: _getParameterValues('itemId'),
        ProjectID: _projectId
  
    } as SPExceptionLoggingProps

    $.ajax({
        url: _currentContext.pageContext.web.absoluteUrl + "/_api/web/lists('" + exceptionLoggingListGUID + "')/items",
        type: "POST",
        data: JSON.stringify(requestData),
        headers:
        {
          "Accept": "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": _formdigest,
          "IF-MATCH": "*",
          'X-HTTP-Method': 'POST'
        },
        success: (data, status, xhr) => {
           //alert("error added successfully");
          // let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
          // window.open(winURL, '_self');
        },
        error: (xhr, status, error) => {
        //    if (xhr.responseText.match('2130575169')) {
        //      alert("The Project Id you entered already exists, please try with a new Project Id")
        //    }
        //  alert(JSON.stringify(xhr.responseText));
          // let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
          // //window.open(winURL,'_self');
        }
      });
}
