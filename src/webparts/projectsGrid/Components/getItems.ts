import "datatables.net";
import "datatables.net-dt";
import "datatables.net-responsive";
import { SPComponentLoader } from "@microsoft/sp-loader";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import * as $ from "jquery";

import { ISPProjectsList } from "../ProjectsGridWebPart";
export var table;

SPComponentLoader.loadCss("https://code.jquery.com/jquery-3.5.1.js");
SPComponentLoader.loadCss(
  "https://cdn.datatables.net/1.10.21/js/jquery.dataTables.min.js"
);
SPComponentLoader.loadCss(
  "https://cdn.datatables.net/1.10.21/js/dataTables.bootstrap4.min.js"
);
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css");
//SPComponentLoader.loadCss("https://cdn.datatables.net/1.10.21/css/dataTables.bootstrap4.min.css");
SPComponentLoader.loadCss("https://cdn.datatables.net/1.10.21/css/jquery.dataTables.min.css");
SPComponentLoader.loadCss("https://cdn.datatables.net/fixedheader/3.1.7/js/dataTables.fixedHeader.min.js");
SPComponentLoader.loadCss("https://cdn.datatables.net/fixedheader/3.1.7/css/fixedHeader.dataTables.min.css");

export function _getallItems(url: string, currentContext: any, absoluteURL: any): Promise<ISPProjectsList[]> {
  $('.dataGrid').empty();
  let requestURL = absoluteURL.concat(url);

  return currentContext.spHttpClient.get(requestURL,
    SPHttpClient.configurations.v1)
    .then(response => {
      return response.json();
    }).then(jsonresponse => {
      return jsonresponse.value;
      console.log(jsonresponse.value);
    }).catch(error => {
      console.log("Error : " + error.message);
    }) as Promise<ISPProjectsList[]>
}

export function _populateGrid(results) {
  $('.dataGrid').append(GenerateTablefromJSON(results));

  table = $('#FilesTable').DataTable({
    "columnDefs": [{
      "targets": '_all',
      "createdCell": function (td, cellData, rowData, row, col) {
        if (cellData == "null") {
          //$(td).css('color', 'red');
          $(td).html(cellData.replace('null','-'));
        }        
      }
    }],
    "searching": false,
    "lengthChange": false,
    "order": [[0, "desc"]]
  });

  $('#FilesTable th.search').css({ 'min-width': '130px' });
  $('#FilesTable th.actionLink').css({ 'min-width': '70px' });


  $('.dataTables_filter input').addClass('form-control');
  $('.dataTables_length label').addClass('col-form-label');

  //$('#FilesTable thead').append('<tr id="columnSearch"></tr>');
  $('#FilesTable thead th.search').each(function (index, th) {
    var title = $(this).text();
    //$('#columnSearch').append('<th><input type="text" class="colSearchInputs" id="' + title + '" placeholder="Search ' + title + '" /></th>');
    $(th).append('<input type="text" class="colSearchInputs" id="' + title + '" placeholder="Search ' + title + '" />');
  });
  $('#FilesTable thead th.actionLink').each(function (index, th) {
    var title = $(this).text();    
    $(th).append('<label/>&nbsp;</label>');
  });

  //search for all columns
  $('.colSearchInputs').on('keyup change', function () {
    table
      .column($(this).closest('th').index())
      .search((<any>this).value)
      .draw();
  });
}
function GenerateTablefromJSON(data) {
  var tablecontent =
    '<table id="FilesTable" class="table table-hover table-responsive cell-border" cellspacing="0" width="100%">' +
    '<thead><tr id="columnFilters">' +
    '<th class="search">Project ID</th>' +
    '<th class="search">Project Name</th>' +
    '<th class="search">Client Name</th>' +
    '<th class="actionLink">Update Details</th>' +
    '<th class="actionLink">Create Issues</th>' +
    '<th class="actionLink">View Issues</th>' +
    '<th class="actionLink">Create Risks</th>' +
    '<th class="actionLink">View Risks</th>' +
    '<th class="actionLink">Create Milestone</th>' +
    '<th class="actionLink">View Milestone</th>' +
    '<th class="search">Project Description</th>' +
    '<th class="search">Project Type</th>' +
    '<th class="search">Region</th>' +
    '<th class="search">Project Budget</th>' +
    '<th class="search">Planned Start</th>' +
    '<th class="search">Planned End</th>' +
    '<th class="search">Project Mode</th>' +
    '<th class="search">Status</th>' +
    '<th class="search">Delivery Manager</th>' +
    '<th class="search">Project Manager</th>' +
    '<th class="search">Progress</th>' +
    '<th class="search">Actual Start</th>' +
    '<th class="search">Actual End</th>' +
    '<th class="search">Revised Budget</th>' +
    '<th class="search">Total Cost</th>' +
    '<th class="search">Invoiced amount</th>' +
    '<th class="search">Scope</th>' +
    '<th class="search">Schedule</th>' +
    '<th class="search">Resource</th>' +
    '<th class="search">Project Cost</th>' +
    '</tr></thead>' +
    '<tbody>';

  for (var i = 0; i < data.length; i++) {
    tablecontent += '<tr id="' + data[i].ID + 'row">';
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].ProjectID + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Project_x0020_Name + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Client_x0020_Name + "</td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id=UpdateDetails'" + data[i].Id +
      "' target='_blank' style='color: teal' class='confirmEditFileLink'>" +
      "<i class='glyphicon glyphicon-pencil' title='Edit File'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: orange' class='confirmEditFileLink'>" +
      "<i class='glyphicon glyphicon-plus' title='Edit File'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: orange' class='confirmEditFileLink'>" +
      "<i class='glyphicon glyphicon-list-alt' title='Edit File'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: red' class='confirmEditFileLink'>" +
      "<i class='glyphicon glyphicon-plus' title='Edit File'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: red' class='confirmEditFileLink'>" +
      "<i class='glyphicon glyphicon-list-alt' title='Edit File'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: blue' class='confirmEditFileLink'>" +
      "<i class='glyphicon glyphicon-plus' title='Edit File'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: blue' class='confirmEditFileLink'>" +
      "<i class='glyphicon glyphicon-list-alt' title='Edit File'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Project_x0020_Description + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Project_x0020_Type + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Region + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Project_x0020_Budget + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].PlannedStart + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Planned_x0020_End + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Project_x0020_Mode + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Status + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Delivery_x0020_Manager + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Project_x0020_Manager + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Progress + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Actual_x0020_Start + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Actual_x0020_End + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Revised_x0020_Budget + "</td>";

    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Total_x0020_Cost + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Invoiced_x0020_amount + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Scope + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Schedule + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Resource + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Project_x0020_Cost + "</td>";
    tablecontent += '</tr>';
  }
  tablecontent += '</tbody></table>';
  return tablecontent
}