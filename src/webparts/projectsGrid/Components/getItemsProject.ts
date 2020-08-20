import "datatables.net";
import "datatables.net-dt";
import "datatables.net-responsive";
import "datatables.net-fixedheader";
import "datatables.net-fixedheader-dt";
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
SPComponentLoader.loadCss("https://cdn.datatables.net/fixedheader/3.1.7/js/dataTables.fixedHeader.min.js");
SPComponentLoader.loadCss("https://cdn.datatables.net/1.10.21/css/jquery.dataTables.min.css");
SPComponentLoader.loadCss("https://cdn.datatables.net/fixedheader/3.1.7/css/fixedHeader.dataTables.min.css")


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

export function _populateGrid(results, currentContext) {
  $('.dataGrid').append(GenerateTablefromJSON(results, currentContext));

  table = $('#FilesTable').DataTable({
    "columnDefs": [{
      "targets": '_all',
      "createdCell": function (td, cellData, rowData, row, col) {
        if (cellData == "null") {
          //$(td).css('color', 'red');
          $(td).html(cellData.replace('null', '-'));
        }
      }
    }],
    //"dom": 'lftrip',//(l)ength,(f)iltering,(t)able,(i)nformation,(p)aging, P(r)ocessing
    "dom": "<<t>ip>",
    //"lengthChange": false,   
    "order": [[0, "desc"]],
    "fixedHeader": {
      header:true
    }
  });

  $('#FilesTable th.search').css({ 'min-width': '130px' });
  $('#FilesTable th.actionLink').css({ 'min-width': '130px' });

  $('.dataTables_filter input').addClass('form-control');
  $('.dataTables_length label').addClass('col-form-label');

  $('#FilesTable thead tr').clone(false).appendTo('#FilesTable thead');
  $('#FilesTable thead tr:eq(1) th').removeClass("sorting sorting_desc");

  $('#FilesTable thead tr:eq(1) th.search').each(function (i) {
    var title = $(this).text();
    $(this).html('<input type="text" class="colSearchInputs" id="' + title + '" placeholder="Search ' + title + '" />');
    $('.colSearchInputs').on('keyup change', function () {
      if (table.column(i).search() !== (<any>(this)).value) {
        table
          .column($(this).closest('th').index())
          .search((<any>(this)).value)
          .draw();
      }
    });
  });
  $('#FilesTable thead tr:eq(1) th.actionLink').each(function (index, th) {
    $(this).text("");
  });
  $('#FilesTable thead tr:eq(1) th.dropdown').each(function () {
    var title = $(this).text();
    var ddColumn = table.column($(this).index());
    var select = $('<select><option value="">Select ' + title + '</option></select>')
      .appendTo($(this).empty())
      .on('change', function () {
        ddColumn
          .search($(this).val())
          .draw();
      });

    ddColumn.data().unique().sort().each(function (d, j) {
      select.append('<option value="' + d + '">' + d + '</option>')
    });
  });
}
function GenerateTablefromJSON(data, currentContext) {
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
    '<th class="search">Project Phase</th>' +
    '<th class="dropdown">Project Type</th>' +
    '<th class="dropdown">Region</th>' +
    '<th class="search">Project Budget</th>' +
    '<th class="search">Planned Start</th>' +
    '<th class="search">Planned End</th>' +
    '<th class="dropdown">Project Mode</th>' +
    '<th class="dropdown">Status</th>' +
    '<th class="search">Delivery Manager</th>' +
    '<th class="search">Project Manager</th>' +
    '<th class="search">Progress</th>' +
    '<th class="search">Actual Start</th>' +
    '<th class="search">Actual End</th>' +
    '<th class="search">Revised Budget</th>' +
    '<th class="search">Total Cost</th>' +
    '<th class="search">Invoiced amount</th>' +
    '<th class="dropdown">Scope</th>' +
    '<th class="dropdown">Schedule</th>' +
    '<th class="dropdown">Resource</th>' +
    '<th class="dropdown">Project Cost</th>' +
    '</tr></thead>' +
    '<tbody>';

  for (var i = 0; i < data.length; i++) {
    let projectUpdateLink = currentContext.pageContext.web.absoluteUrl + '/SitePages/UpdateProject.aspx?page=edit&id=' + data[i].ID;
    let createIssuesLink = currentContext.pageContext.web.absoluteUrl + '/SitePages/Track-Issues.aspx?type=new&ProjectID=' + data[i].ProjectID;
    let viewIssuesLink = currentContext.pageContext.web.absoluteUrl  + '/SitePages/Issue-Grid.aspx?FilterField1=ProjectID&FilterValue1=' + data[i].ProjectID;
    let createRisksLink = currentContext.pageContext.web.absoluteUrl  + '/SitePages/Track-Risks.aspx?type=new&ProjectID=' + data[i].ProjectID;
    let viewRisksLink = currentContext.pageContext.web.absoluteUrl  + '/SitePages/Risk-Grid.aspx?FilterField1=ProjectID&FilterValue1=' + data[i].ProjectID;
    let creatMlestoneLink = currentContext.pageContext.web.absoluteUrl  + '/SitePages/Track-Milestone.aspx?type=new&ProjectID=' + data[i].ProjectID;
    let viewMilestoneLink = currentContext.pageContext.web.absoluteUrl  + '/SitePages/Milestone-Grid.aspx?FilterField1=ProjectID&FilterValue1=' + data[i].ProjectID;

    tablecontent += '<tr id="' + data[i].ID + 'row">';
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].ProjectID + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Project_x0020_Name + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Client_x0020_Name + "</td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id=UpdateDetails'" + data[i].Id +
      "' target='_blank' style='color: teal' class='confirmEditFileLink' href=" + projectUpdateLink + " data-interception='off'>" +
      "<i class='fa fa-pencil' aria-hidden='true'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: orange' class='confirmEditFileLink' href=" + createIssuesLink + " data-interception='off'>" +
      "<i class='fa fa-plus' aria-hidden='true'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: orange' class='confirmEditFileLink' href=" + viewIssuesLink + " data-interception='off'>" +
      "<i class='fa fa-list-alt' aria-hidden='true'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: red' class='confirmEditFileLink' href=" + createRisksLink + " data-interception='off'>" +
      "<i class='fa fa-plus' aria-hidden='true'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: red' class='confirmEditFileLink' href=" + viewRisksLink + " data-interception='off'>" +
      "<i class='fa fa-list-alt' aria-hidden='true'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: blue' class='confirmEditFileLink' href=" + creatMlestoneLink + " data-interception='off'>" +
      "<i class='fa fa-plus' aria-hidden='true'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += "<td class='" + data[i].ProjectID + "rowItem'><a id='" + data[i].Id +
      "' target='_blank' style='color: blue' class='confirmEditFileLink' href=" + viewMilestoneLink + " data-interception='off'>" +
      "<i class='fa fa-list-alt' aria-hidden='true'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Project_x0020_Description + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Project_x0020_Phase + "</td>";
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