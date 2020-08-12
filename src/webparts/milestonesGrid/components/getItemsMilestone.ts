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
import { ISPMilestoneList } from "../MilestonesGridWebPart";
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

export function _getallItems(url: string, currentContext: any, absoluteURL: any): Promise<ISPMilestoneList[]> {
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
    }) as Promise<ISPMilestoneList[]>
}

export function _populateGrid(results) {
  $('.dataGrid').append(GenerateTablefromJSON(results));

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
    "order": [[0, "desc"]]
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
function GenerateTablefromJSON(data) {
  var tablecontent =
    '<table id="FilesTable" class="table table-hover table-responsive cell-border" cellspacing="0" width="100%">' +
    '<thead><tr id="columnFilters">' +
    '<th class="actionLink">Update Milestone</th>' +
    '<th class="search">Project ID</th>' +
    '<th class="search">Milestone</th>' +
    '<th class="search">Phase</th>' +
    '<th class="search">PlannedStart</th>' +
    '<th class="search">PlannedEnd</th>' +
    '<th class="search">MilestoneStatus</th>' +
    '<th class="search">Remarks</th>' +
    '<th class="search">ActualStart</th>' +
    '<th class="search">ActualEnd</th>' +
    '<th class="search">Created</th>' +
    '<th class="search">Modified</th>' +
    '</tr></thead>' +
    '<tbody>';

  for (var i = 0; i < data.length; i++) {
    let MilestoneDetails = `https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Track-Milestone.aspx?page=edit,id=` + data[i].Id 

    tablecontent += '<tr id="' + data[i].Id + 'row">';
    tablecontent += "<td class='" + data[i].Id + "rowItem'><a id=IssueDetails" + data[i].Id +
      "' target='_blank' style='color: teal' class='confirmEditFileLink' href=" + MilestoneDetails + " data-interception='off'>" +
      "<i class='fa fa-pencil' aria-hidden='true'></i></a>&nbsp&nbsp&nbsp;&nbsp;</a></td>";

    // tablecontent += '<tr id="' + data[i].Id + 'row">';
    // tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Id + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].ProjectID + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Milestone + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Phase + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].PlannedStart + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].PlannedEnd + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].MilestoneStatus + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Remarks + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].ActualStart + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].ActualEnd + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Created + "</td>";
    tablecontent += '<td class="' + data[i].ProjectID + 'rowItem">' + data[i].Modified + "</td>";
    tablecontent += '</tr>';
  }
  tablecontent += '</tbody></table>';
  return tablecontent
}