import * as myJQuery from 'jquery'
import { SPComponentLoader } from "@microsoft/sp-loader";
import { DateTimeFieldFormatType } from 'sp-pnp-js';
import { table } from "./getItems";


SPComponentLoader.loadCss("//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css");
//SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.15.1/moment.min.js");
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/js/bootstrap.min.js");
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datetimepicker/4.7.14/js/bootstrap-datetimepicker.min.js");
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.4.1/css/bootstrap-datepicker3.css")
import 'jqueryui';


export function _customStyle() {
    //myJQuery(document).ready(function () {
    myJQuery('#FilesTable th').css({ 'min-width': '130px' });
    myJQuery('.dataTables_filter input').addClass('form-control');
    myJQuery('.dataTables_length label').addClass('col-form-label');
    
    myJQuery('#FilesTable thead th').each(function () {
        var title = myJQuery(this).text();
        myJQuery('#columnSearch').append('<th><input type="text" class="colSearchInputs" id="' + title + '" placeholder="Search ' + title + '" /></th>');
    });

    //search for all columns
    myJQuery('.colSearchInputs').on('keyup change', function () {
        table
            .column(myJQuery(this).closest('th').index())
            .search((<any>this).value)
            .draw();
    });
    //});
}