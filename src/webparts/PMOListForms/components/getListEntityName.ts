import * as $ from 'jquery';
export var listType: any;

export function _getListEntityName(context, listGUID){
    $.ajax({
        url:  context.pageContext.web.absoluteUrl+ "/_api/Web/Lists('"+ listGUID +"')/ListItemEntityTypeFullName",  
        method: "GET",
        headers: {
        accept: "application/json;odata=verbose",
        },
        success:(data, status, xhr) => 
        {  
            listType = data.d.ListItemEntityTypeFullName;
        },  
        error: (xhr, status, error)=>
        {  
            console.log(xhr);
        }  
    });
}