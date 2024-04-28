import * as $ from 'jquery';
import 'bootstrap/dist/css/bootstrap.css';
import "datatables.net-dt/css/jquery.dataTables.min.css";
import "datatables.net/js/jquery.dataTables.min.js";

export function clearDataTable(tableId: string) {
  if ((<any>$).fn.dataTable.isDataTable('#' + tableId)) {
    var table = ($('#' + tableId) as any).DataTable();
    table
      .clear()
      .draw();
  }
}

export function destroyDataTable(tableId: string) {
  if ((<any>$).fn.dataTable.isDataTable('#' + tableId)) {
    var table = ($('#' + tableId) as any).DataTable();
    table.destroy();
  }
}

export function orderingSingleColumn(tableId: string) {
  if ((<any>$).fn.dataTable.isDataTable('#' + tableId)) {
    var table = ($('#' + tableId) as any).DataTable();
    // Sort by column 1 and then re-draw
    table
        .order( [ 0, 'asc' ] )
        .draw();
  }
}

export function orderingMultipleColumn(tableId: string) {
  if ((<any>$).fn.dataTable.isDataTable('#' + tableId)) {
    var table = ($('#' + tableId) as any).DataTable();
    // Sort by columns 1 and 2 and redraw
    table
        .order( [ 1, 'asc' ], [ 2, 'asc' ] )
        .draw();
  }
}

export async function applyDataTableNoPagingFilter(tableId: string) {
  try {
    let oTable: any;
    oTable = ($('#' + tableId) as any).DataTable({
      "bDestroy": true,
      "responsive": true,
      "autoWidth": true,
      "bInfo": false,
      "bLengthChange": false,
      "bFilter": false,
      "bPaginate": false,
      "language": {
        paginate: {
          "next": '<i class="fa fa-angle-right" aria-hidden="true"></i>',
          "previous": '<i class="fa fa-angle-left" aria-hidden="true"></i>',
          "first": '<i class="fa fa-angle-double-left" aria-hidden="true"></i>',
          "last": '<i class="fa fa-angle-double-right" aria-hidden="true"></i>',
        }
      },
    });
  } catch (error) {
    return error;
  }
}

/**
   * @param {string} tableId
   * @param {any[]} columns
   * @param {any[]} rows
   * @returns
   * @memberof spservices
   */

export async function applyDataTableGroupBy(tableId: string, groupColumn: number, effortColumn: number, colspan1: number, colspan2: number) {
  try {
    let oTable: any;
    oTable = ($('#' + tableId) as any).DataTable({
      "bDestroy"    : true,
      "bInfo"       : false,
      "oLanguage"   : {
                        "sLengthMenu": "_MENU_ Per page",
                      }, 
      "columnDefs"  : [
                        { "visible": false, "targets": groupColumn }
                      ],
      "order"       : [[ groupColumn, 'asc' ]],
      "drawCallback": function ( settings ) {
        var api   = this.api();
        var rows  = api.rows( {page:'current'} ).nodes();
        var last  = null;     
        var sum   = 0; 
        api.column(groupColumn, {page:'current'} ).data().each( function ( group, i ) {
          if ( last !== group ) {
            api.rows().data().each( function(item){
              if (item[groupColumn] == group){
                sum = sum + (+item[effortColumn]);
              }
            });
            $(rows).eq( i ).before(
              '<tr class="group"><td colspan="'+colspan1+'" style="background-color: #e7e7e7;"><b>'+group+'</b></td><td colspan="'+colspan2+'" style="background-color: #e7e7e7;"><b>'+ sum.toFixed(2) +'</b></td></tr>'
            );
            last = group;
            sum  = 0;
          }
        });
      }
    });
  } 
  catch (error) {
    return error;
  }
}

export async function applyDataTable(tableId: string, rows: any[] = [], columns: any[] = [], searchElemId: string = "", pageLenSelectElem: string = "", searchBtnElemId: string = "") {
  try {
    let oTable: any;
    if (columns.length > 0) {debugger
      oTable = ($('#' + tableId) as any).DataTable({
        "bDestroy": true,
        "responsive": true,
        "autoWidth": true,
        "bInfo": false,
        "bLengthChange": false,
        "bFilter": false,
        "oLanguage": {
          "sLengthMenu": "_MENU_ Per page",
        },    
        "order": [],
        //"scrollX": true,
        "language": {
          paginate: {
            "next": '<span>NEXT <i class="fa fa-arrow-right"></i></span>',
            "previous": '<span><i class="fa fa-arrow-left"></i> PREV</span>',
            "first": '<i class="fa fa-angle-double-left" aria-hidden="true"></i>',
            "last": '<i class="fa fa-angle-double-right" aria-hidden="true"></i>',
          }
        },
        //searching: searchElemId == "" ? true : false,
        data: rows,
        columns: columns
      }); 
    } else if (rows.length > 0) {
      oTable = ($('#' + tableId) as any).DataTable({
        "bDestroy": true,
        "responsive": true,
        "autoWidth": true,
        "bInfo": false,
        "bLengthChange": false,
        "bFilter": false,
        "oLanguage": {
          "sLengthMenu": "_MENU_ Per page",
        },
        "order": [],
        //"scrollX": true,
        "language": {
          paginate: {
            "next": '<span>NEXT <i class="fa fa-arrow-right"></i></span>',
            "previous": '<span><i class="fa fa-arrow-left"></i> PREV</span>',
            "first": '<i class="fa fa-angle-double-left" aria-hidden="true"></i>',
            "last": '<i class="fa fa-angle-double-right" aria-hidden="true"></i>',
          }
        },
        //searching: searchElemId == "" ? true : false,
        data: rows
      }); 
    } else {
      oTable = ($('#' + tableId) as any).DataTable({
        "bDestroy": true,
        "responsive": true,
        "autoWidth": true,
        "bInfo": false,
        "order": [],
        "oLanguage": {
          "sLengthMenu": "_MENU_ Per page",
        },       
        //"scrollX": true,
        //"bLengthChange": false,
        //"bFilter": false,
        "language": {
          paginate: {
            "next": '<span>NEXT <i class="fa fa-arrow-right"></i></span>',
            "previous": '<span><i class="fa fa-arrow-left"></i> PREV</span>',
            "first": '<i class="fa fa-angle-double-left" aria-hidden="true"></i>',
            "last": '<i class="fa fa-angle-double-right" aria-hidden="true"></i>',
          }
        },
        //searching: searchElemId == "" ? true : false,
      });
    }

    if (searchElemId != "") {
      $("#" + tableId + "_wrapper .dataTables_filter").hide();
      //pay attention to capital D, which is mandatory to retrieve "api" datatables' object, as @Lionel said
      if (searchBtnElemId == "") {
        $('#' + searchElemId).keyup(function (event) {
          if (event.which == 13) {
            event.preventDefault();
          } else {
            oTable.search($(this).val()).draw();
          }
        });
      } else {
        $('#' + searchBtnElemId).on('click', function () {
          var val = $('#' + searchElemId).val();
          oTable.search(val).draw();
        });
      }
    }

    if (pageLenSelectElem != "") {
      $("#" + tableId + "_wrapper .dataTables_length").hide();
      $('#' + pageLenSelectElem).on('change', function () {
        var len = $(this).val();
        oTable.page.len(len).draw();
      });
    }
  } catch (error) {
    return error;
  }
}