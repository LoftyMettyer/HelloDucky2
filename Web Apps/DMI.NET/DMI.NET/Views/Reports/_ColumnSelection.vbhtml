@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Enums
@imports HR.Intranet.Server.Enums
@Imports DMI.NET.ViewModels.Reports
@Inherits System.Web.Mvc.WebViewPage(Of ColumnsViewModel)

<style>

	#columnsAvailable {
		height: 500px;
		overflow: auto;
	}

</style>

@*Look at replacing with proper jqgrid with a subgrid*@
@*http://trirand.com/blog/jqgrid/jqgrid.html*@

<div id="columnsAvailable" style="float:left">

	Columns / Calculations Available :
	<br/>

	@Code
		If Model.DisplayTableSelection Then
			@<select name="SelectedTableID" id="SelectedTableID" onclick="getAvailableTableColumnsCalcs();"></select>
End If
End Code

	<br />
	@Html.RadioButton("columnSelectiontype", ColumnSelectionType.Columns, True, New With {.onclick = "toggleColumnsCalculations('column')"})
	Columns
	@Html.RadioButton("columnSelectiontype", ColumnSelectionType.Calculations, True, New With {.onclick = "toggleColumnsCalculations('calc')"})
	Calculations
	<br/>

	<table id="AvailableColumns" class="scroll" cellpadding="0" cellspacing="0"></table>
</div>

<div style="float:left">
	<input type="button" id="btnColumnAdd" value="Add" onclick="addColumnToSelected(0);" />
	<br />
	<input type="button" id="btnColumnAddAll" value="Add All" onclick="addAllColumnsToSelected();" />
	<br />
	<input type="button" id="btnColumnRemove" value="Remove" onclick="removeSelectedColumn();" />
	<br />
	<input type="button" id="btnColumnRemoveAll" value="Remove All" onclick="removeAllSelectedColumns();" />
	<br />
	<input type="button" id="btnColumnMoveUp" value="Move Up" />
	<br />
	<input type="button" id="btnColumnMoveDown" value="Move Down" />
</div>

<div id="columnsSelected" style="float:left">
	@Html.SelectedReportColumns("Columns.Selected", Model.Selected, Nothing)
</div>

<input type="hidden" name="Columns.BaseTableID" value="@Model.BaseTableID" />


  <script type="text/javascript">

		function toggleColumnsCalculations(type) {
		}

		function addColumnToSelected(rowID) {

			if (rowID == 0) {
				iRowID = $("#AvailableColumns").getGridParam('selrow');
			}

			var gridData = $('#AvailableColumns').getRowData(rowID);



			$('#AvailableColumns').trigger('reloadGrid');

		}

		function addAllColumnsToSelected() {
			//TODO
		}

		function changeAvailableReportTable(event) {

			// Warn user
			// Reload the available columns

		}

		function removeSelectedColumn() {
			//TODO
		}

    function columndefinition_rowcolchange() {

    	iRowID = $("#ColumnsSelected").getGridParam('selrow') - 1;
    	$("[id^=columnproperty], [id$=Breakdown]").hide();
    	$("#columnproperty" + iRowID + "Breakdown").show();

    }

    function getAvailableTableColumnsCalcs() {

    	$("#AvailableColumns").jqGrid({
    		url: 'Reports/GetColumnsForTable?TableID=' & $("#SelectedTableID").val(),
    		datatype: 'json',
    		mtype: 'GET',
    		jsonReader: {
    			root: "rows", //array containing actual data
    			page: "page", //current page
    			total: "total", //total pages for the query
    			records: "records", //total number of records
    			repeatitems: false,
    			id: "id" //index of the column with the PK in it
    		},
    		colNames: ['id', 'Name'],
    		colModel: [
					{ name: 'id', index: 'id', hidden: true },
					{ name: 'Name', index: 'Name', width: 40, sortable: false }],
    		viewrecords: true,
    		width: 400,
    		sortname: 'Name',
    		sortorder: "desc",
    		ondblClickRow: function (rowid) {

    			$(this).jqGrid('editGridRow', rowid);
    			// move to selected

    			addColumnToSelectedO(rowid)

    		}
    	});
    	
    }


		// Initialise
    $(function () {

    	$("[id^=columnproperty], [id$=Breakdown]").hide();
    	$("#columnproperty0Breakdown").show();

    	tableToGrid("#ColumnsSelected", {
    		onSelectRow: function (rowID) {
    			columndefinition_rowcolchange();
    		},
    		colNames: ['id', 'Name'],
    		colModel: [
					{ name: 'id', hidden: true },
					{ name: 'Name', sortable: false }
    		],
    		cmTemplate: { sortable: false },
    		rowNum: 1000
    	});


      $('#viewSelectedColumns').change(function () {
        var value = $(this).val();

        showThisColumnProperties(value);

      });


    });

  </script>
