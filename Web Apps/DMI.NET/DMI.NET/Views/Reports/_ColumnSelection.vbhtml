@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Enums
@imports HR.Intranet.Server.Enums
@Imports DMI.NET.ViewModels.Reports
@Imports DMI.NET.Code.Extensions
@Imports DMI.NET.Models
@Inherits System.Web.Mvc.WebViewPage(Of ReportBaseModel)

<style>

	#columnsAvailable {
		height: 800px;
		overflow: auto;
	}

</style>

@Html.HiddenFor(Function(m) m.ColumnsAsString, New With {.id = "txtCSAAS"})


<div id="columnsAvailable" style="float:left">

	Columns / Calculations Available :
	<br/>

	<select name="SelectedTableID" id="SelectedTableID" onchange="getAvailableTableColumnsCalcs();"></select>

	<br />
	@Html.RadioButton("columnSelectiontype", CInt(ColumnSelectionType.Columns), True, New With {.onclick = "toggleColumnsCalculations('column')"})
	Columns
	@Html.RadioButton("columnSelectiontype", CInt(ColumnSelectionType.Calculations), False, New With {.onclick = "toggleColumnsCalculations('calc')"})
	Calculations
	<br/>

	<table id="AvailableColumns" class="scroll" cellpadding="0" cellspacing="0"></table>
</div>

<div style="float:left">
	<input type="button" id="btnColumnAdd" value="Add" onclick="addColumnToSelected(0);" />
	<br />
	<input type="button" id="btnColumnAddAll" value="Add All" onclick="addAllColumnsToSelected();" />
	<br />
	<br />
	<input type="button" id="btnColumnRemove" value="Remove" disabled onclick="removeSelectedColumn();" />
	<br />
	<input type="button" id="btnColumnRemoveAll" value="Remove All" disabled onclick="removeAllSelectedColumns();" />
	<br />
	<br />
	<div class="customReportsOnly">
		<input type="button" id="btnColumnMoveUp" value="Move Up" disabled onclick="moveSelectedColumn('up');" />
		<br />
		<input type="button" id="btnColumnMoveDown" value="Move Down" disabled onclick="moveSelectedColumn('down');" />
	</div>
</div>

<div id="columnsSelected" style="float:left">
	<table id="SelectedColumns" class="scroll" cellpadding="0" cellspacing="0"></table>
	<br/>

	<div>
		<div class="customReportsOnly">
			<label for="SelectedColumnHeading">Heading:</label>
			<input type='text' id="SelectedColumnHeading" onchange="updateColumnsSelectedGrid();" />
		</div>
		<br />

		Size:
		<input type='text' id="SelectedColumnSize" onchange="updateColumnsSelectedGrid();" />

		<div class="decimalsOnly">
			<label for="SelectedColumnDecimals">Decimals:</label>
			<input type='text' id="SelectedColumnDecimals" onchange="updateColumnsSelectedGrid();" />
		</div>
		<br />

		<div class="customReportsOnly">
			<div class="numericOnly">
				<label for="SelectedColumnIsAverage">Average:</label>
				<input type="checkbox" id="SelectedColumnIsAverage" onchange="updateColumnsSelectedGrid();" />

				<label for="SelectedColumnIsCount">Count:</label>
				<input type='checkbox' id="SelectedColumnIsCount" onchange="updateColumnsSelectedGrid();" />

				<label for="SelectedColumnIsTotal">Total:</label>
				<input type='checkbox' id="SelectedColumnIsTotal" onchange="updateColumnsSelectedGrid();" />
			</div>
			<br />

			<label for="SelectedColumnIsHidden">Hidden:</label>
			<input type='checkbox' id="SelectedColumnIsHidden" onchange="updateColumnsSelectedGrid();" />

			<div class="numericOnly">
				<label for="SelectedColumnIsGroupWithNext">Group With Next:</label>
				<input type='checkbox' id="SelectedColumnIsGroupWithNext" onchange="updateColumnsSelectedGrid();" />
			</div>

		</div>
	</div>

</div>

<input type="hidden" name="Columns.BaseTableID" value="@Model.BaseTableID" />


  <script type="text/javascript">

		function moveSelectedColumn(direction) {
			OpenHR.MoveItemInGrid($("#SelectedColumns"), direction);
		}

		function toggleColumnsCalculations(type) {
			getAvailableTableColumnsCalcs();
		}

		function addColumnToSelected(rowID) {

			if (rowID == 0) {
				rowID = $("#AvailableColumns").getGridParam('selrow');
			}

			var datarow = $("#AvailableColumns").getRowData(rowID);
	
			datarow.ReportType = '@Model.ReportType';
			datarow.ReportID = '@Model.ID';
			datarow.Heading = datarow.Name;
			datarow.Sequence = $("#SelectedColumns").jqGrid('getGridParam', 'records') + 1;
			datarow.IsAverage = false;
			datarow.IsCount = false;
			datarow.IsTotal = false;
			datarow.IsHidden = false;
			datarow.IsGroupWithNext = false;

			OpenHR.postData("Reports/AddReportColumn", datarow);

			$("#SelectedColumns").jqGrid('addRowData', datarow.ID, datarow);
			$('#SelectedColumns').jqGrid("setSelection", rowID);

			$("#AvailableColumns").jqGrid('delRowData', rowID);
			button_disable($("#btnSortOrderAdd")[0], false);

		}

		function addAllColumnsToSelected() {

			var rows = $("#AvailableColumns").jqGrid('getDataIDs');

			for (var i = 0; i < rows.length; i++) {
				var datarow = $("#AvailableColumns").getRowData(rows[i]);
				addColumnToSelected(datarow.ID);
			}

		}

		function removeSelectedColumn() {

			rowID = $("#SelectedColumns").getGridParam('selrow');
			var datarow = $("#SelectedColumns").getRowData(rowID);

			OpenHR.postData("Reports/RemoveReportColumn", datarow);

			$("#AvailableColumns").jqGrid('addRowData', datarow.ID, datarow);
			$("#AvailableColumns").jqGrid("sortGrid", "Name", true)
			$("#SelectedColumns").jqGrid('delRowData', rowID);

		}

    function getAvailableTableColumnsCalcs() {

    	var URL;

    	$("#AvailableColumns").jqGrid('GridUnload');

    	if ($("#columnSelectiontype:checked").val() == 0) {
				URL = 'Reports/GetColumnsForTable?TableID=' + $("#SelectedTableID").val();
    	}
    	else {
				URL = 'Reports/GetCalculationsForTable?TableID=' + $("#SelectedTableID").val();
			}

    	$("#AvailableColumns").jqGrid({
    		url: URL,
    		datatype: 'json',
    		mtype: 'GET',
    		jsonReader: {
    			root: "rows", //array containing actual data
    			page: "page", //current page
    			total: "total", //total pages for the query
    			records: "records", //total number of records
    			repeatitems: false,
    			id: "ID" //index of the column with the PK in it
    		},
    		colNames: ['ID', 'IsExpression', 'Name', 'DataType', 'Size', 'Decimals'],
    		colModel: [
					{ name: 'ID', index: 'ID', hidden: true },
					{ name: 'IsExpression', index: 'IsExpression3', hidden: true },
					{ name: 'Name', index: 'Name', width: 40, sortable: false },
    			{ name: 'DataType', index: 'DataType', hidden: true },
					{ name: 'Size', index: 'Size',  hidden: true },
					{ name: 'Decimals', index: 'Decimals', hidden: true }],
    		viewrecords: true,
    		width: 400,
    		sortname: 'Name',
    		sortorder: "desc",
    		rowNum: 10000,
    		ondblClickRow: function (rowid) {
    			addColumnToSelected(rowid);
    		}
    	});
    	
    	// TODO Loop through available removing any currently selected

    }

    function updateColumnsSelectedGrid() {

    	var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');
    	var dataRow = $('#SelectedColumns').jqGrid('getRowData', rowId);

    	dataRow.Heading = $("#SelectedColumnHeading").val();
    	dataRow.Size = $("#SelectedColumnSize").val();
    	dataRow.Decimals = $("#SelectedColumnDecimals").val();
    	dataRow.IsAverage = $('#SelectedColumnIsAverage').is(':checked');
    	dataRow.IsCount = $("#SelectedColumnIsCount").is(':checked');
    	dataRow.IsTotal = $("#SelectedColumnIsTotal").is(':checked');
    	dataRow.IsHidden = $("#SelectedColumnIsHidden").is(':checked');
    	dataRow.IsGroupWithNext = $("#SelectedColumnIsGroupWithNext").is(':checked');
	
    	$('#SelectedColumns').jqGrid('setRowData', rowId, dataRow);

    }

    function attachGridToSelectedColumns() {

    	$("#SelectedColumns").jqGrid({
    		datatype: "jsonstring",
    		datastr: '@Model.Columns.ToJsonResult',
    		mtype: 'GET',
    		jsonReader: {
    			root: "rows", //array containing actual data
    			page: "page", //current page
    			total: "total", //total pages for the query
    			records: "records", //total number of records
    			repeatitems: false,
    			id: "Sequence" //index of the column with the PK in it
    		},
    		colNames: ['ID', 'IsExpression',	'Name',	'Sequence',	'Heading',	'DataType',
    							'Size', 'Decimals', 'IsAverage', 'IsCount', 'IsTotal', 'IsHidden', 'IsGroupWithNext', 'ReportID', 'ReportType'],
    		colModel: [
					{ name: 'ID', index: 'ID', hidden: true },
					{ name: 'IsExpression', index: 'IsExpression2', hidden: true},
					{ name: 'Name', index: 'Name', sortable: false},
					{ name: 'Sequence', index: 'Sequence', hidden: true },
					{ name: 'Heading', index: 'Heading', hidden: true },
					{ name: 'DataType', index: 'DataType', hidden: true },
					{ name: 'Size', index: 'Size', hidden: true },
					{ name: 'Decimals', index: 'Decimals', hidden: true },
					{ name: 'IsAverage', index: 'IsAverage', hidden: true },
					{ name: 'IsCount', index: 'IsCount', hidden: true },
					{ name: 'IsTotal', index: 'IsTotal', hidden: true },
					{ name: 'IsHidden', index: 'IsHidden', hidden: true },
					{ name: 'IsGroupWithNext', index: 'IsGroupWithNext', hidden: true },
    			{ name: 'ReportID', index: 'ReportID', hidden: true },
					{ name: 'ReportType', index: 'ReportType', hidden: true }],
    		viewrecords: true,
    		width: 400,
    		sortname: 'Sequence',
    		sortorder: "asc",
    		rowNum: 10000,
    		beforeSelectRow: function(id) {
    			updateColumnsSelectedGrid();
    			return true;
    		},
    		onSelectRow: function (id) {

    			var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');
    			var dataRow = $("#SelectedColumns").getRowData(rowId)
    			var allRows = $("#SelectedColumns")[0].rows;

    			var isTopRow = (rowId == allRows[1].id);
    			var isBottomRow = (rowId == allRows[allRows.length - 1].id);
    			var isNumeric = (dataRow.DataType == '2' || dataRow.DataType == '4');

    			$(".numericOnly :input").attr("disabled", !isNumeric);

    			if (isNumeric) {
    				$(".numericOnly").css("color", "#000000");
    			} else {
    				$(".numericOnly").css("color", "#A59393");
    			}


    			$("#SelectedColumnHeading").val(dataRow.Heading);

    			$("#SelectedColumnSize").val(dataRow.Size);
    			$("#SelectedColumnDecimals").val(dataRow.Decimals);
    			$('#SelectedColumnIsAverage').prop('checked', JSON.parse(dataRow.IsAverage));
    			$('#SelectedColumnIsCount').prop('checked', JSON.parse(dataRow.IsCount));
    			$('#SelectedColumnIsTotal').prop('checked', JSON.parse(dataRow.IsTotal));
    			$('#SelectedColumnIsHidden').prop('checked', JSON.parse(dataRow.IsHidden));
    			$('#SelectedColumnIsGroupWithNext').prop('checked', JSON.parse(dataRow.IsGroupWithNext));

    			// Enable / Disable relevant buttons
    			button_disable($("#btnColumnRemove")[0], false);
    			button_disable($("#btnColumnRemoveAll")[0], false);
    			button_disable($("#btnColumnMoveUp")[0], isTopRow);
    			button_disable($("#btnColumnMoveDown")[0], isBottomRow);

    		},
    		gridComplete: function () {
					// Highlight top row
    			var ids = $(this).jqGrid("getDataIDs");
    			if (ids && ids.length > 0)
    				$(this).jqGrid("setSelection", ids[0]);
    		}
    	});

    	$("#SelectedColumns").jqGrid('sortableRows');


    }

		// Initialise
    $(function () {

    	if ('@Model.ReportType' == '@UtilityType.utlMailMerge') {
    		$(".customReportsOnly").hide();
    	}

    	attachGridToSelectedColumns();

    });

  </script>
