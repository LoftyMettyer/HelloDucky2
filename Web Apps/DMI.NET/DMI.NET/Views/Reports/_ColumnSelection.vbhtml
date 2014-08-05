@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.ViewModels.Reports
@Imports DMI.NET.Code.Extensions
@Imports DMI.NET.Models
@Inherits System.Web.Mvc.WebViewPage(Of ReportBaseModel)

@Html.HiddenFor(Function(m) m.ColumnsAsString, New With {.id = "txtCSAAS"})
<fieldset>
	<fieldset id="columnsAvailable">
		<legend class="fontsmalltitle">Columns / Calculations Available :</legend>
		<fieldset id="columncalculations">
			<legend>
				@Html.RadioButton("columnSelectiontype", CInt(ColumnSelectionType.Columns), True, New With {.onclick = "toggleColumnsCalculations('column')"})
				<span style="padding-right:30px">Columns</span>
				@Html.RadioButton("columnSelectiontype", CInt(ColumnSelectionType.Calculations), False, New With {.onclick = "toggleColumnsCalculations('calc')"})	Calculations
			</legend>
			<select name="SelectedTableID" id="SelectedTableID" onchange="getAvailableTableColumnsCalcs();"></select>
			<table id="AvailableColumns"></table>
		</fieldset>
	</fieldset>

	<fieldset id="columnbuttons">
		<div id="colbtngrp1">
			<input type="button" id="btnColumnAdd" value="Add" onclick="addColumnToSelected(0);" />
			<input type="button" id="btnColumnAddAll" value="Add All" onclick="addAllColumnsToSelected();" />
		</div>
		<div id="colbtngrp2">
			<input type="button" id="btnColumnRemove" value="Remove" onclick="removeSelectedColumn();" />
			<input type="button" id="btnColumnRemoveAll" value="Remove All" onclick="removeAllSelectedColumns();" />
		</div>
		<div id="colbtngrp3">
			<input type="button" id="btnColumnMoveUp" value="Move Up" disabled onclick="moveSelectedColumn('up');" />
			<input type="button" id="btnColumnMoveDown" value="Move Down" disabled onclick="moveSelectedColumn('down');" />
		</div>
	</fieldset>

	<fieldset id="columnsSelected">
		<legend class="fontsmalltitle">Columns / Calculations Selected :</legend>
		<table id="SelectedColumns" class="scroll" cellpadding="0" cellspacing="0"></table>
		<br />

		<div class="customReportsOnly" style="margin-left:20px;">
			<div class="numericOnly width35 floatleft" style="color: rgb(0, 0, 0)">
				<div class="width100">
					<input class=" ui-widget ui-corner-all" id="SelectedColumnIsAverage" onchange="updateColumnsSelectedGrid();" type="checkbox">
					<label for="SelectedColumnIsAverage">Average</label>
				</div>
				<div class="width100">
					<input class="ui-widget ui-corner-all" id="SelectedColumnIsCount" onchange="updateColumnsSelectedGrid();" type="checkbox">
					<label for="SelectedColumnIsCount">Count</label>
				</div>
				<div class="width100">
					<input class="ui-widget ui-corner-all" id="SelectedColumnIsTotal" onchange="updateColumnsSelectedGrid();" type="checkbox">
					<label for="SelectedColumnIsTotal">Total</label>
				</div>
			</div>

			<div>
				<div class="width65 floatleft">
					<input class="ui-widget ui-corner-all" id="SelectedColumnIsHidden" onchange="updateColumnsSelectedGrid();" type="checkbox">
					<label for="SelectedColumnIsHidden">Hidden</label>
					<div class="numericOnly" style="color: rgb(0, 0, 0);">
						<input class="ui-widget ui-corner-all" id="SelectedColumnIsGroupWithNext" onchange="updateColumnsSelectedGrid();" type="checkbox">
						<label for="SelectedColumnIsGroupWithNext">Group With Next</label>
					</div>
					<div class="baseTableOnly" style="color: rgb(165, 147, 147);">
						<input disabled="disabled" class="ui-widget ui-corner-all" id="SelectedColumnIsRepeated" onchange="updateColumnsSelectedGrid();" type="checkbox">
						<label for="SelectedColumnIsRepeated">Repeat on child rows</label>
					</div>
				</div>
			</div>
		</div>
	</fieldset>
</fieldset>
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
	
			datarow.Name = $("#SelectedTableID option:selected").text() + '.' + datarow.Name;
			datarow.ReportType = '@Model.ReportType';
			datarow.ReportID = '@Model.ID';
			datarow.Heading = datarow.Name;
			datarow.Sequence = $("#SelectedColumns").jqGrid('getGridParam', 'records') + 1;
			datarow.IsAverage = false;
			datarow.IsCount = false;
			datarow.IsTotal = false;
			datarow.IsHidden = false;
			datarow.IsGroupWithNext = false;
			datarow.IsRepeated = false;
			datarow.TableID = $("#SelectedTableID option:selected").val();

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
    	var sType;

    	$("#AvailableColumns").jqGrid('GridUnload');

    	if ($("#columnSelectiontype:checked").val() == 0) {
    		sType = "C";
    	}
    	else {
    		sType = "E";
			}

    	$("#AvailableColumns").jqGrid({
    		url: 'Reports/GetAvailableItemsForTable?TableID=' + $("#SelectedTableID").val() + '&&ReportID=' + '@Model.ID' + '&&ReportType=' + '@Model.ReportType' + '&&selectionType=' + sType,
    		datatype: 'json',
    		mtype: 'GET',
    		jsonReader: {
    			root: "rows", //array containing actual data
    			page: "page", //current page
    			total: "total", //total pages for the query
    			records: "records", //total number of records
    			repeatitems: false,
    			id: "ID"
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
    		width: 300,
    		height: 320,
    		sortname: 'Name',
    		sortorder: "desc",
    		rowNum: 10000,
    		ondblClickRow: function (rowid) {
    			addColumnToSelected(rowid);
    		}
    	});    	
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
    	dataRow.IsRepeated = $("#SelectedColumnIsRepeated").is(':checked');

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
    		colNames: ['ID', 'TableID', 'IsExpression',	'Name',	'Sequence',	'Heading',	'DataType',
    							'Size', 'Decimals', 'IsAverage', 'IsCount', 'IsTotal', 'IsHidden', 'IsGroupWithNext', 'IsRepeated', 'ReportID', 'ReportType'],
    		colModel: [
					{ name: 'ID', index: 'ID', hidden: true },
					{ name: 'TableID', index: 'TableID', hidden: true },
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
					{ name: 'IsRepeated', index: 'IsRepeated', hidden: true },
    			{ name: 'ReportID', index: 'ReportID', hidden: true },
					{ name: 'ReportType', index: 'ReportType', hidden: true }],
    		viewrecords: true,
    		width: 300,
    		height: 320,
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
    			var isBaseOrParentTableColumn = (dataRow.TableID == $("#BaseTableID").val()) || (dataRow.TableID == $("#txtParent1ID").val()) || (dataRow.TableID == $("#txtParent2ID").val());
    			var isThereChildColumns = true;

    			$(".numericOnly :input").attr("disabled", !isNumeric);
    			$(".baseTableOnly :input").attr("disabled", !isBaseOrParentTableColumn);

    			if (isNumeric) {
    				$(".numericOnly").css("color", "#000000");
    			} else {
    				$(".numericOnly").css("color", "#A59393");
    			}

    			if (isBaseOrParentTableColumn && isThereChildColumns) {
    				$(".baseTableOnly").css("color", "#000000");
    			} else {
    				$(".baseTableOnly").css("color", "#A59393");
    			}

    			$("#SelectedColumnHeading").val(dataRow.Heading);

    			$("#SelectedColumnSize").val(dataRow.Size);
    			$("#SelectedColumnDecimals").val(dataRow.Decimals);
    			$('#SelectedColumnIsAverage').prop('checked', JSON.parse(dataRow.IsAverage));
    			$('#SelectedColumnIsCount').prop('checked', JSON.parse(dataRow.IsCount));
    			$('#SelectedColumnIsTotal').prop('checked', JSON.parse(dataRow.IsTotal));
    			$('#SelectedColumnIsHidden').prop('checked', JSON.parse(dataRow.IsHidden));
    			$('#SelectedColumnIsGroupWithNext').prop('checked', JSON.parse(dataRow.IsGroupWithNext));
    			$('#SelectedColumnIsRepeated').prop('checked', JSON.parse(dataRow.IsRepeated));

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

			$(".spinner").spinner({
				min: 0,
				max: 10,
				showOn: 'both'
			}).css("width", "15px");


    	if ('@Model.ReportType' == '@UtilityType.utlMailMerge') {
    		$(".customReportsOnly").hide();
    	}
    });

  </script>
