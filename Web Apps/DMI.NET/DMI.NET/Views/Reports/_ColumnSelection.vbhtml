﻿@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.ViewModels.Reports
@Imports DMI.NET.Code.Extensions
@Imports DMI.NET.Models
@Inherits System.Web.Mvc.WebViewPage(Of ReportBaseModel)

@Html.HiddenFor(Function(m) m.ColumnsAsString, New With {.id = "txtCSAAS"})

<div class="nowrap">

	<div class="tablerow">
		<fieldset id="selectedTable">
			<legend class="fontsmalltitle width100">Base Table :</legend>
			<select name="SelectedTableID" id="SelectedTableID" class="enableSaveButtonOnComboChange" onchange="getAvailableTableColumnsCalcs();"></select>
			<br />
			@Html.RadioButton("columnSelectiontype", CInt(ColumnSelectionType.Columns), True,
												New With {.onclick = "toggleColumnsCalculations('column')", .class = "radioColumnType", .id = "columnSelectiontype_0"})
			<span style="padding-right:30px">Columns</span>
			@Html.RadioButton("columnSelectiontype", CInt(ColumnSelectionType.Calculations), False,
												New With {.onclick = "toggleColumnsCalculations('calc')", .class = "radioColumnType", .id = "columnSelectiontype_1"})	Calculations
		</fieldset>
	</div>

	<div class="tablerow coldefinition">
		<div class="tablecell">
			<fieldset id="columnsAvailable">
				<legend class="fontsmalltitle">Columns / Calculations Available :</legend>
				<table id="AvailableColumns"></table>
			</fieldset>
		</div>

		<div class="tablecell">
			<fieldset class="" id="columnbuttons">
				<div id="colbtngrp1">
					<input type="button" id="btnColumnAdd" class="enableSaveButtonOnClick" value="Add" onclick="addColumnToSelected();" />
					<input type="button" id="btnColumnAddAll" class="enableSaveButtonOnClick" value="Add All" onclick="addAllColumnsToSelected();" />
				</div>
				<div id="colbtngrp2">
					<input type="button" id="btnColumnRemove" value="Remove" onclick="requestRemoveSelectedColumns();" />
					<input type="button" id="btnColumnRemoveAll" class="enableSaveButtonOnClick" value="Remove All" onclick="requestRemoveAllSelectedColumns();" />
				</div>
				<div id="colbtngrp3" class="customReportsOnly">
					<input type="button" id="btnColumnMoveUp" value="Move Up" disabled onclick="moveSelectedColumn('up');" />
					<input type="button" id="btnColumnMoveDown" class="enableSaveButtonOnClick" value="Move Down" disabled onclick="moveSelectedColumn('down');" />
				</div>
			</fieldset>
		</div>

		<div class="tablecell">
			<fieldset class="left" id="columnsSelected">
				<legend class="fontsmalltitle">Columns / Calculations Selected :</legend>
				<table id="SelectedColumns" class="scroll" cellpadding="0" cellspacing="0"></table>
			</fieldset>
		</div>
	</div>
	<div class="tablerow coldefinition">
		<div class="tablecell">
			<fieldset class="customReportsOnly" id="CustomDefinitionReportOptions">
				<legend class="fontsmalltitle">Report Options :</legend>
				<div>
					@Html.CheckBoxFor(Function(m) m.IsSummary)
					@Html.LabelFor(Function(m) m.IsSummary)
				</div>
				<div>
					@Html.CheckBoxFor(Function(m) m.IgnoreZerosForAggregates)
					@Html.LabelFor(Function(m) m.IgnoreZerosForAggregates)
				</div>
			</fieldset>
		</div>
		<div class="tablecell"></div>
		<div class="tablecell">
			<fieldset>
				<div id="definitionColumnProperties">
					<div class="formfieldfill customReportsOnly">
						<label for="SelectedColumnHeading">Heading :</label>
						<span><input type='text' id="SelectedColumnHeading" maxlength="50" onchange="updateColumnsSelectedGrid();" /></span>
					</div>
					<div class="formfieldfill">
						<label for="SelectedColumnSize">Size :</label>
						<span><input class="" type='number' id="SelectedColumnSize" onchange="updateColumnsSelectedGrid();" /></span>
					</div>
					<div class="formfieldfill decimalsOnly">
						<label for="SelectedColumnDecimals">Decimals :</label>
						<span><input class="" type='number' id="SelectedColumnDecimals" onchange="updateColumnsSelectedGrid();" /></span>
					</div>

					<div class="tablelayout customReportsOnly colAggregates">
						<div class="tablerow" style="color: rgb(0, 0, 0)">
							<div class="tablecell numericOnly">
								<input class=" ui-widget ui-corner-all" id="SelectedColumnIsAverage" onchange="updateColumnsSelectedGrid();" type="checkbox">
								<label for="SelectedColumnIsAverage">Average</label>
							</div>
							<div class="tablecell cannotBeHidden">
								<input class="ui-widget ui-corner-all" id="SelectedColumnIsCount" onchange="updateColumnsSelectedGrid();" type="checkbox">
								<label for="SelectedColumnIsCount">Count</label>
							</div>
							<div class="tablecell numericOnly">
								<input class="ui-widget ui-corner-all" id="SelectedColumnIsTotal" onchange="updateColumnsSelectedGrid();" type="checkbox">
								<label for="SelectedColumnIsTotal">Total</label>
							</div>
						</div>


						<div class="tablerow">
							<div>
								<input class="ui-widget ui-corner-all" id="SelectedColumnIsHidden" onchange="changeColumnIsHidden();" type="checkbox">
								<label id="labelSelectedColumnIsHidden" for="SelectedColumnIsHidden">Hidden</label>
							</div>
							<div class="tablecell canGroupWithNext" style="color: rgb(0, 0, 0);">
								<input class="ui-widget ui-corner-all" id="SelectedColumnIsGroupWithNext" onchange="changeColumnIsGroupWithNext();" type="checkbox">
								<label id="labelSelectedColumnIsGroupWithNext" for="SelectedColumnIsGroupWithNext">Group with next</label>
							</div>
							<div class="tablecell baseTableOnly">
								<input class="ui-widget ui-corner-all" id="SelectedColumnIsRepeated" onchange="updateColumnsSelectedGrid();" type="checkbox">
								<label id="labelSelectedColumnRepeatOnChild" for="SelectedColumnIsRepeated">Repeat on child records</label>
							</div>
						</div>

					</div>

				</div>
			</fieldset>
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

	function addColumnToSelected() {

		var rowID;

		$('#SelectedColumns').jqGrid('resetSelection');

		var selectedRows = $('#AvailableColumns').jqGrid('getGridParam', 'selarrrow');

		for (var i = 0; i <= selectedRows.length - 1; i++) {
			rowID = selectedRows[i];
			var datarow = getDatarowFromAvailable(selectedRows[i]);

			OpenHR.postData("Reports/AddReportColumn", datarow);

			$("#SelectedColumns").jqGrid('addRowData', datarow.ID, datarow);
			$('#SelectedColumns').jqGrid("setSelection", rowID);

			if (datarow.IsExpression == "false") {
				$("#SortOrdersAvailable").val(parseInt($("#SortOrdersAvailable").val()) + 1);
				button_disable($("#btnSortOrderAdd")[0], ($("#SortOrdersAvailable").val() == 0));
			}

		}

		if ('@Model.ReportType' == '@UtilityType.utlMailMerge') {
			$("#SelectedColumns").setGridParam({ sortname: 'Name', sortorder: 'asc' }).trigger('reloadGrid');
		}

		var ids = $("#AvailableColumns").getDataIDs();
		var nextIndex = $("#AvailableColumns").getInd(rowID);

		// Remove selected columns from available
		for (var i = selectedRows.length - 1; i >= 0; i--) {
			$("#AvailableColumns").delRowData(selectedRows[i]);
		}

		$("#AvailableColumns").jqGrid("setSelection", ids[nextIndex], true);
		refreshcolumnPropertiesPanel();

	}

	function getDatarowFromAvailable(index) {

		var datarow = $("#AvailableColumns").getRowData(index);

		datarow.ReportType = '@Model.ReportType';
		datarow.ReportID = '@Model.ID';
		datarow.Heading = datarow.Name.substr(0, 50);

		if (datarow.IsExpression == "false") {
			datarow.Name = $("#SelectedTableID option:selected").text() + '.' + datarow.Name;
		}

		datarow.Sequence = $("#SelectedColumns").jqGrid('getGridParam', 'records') + 1;
		datarow.IsAverage = false;
		datarow.IsCount = false;
		datarow.IsTotal = false;
		datarow.IsHidden = false;
		datarow.IsGroupWithNext = false;
		datarow.IsRepeated = false;
		datarow.TableID = $("#SelectedTableID option:selected").val();

		return datarow;
	}

	function addAllColumnsToSelected() {

		var sType;
		if ($('input[name=columnSelectiontype]:checked').val() == 0) {
			sType = "C";
		}
		else {
			sType = "E";
		}

		var allRows = $('#AvailableColumns').jqGrid('getDataIDs');
		var postData = {
			ReportID: '@Model.ID',
			ReportType: '@Model.ReportType',
			SelectionType: sType,
			ColumnsTableID: $("#SelectedTableID").val(),
			TableName: $("#SelectedTableID option:selected").text(),
			Columns: allRows
		};

		OpenHR.postData("Reports/AddAllReportColumns", postData);

		for (var i = 0; i <= allRows.length - 1; i++) {
			rowID = allRows[i];
			var datarow = getDatarowFromAvailable(allRows[i]);

			$("#SelectedColumns").jqGrid('addRowData', datarow.ID, datarow);

			if (datarow.IsExpression == "false") {
				$("#SortOrdersAvailable").val(parseInt($("#SortOrdersAvailable").val()) + 1);
				button_disable($("#btnSortOrderAdd")[0], ($("#SortOrdersAvailable").val() == 0));
			}

		}

		$('#SelectedColumns').jqGrid("setSelection", rowID);
		$('#AvailableColumns').jqGrid('clearGridData');

		refreshcolumnPropertiesPanel();

	}

	function requestRemoveAllSelectedColumns() {

		if ($("#SortOrders").jqGrid('getGridParam', 'records') > 0) {
			OpenHR.modalPrompt("Removing all the columns will also remove them from the report sort order." +
					"<br/><br/>Are you sure you wish to continue ?", 4, "").then(function (answer) {
						if (answer == 6) {
							removeAllSelectedColumns(true);
						}
					});
		}
		else {
			removeAllSelectedColumns(true);
		}
	}

	function requestRemoveSelectedColumns() {

		var selectedRows = $('#SelectedColumns').jqGrid('getGridParam', 'selarrrow');
		var sMessage = "";

		for (var i = 0; i <= selectedRows.length - 1; i++) {
			rowID = selectedRows[i];

			if ($("#SortOrders #" + rowID).length > 0) {
				var datarow = $("#SelectedColumns").getRowData(selectedRows[i]);
				sMessage += datarow.Name + "<br/>";
			}
		}

		if (sMessage.length > 0) {
			OpenHR.modalPrompt("Removing the following columns will also be removed from the sort order.<br/>" + sMessage + "" +
					"Are you sure you wish to continue ?", 4, "").then(function (answer) {
						if (answer == 6) {
							removeSelectedColumns();
							enableSaveButton();
						}
					});
		}
		else {
			removeSelectedColumns();
			enableSaveButton();
		}
	}

		function removeSelectedColumns() {

		var thisIndex;

		var selectedRows = $('#SelectedColumns').jqGrid('getGridParam', 'selarrrow');

		var postData = {
			ReportID: '@Model.ID',
			ReportType: '@Model.ReportType',
			ColumnsTableID: $("#SelectedTableID").val(),
			Columns: selectedRows
		};

		for (var i = 0; i <= selectedRows.length - 1; i++) {

			var datarow = $("#SelectedColumns").getRowData(selectedRows[i]);

			// Remove from sort order
			if ($("#SortOrders #" + datarow.ID).length > 0) {
				$("#SortOrders").delRowData(datarow.ID);
			}
			else
				if (datarow.IsExpression == "false") {
					$("#SortOrdersAvailable").val(parseInt($("#SortOrdersAvailable").val()) - 1);
				}

			thisIndex = $("#SelectedColumns").getInd(selectedRows[i]);
		}

		OpenHR.postData("Reports/RemoveReportColumn", postData, getAvailableTableColumnsCalcs);
		refreshSortButtons();

		// Position next selected column
		var recordCount = $("#SelectedColumns").jqGrid('getGridParam', 'records')
		var ids = $("#SelectedColumns").getDataIDs();
		if (thisIndex >= recordCount) { thisIndex = 0; }

		// Remove removed columns
		for (var i = selectedRows.length - 1; i >= 0; i--) {
			$("#SelectedColumns").delRowData(selectedRows[i]);
		}

		if (childColumnsCount() == 0) {
			resetRepeatOnChildRows();
		}

		$("#SelectedColumns").jqGrid("setSelection", ids[thisIndex], true);
		refreshcolumnPropertiesPanel();

	}

	function removeAllSelectedColumns(reloadColumns) {

		var dataSend = {
			ReportID: '@Model.ID',
			ReportType: '@Model.ReportType'
		};

		OpenHR.postData("Reports/RemoveAllReportColumns", dataSend);
		$('#SelectedColumns').jqGrid('clearGridData');

		if (reloadColumns == true) {
			getAvailableTableColumnsCalcs();
		}

		removeAllSortOrders();
		$("#SortOrdersAvailable").val(0);
		button_disable($("#btnSortOrderAdd")[0], true);

		if (childColumnsCount() == 0) {
			resetRepeatOnChildRows();
		}

		refreshcolumnPropertiesPanel();
	}


	function getAvailableTableColumnsCalcs() {
		var sType;
		var bIsBaseTable;

		$("#AvailableColumns").jqGrid('GridUnload');

		bIsBaseTable = ($("#SelectedTableID").val() == $("#BaseTableID").val());
		if (!bIsBaseTable && $("#txtReportType").val() == '@UtilityType.utlMailMerge') {
			$('#columnSelectiontype_0').prop("checked", true);
			$(".radioColumnType").attr("disabled", "disabled");
		}
		else {
			$(".radioColumnType").removeAttr("disabled");
		}

		if ($('input[name=columnSelectiontype]:checked').val() == 0) {
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
				{ name: 'Size', index: 'Size', hidden: true },
				{ name: 'Decimals', index: 'Decimals', hidden: true }],
			viewrecords: true,
			autowidth: false,
			sortname: 'Name',
			sortorder: "desc",
			rowNum: 10000,
			scrollrows: true,
			multiselect: true,
			beforeSelectRow: function (rowid, e) {
				var $this = $(this), rows = this.rows,
						// get id of the previous selected row
						startId = $this.jqGrid('getGridParam', 'selrow'),
						startRow, endRow, iStart, iEnd, i, rowidIndex;

				if (!e.ctrlKey && !e.shiftKey) {
					$this.jqGrid('resetSelection');
				} else if (startId && e.shiftKey) {
					$this.jqGrid('resetSelection');

					// get DOM elements of the previous selected and the currect selected rows
					startRow = rows.namedItem(startId);
					endRow = rows.namedItem(rowid);
					if (startRow && endRow) {
						// get min and max from the indexes of the previous selected
						// and the currect selected rows
						iStart = Math.min(startRow.rowIndex, endRow.rowIndex);
						rowidIndex = endRow.rowIndex;
						iEnd = Math.max(startRow.rowIndex, rowidIndex);
						for (i = iStart; i <= iEnd; i++) {
							// the row with rowid will be selected by jqGrid, so:
							if (i != rowidIndex) {
								$this.jqGrid('setSelection', rows[i].id, false);
							}
						}
					}

					// clear text selection
					if (document.selection && document.selection.empty) {
						document.selection.empty();
					} else if (window.getSelection) {
						window.getSelection().removeAllRanges();
					}
				}
				return true;
			},
			ondblClickRow: function (rowid) {
				if (!isDefinitionReadOnly()) {
					addColumnToSelected(rowid);
					enableSaveButton();
				}
			},
			loadComplete: function (data) {
				var topID = $("#AvailableColumns").getDataIDs()[0]
				$("#AvailableColumns").jqGrid("setSelection", topID);
				refreshcolumnPropertiesPanel();
			}
		});

		$("#AvailableColumns").jqGrid('hideCol', 'cb');

		//if ($('#txtReportType').val() == "utlCustomReport")
		resizeColumnGrids(); //should be in scope; this function resides in Util_Def_CustomReport.vbhtml

	}

	function changeColumnIsHidden() {

		if ($("#SelectedColumnIsHidden").is(':checked')) {
			$('#SelectedColumnIsAverage').prop('checked', false);
			$('#SelectedColumnIsCount').prop('checked', false);
			$('#SelectedColumnIsTotal').prop('checked', false);
			$('#SelectedColumnIsGroupWithNext').prop('checked', false);
		}

		refreshcolumnPropertiesPanel();
		updateColumnsSelectedGrid();
	}

	function changeColumnIsGroupWithNext() {

		if ($("#SelectedColumnIsGroupWithNext").is(':checked')) {
			$('#SelectedColumnIsAverage').prop('checked', false);
			$('#SelectedColumnIsCount').prop('checked', false);
			$('#SelectedColumnIsTotal').prop('checked', false);
			$('#SelectedColumnIsHidden').prop('checked', false);
		}

		refreshcolumnPropertiesPanel();
		updateColumnsSelectedGrid();
	}

	function resetRepeatOnChildRows() {

		var allRows = $('#SelectedColumns').jqGrid('getDataIDs');
		for (var i = 0; i <= allRows.length - 1; i++) {
			var datarow = $("#SelectedColumns").getRowData(allRows[i]);
			datarow.IsRepeated = false;
			$('#SelectedColumns').jqGrid('setRowData', allRows[i], datarow);
		}

		return true;
	}


	function childColumnsCount() {

		var allRows = $('#SelectedColumns').jqGrid('getDataIDs');
		var iChildCount = 0;

		for (var i = 0; i <= allRows.length - 1; i++) {
			var datarow = $("#SelectedColumns").getRowData(allRows[i]);
			if (datarow.TableID != $('#BaseTableID').val()) {
				iChildCount += 1;
			}
		}

		return iChildCount;
	}

	function refreshcolumnPropertiesPanel() {

		var rowCount = $('#SelectedColumns').jqGrid('getGridParam', 'selarrrow').length;
		var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');
		var dataRow = $("#SelectedColumns").getRowData(rowId)
		var allRows = $('#SelectedColumns').jqGrid('getDataIDs');
		var bDisableAdd = ($("#AvailableColumns").getGridParam("reccount") == 0);
		var isTopRow = true;
		var isBottomRow = true;
		var isReadOnly = isDefinitionReadOnly();
		var bRowSelected = false;

		if (allRows.length > 0) {
			bRowSelected = true;
			isTopRow = (rowId == allRows[0]);
			isBottomRow = (rowId == allRows[allRows.length - 1]);
		}

		if (rowCount > 1 || allRows.length == 0) {
			$("#definitionColumnProperties :input").attr("disabled", true);
			$("#SelectedColumnHeading").val("");
			$("#SelectedColumnSize").val("");
			$("#SelectedColumnDecimals").val("");
			$('#SelectedColumnIsAverage').prop('checked', false);
			$('#SelectedColumnIsCount').prop('checked', false);
			$('#SelectedColumnIsTotal').prop('checked', false);
			$('#SelectedColumnIsHidden').prop('checked', false);
			$('#SelectedColumnIsGroupWithNext').prop('checked', false);
			$('#SelectedColumnIsRepeated').prop('checked', false);

			$(".numericOnly").css("color", "#A59393");
			$(".cannotBeHidden").css("color", "#A59393");
			$(".canGroupWithNext").css("color", "#A59393");
			$("#labelSelectedColumnIsHidden").css("color", "#A59393");
			$(".baseTableOnly").css("color", "#A59393");
		}
		else {

			if (!isReadOnly) {
				$("#definitionColumnProperties :input").removeAttr("disabled");
			}

			var isThereChildColumns = (childColumnsCount() > 0);
			var isNumeric = (dataRow.DataType == '2' || dataRow.DataType == '4');
			var isDecimals = (isNumeric == true || dataRow.IsExpression == "true");
			var isBaseOrParentTableColumn = (dataRow.TableID == $("#BaseTableID").val()) || (dataRow.TableID == $("#txtParent1ID").val()) || (dataRow.TableID == $("#txtParent2ID").val());

			var isHidden = $("#SelectedColumnIsHidden").is(':checked');
			var isGroupWithNext = $("#SelectedColumnIsGroupWithNext").is(':checked');

			$(".numericOnly :input").attr("disabled", !isNumeric || isHidden || isGroupWithNext || isReadOnly);
			$(".cannotBeHidden :input").attr("disabled", isHidden || isGroupWithNext || isReadOnly);
			$(".decimalsOnly :input").attr("disabled", !isDecimals || isReadOnly);
			$(".baseTableOnly :input").attr("disabled", !isBaseOrParentTableColumn || !isThereChildColumns || isReadOnly);
			$(".canGroupWithNext :input").attr("disabled", isBottomRow || isHidden || isReadOnly);
			$("#SelectedColumnIsHidden").attr("disabled", isGroupWithNext || isReadOnly);

			if (!isNumeric || isHidden || isGroupWithNext || isReadOnly) {
				$(".numericOnly").css("color", "#A59393");
			} else {
				$(".numericOnly").css("color", "#000000");
			}

			if (isHidden || isGroupWithNext || isReadOnly) {
				$(".cannotBeHidden").css("color", "#A59393");
			} else {
				$(".cannotBeHidden").css("color", "#000000");
			}

			if (isBottomRow || isHidden || isReadOnly) {
				$(".canGroupWithNext").css("color", "#A59393");
			} else {
				$(".canGroupWithNext").css("color", "#000000");
			}

			if (isGroupWithNext || isReadOnly) {
				$("#labelSelectedColumnIsHidden").css("color", "#A59393");
			} else {
				$("#labelSelectedColumnIsHidden").css("color", "#000000");
			}

			if (isBaseOrParentTableColumn && isThereChildColumns && !isReadOnly) {
				$(".baseTableOnly").css("color", "#000000");
			} else {
				$(".baseTableOnly").css("color", "#A59393");
			}

		}

		// Enable / Disable relevant buttons
		button_disable($("#btnColumnAdd")[0], bDisableAdd || isReadOnly);
		button_disable($("#btnColumnAddAll")[0], bDisableAdd || isReadOnly);
		button_disable($("#btnColumnRemove")[0], !bRowSelected || isReadOnly);
		button_disable($("#btnColumnRemoveAll")[0], !bRowSelected || isReadOnly);
		button_disable($("#btnColumnMoveUp")[0], isTopRow || isReadOnly || (rowCount > 1));
		button_disable($("#btnColumnMoveDown")[0], isBottomRow || isReadOnly || (rowCount > 1));

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
				id: "ID" //index of the column with the PK in it
			},
			colNames: ['ID', 'TableID', 'IsExpression', 'Name', 'Sequence', 'Heading', 'DataType',
								'Size', 'Decimals', 'IsAverage', 'IsCount', 'IsTotal', 'IsHidden', 'IsGroupWithNext', 'IsRepeated', 'ReportID', 'ReportType'],
			colModel: [
				{ name: 'ID', index: 'ID', hidden: true },
				{ name: 'TableID', index: 'TableID', hidden: true },
				{ name: 'IsExpression', index: 'IsExpression2', hidden: true },
				{ name: 'Name', index: 'Name', sortable: false },
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
			autowidth: false,
			sortname: 'Sequence',
			sortorder: "asc",
			rowNum: 10000,
			scrollrows: true,
			multiselect: true,
			beforeSelectRow: function (rowid, e) {

				if ($('#SelectedColumns').jqGrid('getGridParam', 'selarrrow').length == 1) {
					updateColumnsSelectedGrid();
				}

				var $this = $(this), rows = this.rows,
						// get id of the previous selected row
						startId = $this.jqGrid('getGridParam', 'selrow'),
						startRow, endRow, iStart, iEnd, i, rowidIndex;

				if (!e.ctrlKey && !e.shiftKey) {
					$this.jqGrid('resetSelection');
				} else if (startId && e.shiftKey) {
					$this.jqGrid('resetSelection');

					// get DOM elements of the previous selected and the currect selected rows
					startRow = rows.namedItem(startId);
					endRow = rows.namedItem(rowid);
					if (startRow && endRow) {
						// get min and max from the indexes of the previous selected
						// and the currect selected rows
						iStart = Math.min(startRow.rowIndex, endRow.rowIndex);
						rowidIndex = endRow.rowIndex;
						iEnd = Math.max(startRow.rowIndex, rowidIndex);
						for (i = iStart; i <= iEnd; i++) {
							// the row with rowid will be selected by jqGrid, so:
							if (i != rowidIndex) {
								$this.jqGrid('setSelection', rows[i].id, false);
							}
						}
					}

					// clear text selection
					if (document.selection && document.selection.empty) {
						document.selection.empty();
					} else if (window.getSelection) {
						window.getSelection().removeAllRanges();
					}
				}
				return true;
			},
			onSelectRow: function (id) {

				var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');
				var dataRow = $("#SelectedColumns").getRowData(rowId)

				$("#SelectedColumnHeading").val(dataRow.Heading);

				$("#SelectedColumnSize").val(dataRow.Size);
				$("#SelectedColumnDecimals").val(dataRow.Decimals);
				$('#SelectedColumnIsAverage').prop('checked', JSON.parse(dataRow.IsAverage));
				$('#SelectedColumnIsCount').prop('checked', JSON.parse(dataRow.IsCount));
				$('#SelectedColumnIsTotal').prop('checked', JSON.parse(dataRow.IsTotal));
				$('#SelectedColumnIsHidden').prop('checked', JSON.parse(dataRow.IsHidden));
				$('#SelectedColumnIsGroupWithNext').prop('checked', JSON.parse(dataRow.IsGroupWithNext));
				$('#SelectedColumnIsRepeated').prop('checked', JSON.parse(dataRow.IsRepeated));

				refreshcolumnPropertiesPanel();

			},
			loadComplete: function (data) {
				var topID = $("#SelectedColumns").getDataIDs()[0]
				$("#SelectedColumns").jqGrid("setSelection", topID);
			}
		});

		if ('@Model.ReportType' == '@UtilityType.utlMailMerge') {
			$("#SelectedColumns").setGridParam({ sortname: 'Name', sortorder: 'asc' }).trigger('reloadGrid');
		}

		if ('@Model.ReportType' == '@UtilityType.utlCustomReport') {
			$("#SelectedColumns").jqGrid('sortableRows');
		}

		$("#SelectedColumns").jqGrid('hideCol', 'cb');

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