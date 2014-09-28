﻿@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.ReportBaseModel)

@Html.HiddenFor(Function(m) m.SortOrdersString, New With {.id = "txtSOAAS"})
@Html.HiddenFor(Function(m) m.SortOrdersAvailable, New With {.id = "SortOrdersAvailable"})

<fieldset style="width:100%">
	<legend class="fontsmalltitle">Sort Order :</legend>

	<div id="divSortOrderDiv" style="float:left" class="width80">
		<table id="SortOrders"></table>
	</div>

	<div class="stretchyfixedbuttoncolumn " id="sortorderbuttons" style="float:left;padding-left:20px">
		<div id="colbtngrp1">
			<input type="button" id="btnSortOrderAdd" value="Add" disabled onclick="addSortOrder();" />
			<input type="button" id="btnSortOrderEdit" value="Edit" disabled onclick="editSortSorder(0);" />
		</div>
		<div id="colbtngrp2">
			<input type="button" id="btnSortOrderRemove" value="Remove" disabled onclick="removeSortOrder()" />
			<input type="button" id="btnSortOrderRemoveAll" value="Remove All" disabled onclick="removeAllSortOrders()" />
		</div>
		<div id="colbtngrp3">
			<input type="button" id="btnSortOrderMoveUp" value="Move Up" disabled onclick="moveSelectedOrder('up')" />
			<input type="button" id="btnSortOrderMoveDown" value="Move Down" disabled onclick="moveSelectedOrder('down')" />
		</div>
	</div>
</fieldset>


	<script type="text/javascript">

		function refreshSortButtons() {

			var isReadonly = isDefinitionReadOnly();
			var bDisableRemoveAll = ($("#SortOrders").getGridParam("reccount") == 0) || ($("#SelectedColumns").getGridParam("reccount") == 0);
			var isTopRow = false;
			var isBottomRow = false;
			var isRowSelected = ($("#SortOrders").jqGrid('getGridParam', 'selrow') > 0);

			if (isRowSelected) {
				var rowId = $("#SortOrders").jqGrid('getGridParam', 'selrow');
				var allRows = $("#SortOrders")[0].rows;
				isTopRow = (rowId == allRows[1].id);
				isBottomRow = (rowId == allRows[allRows.length - 1].id);			
			} 

			button_disable($("#btnSortOrderAdd")[0], ($("#SortOrdersAvailable").val() == 0) || isReadonly);
			button_disable($("#btnSortOrderEdit")[0], !isRowSelected || isReadonly);
			button_disable($("#btnSortOrderRemove")[0], !isRowSelected || isReadonly);
			button_disable($("#btnSortOrderRemoveAll")[0], bDisableRemoveAll || isReadonly);
			button_disable($("#btnSortOrderMoveUp")[0], !isRowSelected || isTopRow || isReadonly);
			button_disable($("#btnSortOrderMoveDown")[0], !isRowSelected || isBottomRow || isReadonly);

		}

		$(function () {

			attachGrid();
			$("#SortOrders").jqGrid('setGridWidth', $("#divSortOrderDiv").width() * .95);

		})

		function removeSortOrder() {

			var recordCount = $("#SortOrders").jqGrid('getGridParam', 'records')
			var ids = $("#SortOrders").getDataIDs();
			var rowID = $("#SortOrders").jqGrid('getGridParam', 'selrow');
			var datarow = $("#SortOrders").getRowData(rowID);
			var thisIndex = $("#SortOrders").getInd(rowID);

			OpenHR.postData('Reports/RemoveSortOrder', datarow);
			$("#SortOrders").jqGrid('delRowData', rowID);

			$("#SortOrdersAvailable").val(parseInt($("#SortOrdersAvailable").val()) + 1);

			if (thisIndex >= recordCount) { thisIndex = 0; }
			$("#SortOrders").jqGrid("setSelection", ids[thisIndex], true);

			refreshSortButtons();

		}

		function removeAllSortOrders() {

			var rows = $("#SortOrders").jqGrid('getDataIDs');
			var nowAvailable = $("#SortOrders").getGridParam("reccount") + parseInt($("#SortOrdersAvailable").val());

			for (var i = 0; i < rows.length; i++) {
				var datarow = $("#SortOrders").getRowData(rows[i]);
				OpenHR.postData('Reports/RemoveSortOrder', datarow);
			}

			$("#SortOrders").jqGrid('clearGridData');

			$("#SortOrdersAvailable").val(nowAvailable);
			refreshSortButtons();

		}


	function attachGrid() {

		$("#SortOrders").jqGrid({

			datatype: 'jsonstring',
			datastr: '@Model.SortOrders.ToJsonResult',
			mtype: 'GET',
			jsonReader: {
				root: "rows", //array containing actual data
				page: "page", //current page
				total: "total", //total pages for the query
				records: "records", //total number of records
				repeatitems: false,
				id: "ID"
			},
			cmTemplate: {sortable:false},
			colNames: ['ID', 'ReportID', 'ReportType', 'TableID', 'ColumnID', 'Column Name', 'Sequence', 'Order',
									'Break On Change', 'Page On Change', 'Value On Change', 'Suppress Repeated'],
			colModel: [
									{ name: 'ID', width: 50, key: true, hidden: true },
									{ name: 'ReportID', width: 50, hidden: true },
									{ name: 'ReportType', width: 50, hidden: true },
									{ name: 'TableID', width: 50, hidden: true },
									{ name: 'ColumnID', width: 50, hidden: true },
									{ name: 'Name', index: 'Name', width: 200 },
									{ name: 'Sequence', index: 'Sequence', width: 150, hidden: true },
									{ name: 'Order', index: 'Order', width: 90, editable: true, formatter: "select", edittype: "select", editoptions: { value: "0:Ascending;1:Descending" } },
									{
										name: 'BreakOnChange', index: 'BreakOnChange', width: 120, align: "center", hidden: true,
										editable: true, edittype: 'checkbox', editoptions: { value: "True:False" }, formatter: "checkbox", formatoptions: { disabled: false }
									},
									{
										name: 'PageOnChange', index: 'PageOnChange', width: 120, align: "center", hidden: true,
										editable: true, edittype: 'checkbox', editoptions: { value: "True:False" }, formatter: "checkbox", formatoptions: { disabled: false }
									},
									{
										name: 'ValueOnChange', index: 'ValueOnChange', width: 120, align: "center", hidden: true,
										editable: true, edittype: 'checkbox', editoptions: { value: "True:False" }, formatter: "checkbox", formatoptions: { disabled: false }
									},
									{
										name: 'SuppressRepeated', index: 'SuppressRepeated', width: 120, align: "center", hidden: true,
										editable: true, edittype: 'checkbox', editoptions: { value: "True:False" }, formatter: "checkbox", formatoptions: { disabled: false }
									}
			],
			viewrecords: true,
			width: 'auto',
			height: '400px',
			scrollrows: true,
			ondblClickRow: function (rowID) {
				editSortSorder(rowID);
				enableSaveButton();
			},
			onSelectRow: function (id) {
				refreshSortButtons();
			},
			loadComplete: function (data) {

				if ('@Model.ReportType' == '@UtilityType.utlCustomReport') {
					$(this).showCol("BreakOnChange");
					$(this).showCol("PageOnChange");
					$(this).showCol("ValueOnChange");
					$(this).showCol("SuppressRepeated");
				}

				$('#SortOrders input[type=checkbox]').on('click', function () { enableSaveButton(); });

				var topID = $("#SortOrders").getDataIDs()[0]
				$("#SortOrders").jqGrid("setSelection", topID);
			}
		});
	}

	function addSortOrder() {
		OpenHR.OpenDialog("Reports/AddSortOrder", "divPopupReportDefinition", { ReportID: "@Model.ID", ReportType: "@Model.ReportType" }, 'auto');
	}

	function editSortSorder(rowID) {

		if (rowID == 0) {
			rowID = $('#SortOrders').jqGrid('getGridParam', 'selrow');
		}

		var gridData = $("#SortOrders").getRowData(rowID);
		OpenHR.OpenDialog("Reports/EditSortOrder", "divPopupReportDefinition", gridData, 'auto');

	}

	function moveSelectedOrder(direction) {
		OpenHR.MoveItemInGrid($("#SortOrders"), direction);
	}

	</script>
