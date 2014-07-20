﻿@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.ReportBaseModel)

@Html.HiddenFor(Function(m) m.SortOrdersString, New With {.id = "txtSOAAS"})

<fieldset class="relatedtables">
	<legend>Sort Order :</legend>

	<div class="stretchyfill">
		<table id="SortOrders"></table>
	</div>

	<div class="stretchyfixed">
		<input type="button" id="btnSortOrderAdd" value="Add" disabled onclick="addSortOrder();" />
		<input type="button" id="btnSortOrderEdit" value="Edit" disabled onclick="editSortSorder(0);" />
		<input type="button" id="btnSortOrderRemove" value="Remove" disabled onclick="OpenHR.RemoveRowFromGrid(SortOrders, 'Reports/RemoveSortOrder')" />
		<input type="button" id="btnSortOrderRemoveAll" value="Remove All" disabled onclick="OpenHR.RemoveAllRowsFromGrid(SortOrders, 'Reports/RemoveSortOrder')" />
		<input type="button" id="btnSortOrderMoveUp" value="Move Up" disabled onclick="moveSortOrderUp()" />
		<input type="button" id="btnSortOrderMoveDown" value="Move Down" disabled onclick="moveSortOrderDown()" />
	</div>

</fieldset>

	<fieldset>
		@code
			If Model.ReportType = UtilityType.utlCustomReport Then
				@Html.SortOrderGrid("Repetition", Model.Repetition, Nothing)
			End If
		End Code
	</fieldset>

	<script type="text/javascript">

		$(function () {
			attachGrid();
		})

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
				colNames: ['ID','ReportID', 'ReportType', 'TableID', 'ColumnID', 'Name', 'Sequence', 'Order',
										'BreakOnChange', 'PageOnChange', 'ValueOnChange', 'SuppressRepeated'],
				colModel: [
										{ name: 'ID', width: 50, key: true, hidden: false },
										{ name: 'ReportID', width: 50, hidden: true },
										{ name: 'ReportType', width: 50, hidden: true },
										{ name: 'TableID', width: 50, hidden: true },
										{ name: 'ColumnID', width: 50, key: true, hidden: false },
										{ name: 'Name', index: 'Name', width: 600 },
										{ name: 'Sequence', index: 'Sequence', width: 150 },
										{ name: 'Order', index: 'Order', width: 150 },
										{ name: 'BreakOnChange', index: 'BreakOnChange', width: 150 },
										{ name: 'PageOnChange', index: 'PageOnChange', width: 150 },
										{ name: 'ValueOnChange', index: 'ValueOnChange', width: 150 },
										{ name: 'SuppressRepeated', index: 'SuppressRepeated', width: 120, align: "center" }
				],
				viewrecords: true,
				width: 400,
				sortname: 'Sequence',
				sortorder: "desc",
				ondblClickRow: function (rowID) {
					editSortSorder(rowID);
				},
				onSelectRow: function (id) {

					var rowId = $(this).jqGrid('getGridParam', 'selrow');
					var allRows = $(this)[0].rows;

					var isTopRow = (rowId == allRows[1].id);
					var isBottomRow = (rowId == allRows[allRows.length - 1].id);

					// Enable / Disable relevant buttons
					button_disable($("#btnSortOrderAdd")[0], false);
					button_disable($("#btnSortOrderEdit")[0], false);
					button_disable($("#btnSortOrderRemove")[0], false);
					button_disable($("#btnSortOrderRemoveAll")[0], false);
					button_disable($("#btnSortOrderMoveUp")[0], isTopRow);
					button_disable($("#btnSortOrderMoveDown")[0], isBottomRow);

				},
				gridComplete: function () {
					// Highlight top row
					var ids = $(this).jqGrid("getDataIDs");
					if (ids && ids.length > 0)
						$(this).jqGrid("setSelection", ids[0]);
				}
			});
		}


		function addSortOrder() {
			OpenHR.OpenDialog("Reports/AddSortOrder", "divPopupReportDefinition", { ReportID: "@Model.ID", ReportType: "@Model.ReportType"});
		}

		function editSortSorder(rowID) {

			if (rowID == 0) {
				rowID = $('#SortOrders').jqGrid('getGridParam', 'selrow');
			}

			var gridData = $("#SortOrders").getRowData(rowID);
			OpenHR.OpenDialog("Reports/EditSortOrder", "divPopupReportDefinition", gridData);

		}

	function moveSortOrderUp() {}

	function moveSortOrderDown() {}

	</script>
