@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.CalendarReportModel)

@Html.HiddenFor(Function(m) m.EventsString, New With {.id = "txtCEAAS"})

<div id="eventDetailsContainer">
	<fieldset>
		<legend class="fontsmalltitle">Calendar Events :</legend>

		<div id="divEventDetails" class="stretchyfill">
			<table id="CalendarEvents"></table>
		</div>

		<div class="stretchyfixed">
			<input type="button" id="btnEventDetailsAdd" value="Add..." onclick="eventAdd();" />
			<br />
			<input type="button" id="btnEventDetailsEdit" value="Edit..." disabled onclick="eventEdit();" />
			<br />
			<input type="button" id="btnEventDetailsRemove" class="enableSaveButtonOnClick" value="Remove" disabled onclick="removeEvent()" />
			<br />
			<input type="button" id="btnEventDetailsRemoveAll" class="enableSaveButtonOnClick" value="Remove All" disabled onclick="removeAllEvents()" />
		</div>
	</fieldset>
</div>

<script type="text/javascript">

	$(function () {
		attachCalendarEventsGrid();
	});

	function attachCalendarEventsGrid() {
		//create the column layout:
		var gridWidth = $('#divEventDetails').width();

		$("#CalendarEvents").jqGrid({
			datatype: "jsonstring",
			datastr: '@Model.Events.ToJsonResult',
			mtype: 'GET',
			jsonReader: {
				root: "rows", //array containing actual data
				page: "page", //current page
				total: "total", //total pages for the query
				records: "records", //total number of records
				repeatitems: false,
				id: "EventKey" //index of the column with the PK in it
			},
			colNames: ['ID', 'EventKey', 'ReportID', 'ReportType',
					'Name', 'TableID', 'Table', 'FilterID', 'EventEndType', 'EventStartDateID', 'EventStartSessionID', 'EventEndDateID', 'EventEndSessionID', 'EventDurationID',
					'LegendType', 'LegendCharacter', 'LegendLookupTableID', 'LegendLookupColumnID', 'LegendLookupCodeID', 'LegendEventColumnID',
					'EventDesc1ColumnID', 'EventDesc2ColumnID', 'Filter', 'FilterViewAccess',
					'Start Date', 'Start Session', 'End Date', 'End Session', 'Duration', 'Key',
					'Description 1', 'Description 2'
			],
			colModel: [
				{ name: 'ID', index: 'ID', sorttype: 'int', hidden: true },
				{ name: 'EventKey', index: 'EventKey', sorttype: 'int', hidden: true },
				{ name: 'ReportID', index: 'ReportID', sorttype: 'int', hidden: true },
				{ name: 'ReportType', index: 'ReportType', hidden: true },
				{ name: 'Name', index: 'Name', sorttype: 'text' },
				{ name: 'TableID', index: 'TableID', width: 100, hidden: true },
				{ name: 'TableName', index: 'TableName', width: 100 },
				{ name: 'FilterID', index: 'FilterID', width: 100, hidden: true },
				{ name: 'EventEndType', index: 'EventEndType', sorttype: 'int', hidden: true },
				{ name: 'EventStartDateID', index: 'EventStartDateID', sorttype: 'int', hidden: true },
				{ name: 'EventStartSessionID', index: 'EventStartSessionID', sorttype: 'int', hidden: true },
				{ name: 'EventEndDateID', index: 'EventEndDateID', sorttype: 'int', hidden: true },
				{ name: 'EventEndSessionID', index: 'EventEndSessionID', sorttype: 'int', hidden: true },
				{ name: 'EventDurationID', index: 'EventDurationID', sorttype: 'int', hidden: true },
				{ name: 'LegendType', index: 'LegendType', sorttype: 'text', hidden: true },
				{ name: 'LegendCharacter', index: 'LegendCharacter', sorttype: 'text', hidden: true },
				{ name: 'LegendLookupTableID', index: 'LegendLookupTableID', sorttype: 'int', hidden: true },
				{ name: 'LegendLookupColumnID', index: 'LegendLookupColumnID', sorttype: 'int', hidden: true },
				{ name: 'LegendLookupCodeID', index: 'LegendLookupCodeID', sorttype: 'int', hidden: true },
				{ name: 'LegendEventColumnID', index: 'LegendEventColumnID', sorttype: 'int', hidden: true },
				{ name: 'EventDesc1ColumnID', index: 'EventDesc1ColumnID', sorttype: 'int', hidden: true },
				{ name: 'EventDesc2ColumnID', index: 'EventDesc2ColumnID', sorttype: 'int', hidden: true },
				{ name: 'FilterName', index: 'FilterName', sorttype: 'text' },
				{ name: 'FilterViewAccess', index: 'FilterViewAccess', hidden: true, classes: 'ViewAccess' },
				{ name: 'EventStartDateName', index: 'EventStartDateName', sorttype: 'text' },
				{ name: 'EventStartSessionName', index: 'EventStartSessionName', sorttype: 'text' },
				{ name: 'EventEndDateName', index: 'EventEndDateName', sorttype: 'text' },
				{ name: 'EventEndSessionName', index: 'EventEndSessionName', sorttype: 'text' },
				{ name: 'EventDurationName', index: 'EventDurationName', sorttype: 'text' },
				{ name: 'LegendTypeName', index: 'LegendTypeName', sorttype: 'text' },
				{ name: 'EventDesc1ColumnName', index: 'EventDesc1ColumnName', sorttype: 'text' },
				{ name: 'EventDesc2ColumnName', index: 'EventDesc2ColumnName', sorttype: 'text' }
			],
			rowNum: 10,
			rowTotal: 50,
			rowList: [10, 20, 30],
			pager: '#pcrud',
			autowidth: false,
			sortname: 'Name',
			loadonce: true,
			viewrecords: true,
			sortorder: "desc",
			width: gridWidth,
			ondblClickRow: function (rowID) {
				eventEdit(rowID);
			},
			onSelectRow: function (id) {

				var isReadOnly = isDefinitionReadOnly();
				button_disable($("#btnEventDetailsEdit")[0], isReadOnly);
				button_disable($("#btnEventDetailsRemove")[0], isReadOnly);
				button_disable($("#btnEventDetailsRemoveAll")[0], isReadOnly);

			},
			gridComplete: function () {

				button_disable($("#btnEventDetailsEdit")[0], true);
				button_disable($("#btnEventDetailsRemove")[0], true);
				button_disable($("#btnEventDetailsRemoveAll")[0], true);

				// Highlight top row
				var ids = $(this).jqGrid("getDataIDs");
				if (ids && ids.length > 0)
					$(this).jqGrid("setSelection", ids[0]);
			}
		});

		$("#Events").jqGrid('navGrid', '#pcrud', {});
	}


	function eventAdd() {
		OpenHR.OpenDialog("Reports/AddCalendarEvent", "divPopupReportDefinition", { ReportID: "@Model.ID" }, '1000');
	}

	function eventEdit() {

		var rowID = $('#CalendarEvents').jqGrid('getGridParam', 'selrow');
		var datarow = $("#CalendarEvents").getRowData(rowID);

		OpenHR.OpenDialog("Reports/EditCalendarEvent", "divPopupReportDefinition", datarow, '1000');

	}

	function removeEvent() {

		var recordCount = $("#CalendarEvents").jqGrid('getGridParam', 'records')
		var ids = $("#CalendarEvents").getDataIDs();
		var rowID = $('#CalendarEvents').jqGrid('getGridParam', 'selrow');
		var datarow = $("#CalendarEvents").getRowData(rowID);
		var thisIndex = $("#CalendarEvents").getInd(rowID);

		OpenHR.postData("Reports/RemoveCalendarEvent", datarow)
		$('#CalendarEvents').jqGrid('delRowData', rowID)

		if (thisIndex >= recordCount) { thisIndex = 0; }
		$("#CalendarEvents").jqGrid("setSelection", ids[thisIndex], true);

		checkIfDefinitionNeedsToBeHidden(0);
	}

	function removeAllEvents() {

		var i;

		var data = {};
		var grid = $('#CalendarEvents');
		var rows = grid.jqGrid('getDataIDs');

		for (i = 0; i < rows.length; i++) {
			var datarow = grid.jqGrid('getRowData', rows[i]);
			OpenHR.postData("Reports/RemoveCalendarEvent", datarow)
		}

		$('#CalendarEvents').jqGrid('clearGridData');
		checkIfDefinitionNeedsToBeHidden(0);
	}


</script>
