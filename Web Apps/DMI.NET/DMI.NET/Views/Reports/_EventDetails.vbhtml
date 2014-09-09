@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.CalendarReportModel)

@Html.HiddenFor(Function(m) m.EventsString, New With {.id = "txtCEAAS"})

<fieldset class="relatedtables">
	<legend class="fontsmalltitle">Calendar Events :</legend>

	<div class="width80 floatleft overflowauto">
		<input type="hidden" id="CalendarEventsViewAccess" />
		<table id="CalendarEvents"></table>
	</div>

	<div class="stretchyfixed floatleft">
		<input type="button" id="btnEventDetailsAdd" value="Add..." onclick="eventAdd();" />
		<br />
		<input type="button" id="btnEventDetailsEdit" value="Edit..." disabled onclick="eventEdit(0);" />
		<br />
		<input type="button" id="btnEventDetailsRemove" value="Remove" disabled onclick="removeEvent()" />
		<br />
		<input type="button" id="btnEventDetailsRemoveAll" value="Remove All" disabled onclick="removeAllEvents()" />
	</div>
</fieldset>


<script type="text/javascript">

	$(function () {
		attachCalendarEventsGrid();
	})

	function attachCalendarEventsGrid() {
		//create the column layout:
			
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
				{ name: 'EventKey', index: 'EventKey', sorttype: 'text', hidden: true },
				{ name: 'ReportID', index: 'ReportID', sorttype: 'int', hidden: true },
				{ name: 'ReportType', index: 'ReportType', hidden: true },
				{ name: 'Name', index: 'Name', sorttype: 'text'},
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
				{ name: 'FilterViewAccess', index: 'FilterViewAccess', hidden: true },
				{ name: 'EventStartDateName', index: 'EventStartDateName', sorttype: 'text' },
				{ name: 'EventStartSessionName', index: 'EventStartSessionName', sorttype: 'text' },
				{ name: 'EventEndDateName', index: 'EventEndDateName', sorttype: 'text'},
				{ name: 'EventEndSessionName', index: 'EventEndSessionName', sorttype: 'text'},
				{ name: 'EventDurationName', index: 'EventDurationName', sorttype: 'text'},
				{ name: 'LegendTypeName', index: 'LegendTypeName', sorttype: 'text'},
				{ name: 'EventDesc1ColumnName', index: 'EventDesc1ColumnName', sorttype: 'text' },
				{ name: 'EventDesc2ColumnName', index: 'EventDesc2ColumnName', sorttype: 'text' }
			],
			rowNum: 10,
			rowTotal: 50,
			rowList: [10, 20, 30],
			autowidth: false,			
			pager: '#pcrud',
			sortname: 'Name',
			loadonce: true,
			viewrecords: true,
			sortorder: "desc",
			height: 400,
			ondblClickRow: function (rowID) {
				eventEdit(rowID);
				enableSaveButton();
			},
			onSelectRow: function (id) {

				var isReadOnly = isDefinitionReadOnly();
				button_disable($("#btnEventDetailsEdit")[0], isReadOnly);
				button_disable($("#btnEventDetailsRemove")[0], isReadOnly);
				button_disable($("#btnEventDetailsRemoveAll")[0], isReadOnly);

			},
			gridComplete: function () {
				// Highlight top row
				var ids = $(this).jqGrid("getDataIDs");
				if (ids && ids.length > 0)
					$(this).jqGrid("setSelection", ids[0]);
			}
		});
		$("#Events").jqGrid('navGrid', '#pcrud', {});

	}

	function eventAdd() {
		OpenHR.OpenDialog("Reports/AddCalendarEvent", "divPopupReportDefinition", { ReportID: "@Model.ID" }, 'auto');
	}


	function eventEdit(rowID) {

		if (rowID == 0) {
			rowID = $('#CalendarEvents').jqGrid('getGridParam', 'selrow');
		}

		var gridData = $("#CalendarEvents").getRowData(rowID);
		OpenHR.OpenDialog("Reports/EditCalendarEvent", "divPopupReportDefinition", gridData, 'auto');

	}

	function removeEvent() {

		var rowID = $('#CalendarEvents').jqGrid('getGridParam', 'selrow');
		var datarow = $("#CalendarEvents").getRowData(rowID);
		OpenHR.postData("Reports/RemoveCalendarEvent", datarow)
		$('#CalendarEvents').jqGrid('delRowData', rowID)

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
	}


</script>