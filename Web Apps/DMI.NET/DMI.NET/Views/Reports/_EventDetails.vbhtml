@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.CalendarReportModel)

<div class="left">

@*	@Html.TableFor("Events", Model.Events, Nothing)*@

	<div class="stretchyfill">
		<table id="CalendarEvents"></table>
	</div>

</div>



<div class="right">
	<input type="button" id="btnEventDetailsAdd" value="Add..." onclick="eventAdd();" />
	<br />
	<input type="button" id="btnEventDetailsEdit" value="Edit..." onclick="eventEdit(0);" />
	<br />
	<input type="button" id="btnEventDetailsRemove" value="Remove" />
	<br />
	<input type="button" id="btnEventDetailsRemoveAll" value="Remove All" />
</div>

<script type="text/javascript">

	$(function () {
		attachCalendarEventsGrid();
		attachSortOrderColumns();
	})

	function attachCalendarEventsGrid() {

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
				id: "ID" //index of the column with the PK in it
			},
			colNames: ['ID', 'EventKey', 'CalendarReportID',
					'Name', 'TableID', 'FilterID', 'EventStartDateID', 'EventStartSessionID', 'EventEndDateID', 'EventEndSessionID', 'EventDurationID',
					'LegendType', 'LegendCharacter', 'LegendLookupTableID', 'LegendLookupColumnID', 'LegendLookupCodeID', 'LegendEventColumnID',
					'EventDesc1ColumnID', 'EventDesc2ColumnID', 'FilterHidden', 'FilterName',
					'EventStartSessionName', 'EventEndDateName', 'EventEndSessionName', 'EventDurationName', 'LegendTypeName',
					'EventDesc1ColumnName', 'EventDesc2ColumnName'
			],
			colModel: [
				{ name: 'ID', index: 'ID', sorttype: 'int', hidden: true },
				{ name: 'EventKey', index: 'EventKey', sorttype: 'text', hidden: true },
				{ name: 'CalendarReportID', index: 'CalendarReportID', sorttype: 'int', hidden: true },
				{ name: 'Name', index: 'Name', sorttype: 'text', hidden: false },
				{ name: 'TableID', index: 'TableID', width: 100, hidden: true },
				{ name: 'FilterID', index: 'FilterID', width: 100, hidden: true },
				{ name: 'EventStartDateID', index: 'EventStartDateID', sorttype: 'int', hidden: true },
				{ name: 'EventStartSessionID', index: 'EventStartSessionID', sorttype: 'int', hidden: true },
				{ name: 'EventEndDateID', index: 'EventEndDateID', sorttype: 'int', hidden: true },
				{ name: 'EventEndSessionID', index: 'EventEndSessionID', sorttype: 'int', hidden: true },
				{ name: 'EventDurationID', index: 'EventDurationID', sorttype: 'int', hidden: true },
				{ name: 'LegendType', index: 'LegendType', sorttype: 'text', hidden: true },
				{ name: 'LegendCharacter', index: 'LegendCharacter', sorttype: 'text', hidden: false },
				{ name: 'LegendLookupTableID', index: 'LegendLookupTableID', sorttype: 'int', hidden: true },
				{ name: 'LegendLookupColumnID', index: 'LegendLookupColumnID', sorttype: 'int', hidden: true },
				{ name: 'LegendLookupCodeID', index: 'LegendLookupCodeID', sorttype: 'int', hidden: true },
				{ name: 'LegendEventColumnID', index: 'LegendEventColumnID', sorttype: 'int', hidden: true },
				{ name: 'EventDesc1ColumnID', index: 'EventDesc1ColumnID', sorttype: 'int', hidden: true },
				{ name: 'EventDesc2ColumnID', index: 'EventDesc2ColumnID', sorttype: 'int', hidden: true },
				{ name: 'FilterHidden', index: 'FilterHidden', sorttype: 'text', hidden: true },
				{ name: 'FilterName', index: 'FilterName', sorttype: 'text', hidden: false },
				{ name: 'EventStartSessionName', index: 'EventStartSessionName', sorttype: 'text', hidden: true },
				{ name: 'EventEndDateName', index: 'EventEndDateName', sorttype: 'text', hidden: true },
				{ name: 'EventEndSessionName', index: 'EventEndSessionName', sorttype: 'text', hidden: false },
				{ name: 'EventDurationName', index: 'EventDurationName', sorttype: 'text', hidden: false },
				{ name: 'LegendTypeName', index: 'LegendTypeName', sorttype: 'text', hidden: false },
				{ name: 'EventDesc1ColumnName', index: 'EventDesc1ColumnName', sorttype: 'text', hidden: false },
				{ name: 'EventDesc2ColumnName', index: 'EventDesc2ColumnName', sorttype: 'text', hidden: false }
			],
			rowNum: 10,
			autowidth: true,
			rowTotal: 50,
			rowList: [10, 20, 30],
			shrinkToFit: true,
			pager: '#pcrud',
			sortname: 'Name',
			loadonce: true,
			viewrecords: true,
			sortorder: "desc",
			editurl: 'server.php', // this is dummy existing url
			ondblClickRow: function (rowID) {
				eventEdit(rowID);
			}
		});
		$("#Events").jqGrid('navGrid', '#pcrud', {});

	}

	function	attachSortOrderColumns() {
	
		jQuery("#SortOrderColumns").jqGrid({

			datatype: 'jsonstring',
			datastr: '@Model.Events.ToJsonResult',
			mtype: 'GET',
			jsonReader: {
				root: "rows", //array containing actual data
				page: "page", //current page
				total: "total", //total pages for the query
				records: "records", //total number of records
				repeatitems: false,
				id: "id" //index of the column with the PK in it
			},
			colNames: ['EventKey', 'Name'],
			colModel: [
									{ name: 'EventKey', width: 50, key: true },
									{ name: 'Name', index: 'Name', width: 300 }],
			viewrecords: true,
			width: 400,
			sortname: 'Name',
			sortorder: "desc"
		});
	}

	function eventAdd() {
	}


	function eventEdit(rowID) {

		var gridData = $("#CalendarEvents").getRowData(rowID);
		OpenHR.OpenDialog("Reports/EditCalendarEvent", "divPopupReportDefinition", gridData);

		return;


		var sURL;
		var frmEvent = OpenHR.getForm("workframe", "frmEventDetails");
		frmEvent.eventAction.value = "NEW";
		frmEvent.eventID.value = getEventKey();
		frmEvent.eventFilterHidden.value = "";

		if (frmDefinition.grdEvents.Rows < 999) {
			sURL = "util_def_calendarreportdates_main" +
				"?eventAction=" + escape(frmEvent.eventAction.value) +
				"&eventName=" + escape(frmEvent.eventName.value) +
				"&eventID=" + escape(frmEvent.eventID.value) +
				"&eventTableID=" + escape(frmEvent.eventTableID.value) +
				"&eventTable=" + escape(frmEvent.eventTable.value) +
				"&eventFilterID=" + escape(frmEvent.eventFilterID.value) +
				"&eventFilter=" + escape(frmEvent.eventFilter.value) +
				"&eventFilterHidden=" + escape(frmEvent.eventFilterHidden.value) +
				"&eventStartDateID=" + escape(frmEvent.eventStartDateID.value) +
				"&eventStartDate=" + escape(frmEvent.eventStartDate.value) +
				"&eventStartSessionID=" + escape(frmEvent.eventStartSessionID.value) +
				"&eventStartSession=" + escape(frmEvent.eventStartSession.value) +
				"&eventEndDateID=" + escape(frmEvent.eventEndDateID.value) +
				"&eventEndDate=" + escape(frmEvent.eventEndDate.value) +
				"&eventEndSessionID=" + escape(frmEvent.eventEndSessionID.value) +
				"&eventEndSession=" + escape(frmEvent.eventEndSession.value) +
				"&eventDurationID=" + escape(frmEvent.eventDurationID.value) +
				"&eventDuration=" + escape(frmEvent.eventDuration.value) +
				"&eventLookupType=" + escape(frmEvent.eventLookupType.value) +
				"&eventKeyCharacter=" + escape(frmEvent.eventKeyCharacter.value) +
				"&eventLookupTableID=" + escape(frmEvent.eventLookupTableID.value) +
				"&eventLookupColumnID=" + escape(frmEvent.eventLookupColumnID.value) +
				"&eventLookupCodeID=" + escape(frmEvent.eventLookupCodeID.value) +
				"&eventTypeColumnID=" + escape(frmEvent.eventTypeColumnID.value) +
				"&eventDesc1ID=" + escape(frmEvent.eventDesc1ID.value) +
				"&eventDesc1=" + escape(frmEvent.eventDesc1.value) +
				"&eventDesc2ID=" + escape(frmEvent.eventDesc2ID.value) +
				"&eventDesc2=" + escape(frmEvent.eventDesc2.value) +
				"&relationNames=" + escape(frmEvent.relationNames.value);
			//openDialogCalEvent(sURL, 650, 500, "no", "no");
			openDialog(sURL, (screen.width) / 3.4, (screen.height) / 1.4, "no", "no");

//			frmUseful.txtChanged.value = 1;
		} else {
			var sMessage = "";
			sMessage = "The maximum of 999 events has been selected.";
			OpenHR.messageBox(sMessage, 64, "Calendar Reports");
		}


	}

</script>