@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.CalendarReportModel)

<div class="left">

	@Html.TableFor("Events", Model.Events, Nothing)
	<table id="CalendarEvents"></table>
</div>

<div class="right">
	<input type="button" id="btnEventDetailsAdd" value="Add" />
	<br />
	<input type="button" id="btnEventDetailsEdit" value="Add All" />
	<br />
	<input type="button" id="btnEventDetailsRemove" value="Remove" />
	<br />
	<input type="button" id="btnEventDetailsRemoveAll" value="Remove All" />
</div>

<script type="text/javascript">

	$(function () {
		attachCalendarEventsGrid();
	})

	function	attachCalendarEventsGrid() {
	
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