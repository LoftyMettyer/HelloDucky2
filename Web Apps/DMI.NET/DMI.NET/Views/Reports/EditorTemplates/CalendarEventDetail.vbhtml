﻿@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums
@Inherits System.Web.Mvc.WebViewPage(Of ViewModels.CalendarEventDetailViewModel)

@Code
	Html.BeginForm("PostCalendarEvent", "Reports", FormMethod.Post, New With {.id = "frmPostCalendarEvent"})
End Code

<div class="left">

	@Html.HiddenFor(Function(model) model.ID, New With {.id = "CalendarEventID"})
	@Html.HiddenFor(Function(model) model.EventKey, New With {.id = "EventKey"})
	@Html.HiddenFor(Function(model) model.ReportID)
	@Html.HiddenFor(Function(model) model.FilterHidden)

	<fieldset>
		<legend>Event :</legend>

		@Html.DisplayNameFor(Function(model) model.Name)
		@Html.TextBoxFor(Function(model) model.Name, New With {.id = "EventName"})
		<br />
		@Html.LabelFor(Function(model) model.TableID)
		@Html.TableDropdown("CalendarEventTableID", "EventTableID", Model.TableID, Model.AvailableTables, "changeEventTable(event);")

		<br/>
		<input type="hidden" id="txtEventFilterID" name="FilterID" value="@Model.FilterID" />
		@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "EventFilter", .readonly = "true"})
		@Html.EllipseButton("cmdBasePicklist", "selectRecordOption('event', 'filter')", True)

	</fieldset>

	<fieldset>
		<legend>Event Start :</legend>

		@Html.LabelFor(Function(model) model.EventStartDateID)
		@Html.ColumnDropdown2("EventStartDateID", Model.EventStartDateID, Model.TableID, SQLDataType.sqlDate, False, False)
		<br/>
		@Html.LabelFor(Function(model) model.EventStartSessionID)
		@Html.ColumnDropdown2("EventStartSessionID", Model.EventStartSessionID, Model.TableID, SQLDataType.sqlVarChar, True, False)

	</fieldset>

	<fieldset>
		<legend>Event End :</legend>

		@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.None), Model.EventEndType = CalendarEventEndType.None, New With {.onclick = "changeEventEndType('none')"})
		None
		<br />

		@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.EndDate), Model.EventEndType = CalendarEventEndType.EndDate, New With {.onclick = "changeEventEndType('enddate')"})
		End Date
		<br />
		@Html.LabelFor(Function(model) model.EventEndDateID)
		@Html.ColumnDropdown2("EventEndDateID", Model.EventEndDateID, Model.TableID, SQLDataType.sqlDate, True, False)
		<br />
		@Html.LabelFor(Function(model) model.EventEndSessionID)
		@Html.ColumnDropdown2("EventEndSessionID", Model.EventEndSessionID, Model.TableID, SQLDataType.sqlVarChar, True, False)
		<br/>
		@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.Duration), Model.EventEndType = CalendarEventEndType.Duration, New With {.onclick = "changeEventEndType('duration')"})
		Duration
		@Html.ColumnDropdown2("EventDurationID", Model.EventDurationID, Model.TableID, SQLDataType.sqlNumeric, True, False)
	</fieldset>

</div>

<div class="right">

	<fieldset>
		<legend>Key</legend>

		@Html.RadioButton("LegendType", CInt(CalendarLegendType.Character), Model.LegendType = CalendarLegendType.Character, New With {.onclick = "changeEventLegendType('char')"})
		@Html.DisplayNameFor(Function(model) model.LegendCharacter)
		@Html.TextBoxFor(Function(model) model.LegendCharacter)
		<br />

		@Html.RadioButton("LegendType", CInt(CalendarLegendType.LookupTable), Model.LegendType = CalendarLegendType.LookupTable, New With {.onclick = "changeEventLegendType('lookup')"})
		Lookup Table
		<br />

		@Html.DisplayNameFor(Function(model) model.LegendEventColumnID)
		@Html.ColumnDropdown2("LegendEventColumnID", Model.LegendEventColumnID, Model.TableID, SQLDataType.sqlVarChar, False, True)
		<br/>
		<br/>

		@Html.DisplayNameFor(Function(model) model.LegendLookupTableID)
		@Html.LookupTableDropdown("LegendLookupTableID", "LegendLookupTableID", Model.LegendLookupTableID)

		<br />
		@Html.DisplayNameFor(Function(model) model.LegendLookupColumnID)
		@Html.ColumnDropdown2("LegendLookupColumnID", Model.LegendLookupColumnID, Model.LegendLookupTableID, SQLDataType.sqlVarChar, False, False)
		<br />
		@Html.DisplayNameFor(Function(model) model.LegendLookupCodeID)
		@Html.ColumnDropdown2("LegendLookupCodeID", Model.LegendLookupCodeID, Model.LegendLookupTableID, SQLDataType.sqlVarChar, False, False)


		<br />

</fieldset>

	<fieldset>
		<legend>Event Description</legend>
		@Html.DisplayNameFor(Function(model) model.EventDesc1ColumnID)
		@Html.ColumnDropdown2("EventDesc1ColumnID", Model.EventDesc1ColumnID, Model.TableID, SQLDataType.sqlVarChar, True, False)
		<br/>
		@Html.DisplayNameFor(Function(model) model.EventDesc2ColumnID)
		@Html.ColumnDropdown2("EventDesc2ColumnID", Model.EventDesc2ColumnID, Model.TableID, SQLDataType.sqlVarChar, True, False)

</fieldset>

</div>

<br/>

<input type="button" value="OK" onclick="postThisCalendarEvent();" />
<input type="button" value="Cancel" onclick="closeThisCalendarEvent();" />

@code
	Html.EndForm()
End Code



<script type="text/javascript">

	function changeEventEndType(endType) {

		combo_disable("#EventEndDateID", true)
		combo_disable("#EventEndSessionID", true)
		combo_disable("#EventEndDurationID", true)

		switch (endType) {
			case "duration":
				combo_disable("#EventEndDurationID", false)
				break;
			case "enddate":
				combo_disable("#EventEndDateID", false)
				combo_disable("#EventEndSessionID", false)
				break;
		}
	}

	function changeEventLegendType(type) {

		combo_disable("#LegendEventColumnID", true)
		combo_disable("#LegendLookupTableID", true)
		combo_disable("#LegendLookupColumnID", true)
		combo_disable("#LegendLookupCodeID", true)
		text_disable("#LegendCharacter", true)

		switch (type) {
			case "char":
				text_disable("#LegendCharacter", false)

				break;
			case "lookup":
				combo_disable("#LegendEventColumnID", false)
				combo_disable("#LegendLookupTableID", false)
				combo_disable("#LegendLookupColumnID", false)
				combo_disable("#LegendLookupCodeID", false)
				break
		}

	}

	function changeEventTable(event) {

		// Reload dropdowns from server
		var frmSubmit = $("#frmPostCalendarEvent");
		OpenHR.submitForm(frmSubmit, "divPopupReportDefinition", null, null, "Reports/ChangeEventBaseTable");

	}

	function postThisCalendarEvent() {

		var legendLookupColumnID = $("#LegendLookupColumnID").val()
		if (legendLookupColumnID == null) { legendLookupColumnID = 0 }

		var legendLookupCodeID = $("#LegendLookupCodeID").val()
		if (legendLookupCodeID == null) { legendLookupCodeID = 0 }

		var datarow = {
			ID: $("#CalendarEventID").val(),
			EventKey: '@Model.EventKey',
			ReportID: '@Model.ReportID',
			ReportType: '@CInt(Model.ReportType)',
			Name: $("#EventName").val(),
			TableID: $("#EventTableID").val(),
			FilterID: $("#txtEventFilterID").val(),
			EventEndType: $("#EventEndType").val(),
			EventStartDateID: $("#EventStartDateID").val(),
			EventStartSessionID: $("#EventStartSessionID").val(),
			EventEndType: $("#EventEndType").val(),
			EventEndDateID: $("#EventEndDateID").val(),
			EventEndSessionID: $("#EventEndSessionID").val(),
			EventDurationID: $("#EventDurationID").val(),
			LegendType: $("#LegendType").val(),
			LegendCharacter: $("#LegendCharacter").val(),
			LegendLookupTableID: $("#LegendLookupTableID").val(),
			LegendLookupColumnID: legendLookupColumnID,
			LegendLookupCodeID: legendLookupCodeID,
			LegendEventColumnID: $("#LegendEventColumnID").val(),
			EventDesc1ColumnID: $("#EventDesc1ColumnID").val(),
			EventDesc2ColumnID: $("#EventDesc2ColumnID").val(),
			FilterHidden: $("#FilterHidden").val(),
			FilterName: $("#EventFilter").val(),
			EventStartSessionName: $("#EventStartSessionID option:selected").text(),
			EventEndDateName: $("#EventEndDateID option:selected").text(),
			EventEndSessionName: $("#EventEndSessionID option:selected").text(),
			EventDurationName: $("#EventDurationID option:selected").text(),
			LegendTypeName: $("#LegendEventColumnID option:selected").text(),
			EventDesc1ColumnName: $("#EventDesc1ColumnID option:selected").text(),
			EventDesc2ColumnName: $("#EventDesc2ColumnID option:selected").text()
		};

		// Update client
		$('#CalendarEvents').jqGrid('delRowData', '@Model.EventKey')
		var su = jQuery("#CalendarEvents").jqGrid('addRowData', '@Model.EventKey', datarow);

		// Post to server
		OpenHR.postData("Reports/PostCalendarEvent", datarow)

		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();


	}

	function closeThisCalendarEvent() {
		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();
	}

</script>



