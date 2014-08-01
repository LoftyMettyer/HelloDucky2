@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Imports DMI.NET.ViewModels.Reports
@Inherits System.Web.Mvc.WebViewPage(Of CalendarEventDetailViewModel)

@Code
	Html.BeginForm("PostCalendarEvent", "Reports", FormMethod.Post, New With {.id = "frmPostCalendarEvent"})
End Code

<div class="width50 floatleft">

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
		@Html.TableDropdown("TableID", "EventTableID", Model.TableID, Model.AvailableTables, "changeEventTable(event);")

		<br/>
		<input type="hidden" id="txtEventFilterID" name="FilterID" value="@Model.FilterID" />
		@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtEventFilter", .readonly = "true"})
		@Html.EllipseButton("cmdBasePicklist", "selectEventFilter()", True)

	</fieldset>

	<fieldset>
		<legend>Event Start :</legend>

		@Html.LabelFor(Function(model) model.EventStartDateID)
		@Html.ColumnDropdownFor(Function(m) m.EventStartDateID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlDate}, Nothing)
		<br/>
		@Html.LabelFor(Function(model) model.EventStartSessionID)
		@Html.ColumnDropdownFor(Function(m) m.EventStartSessionID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .Size = 2}, Nothing)

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
		@Html.ColumnDropdownFor(Function(m) m.EventEndDateID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlDate, .AddNone = True}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})

		<br />
		@Html.LabelFor(Function(model) model.EventEndSessionID)
		@Html.ColumnDropdownFor(Function(m) m.EventEndSessionID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .Size = 2}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})

		<br/>
		@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.Duration), Model.EventEndType = CalendarEventEndType.Duration, New With {.onclick = "changeEventEndType('duration')"})
		Duration
		@Html.ColumnDropdownFor(Function(m) m.EventDurationID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlNumeric, .AddNone = True}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})


	</fieldset>

</div>

<div class="width50 floatleft">

	<fieldset>
		<legend>Key</legend>

		@Html.RadioButton("LegendType", CInt(CalendarLegendType.Character), Model.LegendType = CalendarLegendType.Character, New With {.onclick = "changeEventLegendType('Character')"})
		@Html.DisplayNameFor(Function(model) model.LegendCharacter)
		@Html.TextBoxFor(Function(model) model.LegendCharacter, New With {.maxlength = 2})
		<br />

		@Html.RadioButton("LegendType", CInt(CalendarLegendType.LookupTable), Model.LegendType = CalendarLegendType.LookupTable, New With {.onclick = "changeEventLegendType('LookupTable')"})
		Lookup Table
		<br />

		@Html.DisplayNameFor(Function(model) model.LegendEventColumnID)
		@Html.ColumnDropdownFor(Function(m) m.LegendEventColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .ColumnType = ColumnType.Lookup}, _
													 New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})

		<br />
		<br />

		@Html.DisplayNameFor(Function(model) model.LegendLookupTableID)
		@Html.LookupTableDropdown("LegendLookupTableID", "LegendLookupTableID", Model.LegendLookupTableID, "changeEventLookupTable(event);")

		<br />
		@Html.DisplayNameFor(Function(model) model.LegendLookupColumnID)
		@Html.ColumnDropdownFor(Function(m) m.LegendLookupColumnID, New ColumnFilter() _
													 With {.TableID = Model.LegendLookupTableID, .DataType = ColumnDataType.sqlVarChar}, Nothing)

		<br />
		@Html.DisplayNameFor(Function(model) model.LegendLookupCodeID)
		@Html.ColumnDropdownFor(Function(m) m.LegendLookupCodeID, New ColumnFilter() _
													 With {.TableID = Model.LegendLookupTableID, .DataType = ColumnDataType.sqlVarChar}, Nothing)

		<br />

	</fieldset>

	<fieldset>
		<legend>Event Description</legend>
		@Html.DisplayNameFor(Function(model) model.EventDesc1ColumnID)
		@Html.ColumnDropdownFor(Function(m) m.EventDesc1ColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .ShowFullName = True, .IncludeParents = True}, Nothing)

		<br />
		@Html.DisplayNameFor(Function(model) model.EventDesc2ColumnID)
		@Html.ColumnDropdownFor(Function(m) m.EventDesc2ColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .ShowFullName = True, .IncludeParents = True}, Nothing)

	</fieldset>

</div>

<div class="width100 floatleft">
	<input type="button" value="OK" onclick="postThisCalendarEvent();" />
	<input type="button" value="Cancel" onclick="closeThisCalendarEvent();" />
</div>

@code
	Html.EndForm()
End Code



<script type="text/javascript">

	$(function () {
		changeEventLegendType('@Model.LegendType')
	});




	function changeEventLookupTable(event) {

		var frmSubmit = $("#frmPostCalendarEvent");
		OpenHR.submitForm(frmSubmit, "divPopupReportDefinition", null, null, "Reports/ChangeEventLookupTable");

	}

	function selectEventFilter() {

		var tableID = $("#EventTableID option:selected").val();
		var currentID = $("#txtEventFilterID").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name) {
			$("#txtEventFilterID").val(id);
			$("#txtEventFilter").val(name);
		});

	}

	function changeEventEndType(endType) {

		combo_disable("#EventEndDateID", true)
		combo_disable("#EventEndSessionID", true)
		combo_disable("#EventDurationID", true)

		switch (endType) {
			case "duration":
				combo_disable("#EventDurationID", false)
				break;
			case "enddate":
				combo_disable("#EventEndDateID", false)
				combo_disable("#EventEndSessionID", false)
				break;
		}
	}

	function changeEventLegendType(type) {

		combo_disable("#LegendEventColumnID", true);
		combo_disable("#LegendLookupTableID", true);
		combo_disable("#LegendLookupColumnID", true);
		combo_disable("#LegendLookupCodeID", true);
		text_disable("#LegendCharacter", true);

		switch (type) {
			case "Character":
				text_disable("#LegendCharacter", false);
				break;

			case "LookupTable":
				combo_disable("#LegendEventColumnID", false);
				combo_disable("#LegendLookupTableID", false);
				combo_disable("#LegendLookupColumnID", false);
				combo_disable("#LegendLookupCodeID", false);
				break;
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
			TableName: $("#EventTableID option:selected").text(),
			FilterID: $("#txtEventFilterID").val(),
			EventEndType: $("#EventEndType:checked").val(),
			EventStartDateID: $("#EventStartDateID").val(),
			EventStartSessionID: $("#EventStartSessionID").val(),
			EventEndDateID: $("#EventEndDateID").val(),
			EventEndSessionID: $("#EventEndSessionID").val(),
			EventDurationID: $("#EventDurationID").val(),
			LegendType: $("#LegendType:checked").val(),
			LegendCharacter: $("#LegendCharacter").val(),
			LegendLookupTableID: $("#LegendLookupTableID").val(),
			LegendLookupColumnID: legendLookupColumnID,
			LegendLookupCodeID: legendLookupCodeID,
			LegendEventColumnID: $("#LegendEventColumnID").val(),
			EventDesc1ColumnID: $("#EventDesc1ColumnID").val(),
			EventDesc2ColumnID: $("#EventDesc2ColumnID").val(),
			FilterHidden: $("#FilterHidden").val(),
			FilterName: $("#txtEventFilter").val(),
			EventStartDateName: $("#EventStartDateID option:selected").text(),
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



