@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Imports DMI.NET.ViewModels.Reports
@Inherits System.Web.Mvc.WebViewPage(Of CalendarEventDetailViewModel)

@Code
	Html.BeginForm("PostCalendarEvent", "Reports", FormMethod.Post, New With {.id = "frmPostCalendarEvent"})
End Code

@Html.HiddenFor(Function(model) model.ID, New With {.id = "CalendarEventID"})
@Html.HiddenFor(Function(model) model.EventKey, New With {.id = "EventKey"})
@Html.HiddenFor(Function(model) model.ReportID)
@Html.HiddenFor(Function(m) m.FilterViewAccess)

<div class="pageTitleDiv">
	<span class="pageTitle" id="PopupCalendarEventDetails_PageTitle">Event Information</span>
</div>

<div class="width50 floatleft ">
	<fieldset class="">
		<legend class="fontsmalltitle">Event :</legend>
		<table class="width100">
			<tr>
				<td class="width30">
					@Html.DisplayNameFor(Function(model) model.Name)
				</td>

				<td class="width70" colspan="2">
					@Html.TextBoxFor(Function(model) model.Name, New With {.id = "EventName", .class = "width99"})
				</td>
			</tr>

			<tr>
				<td class="width30">
					@Html.LabelFor(Function(model) model.TableID)
				</td>

				<td class="width70" colspan="2">
					@Html.TableDropdown("TableID", "EventTableID", Model.TableID, Model.AvailableTables, "changeEventTable();")
				</td>
			</tr>
			<tr>
				<td class="width30">
					Filter :
					<input type="hidden" id="txtEventFilterID" name="FilterID" value="@Model.FilterID" />
				</td>
				<td class="">
					@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtEventFilter", .readonly = "true", .style = "width:86%"})
					@Html.EllipseButton("cmdBasePicklist", "selectEventFilter()", True)
				</td>
			</tr>
		</table>
	</fieldset>

	<fieldset>
		<legend class="fontsmalltitle">Event Start :</legend>
		<table class="width100">
			<tr>
				<td class="width30">@Html.LabelFor(Function(model) model.EventStartDateID)</td>
				<td class="" colspan="2">@Html.ColumnDropdownFor(Function(m) m.EventStartDateID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlDate}, Nothing)</td>
			</tr>
			<tr>
				<td class="width30">@Html.LabelFor(Function(model) model.EventStartSessionID)</td>
				<td class="" colspan="2">@Html.ColumnDropdownFor(Function(m) m.EventStartSessionID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .Size = 2}, Nothing)</td>
			</tr>
		</table>
	</fieldset>

	<fieldset>
		<legend class="fontsmalltitle">Event End :</legend>
		<table class="width100">
			<tr>
				<td class="width30">
					@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.None), Model.EventEndType = CalendarEventEndType.None, New With {.onclick = "changeEventEndType()"})
					None
				</td>
				<td class="width60"></td>
				<td></td>
			</tr>
			<tr>
				<td class="width30">
					@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.EndDate), Model.EventEndType = CalendarEventEndType.EndDate, New With {.onclick = "changeEventEndType()"})
					End Date
				</td>
				<td></td>
				<td></td>
			</tr>

			<tr>
				<td class="width30 padleft20">@Html.LabelFor(Function(model) model.EventEndDateID)</td>
				<td class="" colspan="2">@Html.ColumnDropdownFor(Function(m) m.EventEndDateID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlDate, .AddNone = True}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})</td>
			</tr>

			<tr>
				<td class="width30 padleft20">@Html.LabelFor(Function(model) model.EventEndSessionID)</td>
				<td class="" colspan="2">@Html.ColumnDropdownFor(Function(m) m.EventEndSessionID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .Size = 2}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})</td>
			</tr>

			<tr>
				<td>
					@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.Duration), Model.EventEndType = CalendarEventEndType.Duration, New With {.onclick = "changeEventEndType()"})
					Duration
				</td>
				<td></td>
				<td></td>
			</tr>

			<tr>
				<td class="width30 padleft20">Length :</td>
				<td class="" colspan="2">
				@Html.ColumnDropdownFor(Function(m) m.EventDurationID, New ColumnFilter() With {.TableID = Model.TableID, .IsNumeric = True, .AddNone = True}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})
				</td>
			</tr>
		</table>

	</fieldset>
</div>
<div class="width50 floatleft ">
	<fieldset class="">
		<legend class="fontsmalltitle">Key :</legend>
		<table class="width100">
			<tr>
				<td class="width30">
					@Html.RadioButton("LegendType", CInt(CalendarLegendType.Character), Model.LegendType = CalendarLegendType.Character, New With {.id = "legendType_character", .onclick = "changeEventLegendType('Character')"})
					@Html.DisplayNameFor(Function(model) model.LegendCharacter)
				</td>
				<td class="width60">
					@Html.TextBoxFor(Function(model) model.LegendCharacter, New With {.maxlength = 2})
				</td>
				<td></td>
			</tr>
			<tr>
				<td>
					@Html.RadioButton("LegendType", CInt(CalendarLegendType.LookupTable), Model.LegendType = CalendarLegendType.LookupTable, New With {.id = "legendType_lookuptable", .onclick = "changeEventLegendType('LookupTable')"})
					Lookup Table
				</td>
				<td></td>
				<td></td>
			</tr>
			<tr>
				<td class="padleft20">@Html.DisplayNameFor(Function(model) model.LegendEventColumnID)</td>
				<td class="" colspan="2">
					@Html.ColumnDropdownFor(Function(m) m.LegendEventColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .ColumnType = ColumnType.Lookup}, _
													 New With {.class = "eventLegendLookupLive", .disabled = (Model.EventEndType = CalendarEventEndType.EndDate), .onchange = "changeLegendEventColumnID()"})
					<select disabled class="eventLegendLookupDummy"><option>None</option></select>
				</td>
			</tr>
			<tr>
				<td class="padleft20">@Html.DisplayNameFor(Function(model) model.LegendLookupTableID)</td>
				<td class="" colspan="2">
					@Html.LookupTableDropdown("LegendLookupTableID", "LegendLookupTableID", Model.LegendLookupTableID, "changeEventLookupTable();" _
																	, New With {.class = "eventLegendLookupLive"})
					<select disabled class="eventLegendLookupDummy"><option>None</option></select>
				</td>
			</tr>
			<tr>
				<td class="padleft20">@Html.DisplayNameFor(Function(model) model.LegendLookupColumnID)</td>
				<td class="" colspan="2">
					@Html.ColumnDropdownFor(Function(m) m.LegendLookupColumnID _
																	, New ColumnFilter() With {.TableID = Model.LegendLookupTableID, .DataType = ColumnDataType.sqlVarChar} _
																	, New With {.class = "eventLegendLookupLive"})
					<select disabled class="eventLegendLookupDummy"><option>None</option></select>
				</td>
			</tr>
			<tr>
				<td class="padleft20">@Html.DisplayNameFor(Function(model) model.LegendLookupCodeID)</td>
				<td class="" colspan="2">
					@Html.ColumnDropdownFor(Function(m) m.LegendLookupCodeID _
																	, New ColumnFilter() With {.TableID = Model.LegendLookupTableID, .DataType = ColumnDataType.sqlVarChar} _
																	, New With {.class = "eventLegendLookupLive"})
					<select disabled class="eventLegendLookupDummy"><option>None</option></select>
				</td>
			</tr>
		</table>
	</fieldset>

	<fieldset class="">
		<legend class="fontsmalltitle">Event Description :</legend>
		<table class="width100">
			<tr>
				<td class="width30">@Html.DisplayNameFor(Function(model) model.EventDesc1ColumnID)</td>
				<td class="" colspan="2">
					@Html.ColumnDropdownFor(Function(m) m.EventDesc1ColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .ShowFullName = True, .IncludeParents = True}, Nothing)
				</td>

			</tr>
			<tr>
				<td>
					@Html.DisplayNameFor(Function(model) model.EventDesc2ColumnID)
				</td>
				<td class="" colspan="2">
					@Html.ColumnDropdownFor(Function(m) m.EventDesc2ColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .ShowFullName = True, .IncludeParents = True}, Nothing)
				</td>

			</tr>
		</table>
	</fieldset>
</div>

<div id="divCalendarDetailButtons" class="width100 floatright">
	<input type="button" value="Cancel" onclick="closeThisCalendarEvent();" />
	<input type="button" value="OK" onclick="postThisCalendarEvent();" />
</div>

@code
	Html.EndForm()
End Code

<script type="text/javascript">

	$(function () {
		refreshCalendarEventDisplay();

		//some styling - alter with caution		
		$("#frmPostCalendarEvent select").css('width', "100%");
		$('#LegendCharacter').width('30');
		$('fieldset').css('border', '0');
		$('table').attr('border', '0');
		$('#frmPostCalendarEvent table').addClass('padleft20');
		$('fieldset').css('padding-right', '0px');

		if (isDefinitionReadOnly()) {
			$("#frmPostCalendarEvent input").prop('disabled', "disabled");
			$("#frmPostCalendarEvent select").prop('disabled', "disabled");
			$("#frmPostCalendarEvent :button").prop('disabled', "disabled");
		}
	});

	function changeEventLookupTable() {

		$.ajax({
			url: 'Reports/GetAvailableCharacterLookupsForTable?TableID=' + $("#LegendLookupTableID").val(),
			datatype: 'json',
			mtype: 'GET',
			success: function (json) {

				var optionItem = "";

				var options = '';
				for (var i = 0; i < json.length; i++) {
					optionItem += "<option value='" + json[i].ID + "'>" + json[i].Name + "</option>";
				}

				$("select#LegendLookupColumnID").html(optionItem);
				$("select#LegendLookupCodeID").html(optionItem);

			}
		});

	}

	function changeLegendEventColumnID() {

		var dropDown = $("#LegendEventColumnID")[0];

		if (dropDown.length == 0) {
			$("#legendType_character").prop('checked', 'checked');
		}
		else {
			var iLookupTableID = dropDown.options[dropDown.selectedIndex].attributes["data-lookuptableid"].value;
			$("#LegendLookupTableID").val(iLookupTableID);
			changeEventLookupTable();
		}

		refreshCalendarEventDisplay();

	}

	function selectEventFilter() {

		var tableID = $("#EventTableID option:selected").val();
		var currentID = $("#txtEventFilterID").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name, access) {
			if (access == "HD" && $("#owner") != '@Session("Username")') {
				$("#txtEventFilterID").val(0);
				$("#txtEventFilter").val('None');
				$("#FilterViewAccess").val('');
				OpenHR.modalMessage("The event filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#txtEventFilterID").val(id);
				$("#txtEventFilter").val(name);
				$("#FilterViewAccess").val(access);
			}

		}, 400, 400);

	}

	function changeEventEndType() {

		var eventEndType = $("input[name='EventEndType']:checked").val();

		if (eventEndType == "0" || eventEndType == "1") {
			$("#EventDurationID").val(0);
		}

		if (eventEndType == "0" || eventEndType == "2") {
			$("#EventEndDateID").val(0);
			$("#EventEndSessionID").val(0);
		}

		refreshCalendarEventDisplay();
	}

	function refreshCalendarEventDisplay() {

		var legendType = $("input[name='LegendType']:checked").val();
		var legendDropDown = $("#LegendEventColumnID")[0];

		text_disable("#LegendCharacter", (legendType != "0"));
		combo_disable("#LegendEventColumnID", (legendType != "1"));
		combo_disable("#LegendLookupTableID", (legendType != "1"));
		combo_disable("#LegendLookupColumnID", (legendType != "1"));
		combo_disable("#LegendLookupCodeID", (legendType != "1"));

		$("#legendType_lookuptable").attr('disabled', (legendDropDown.length == 0));
		if (legendType == "0") {
			$(".eventLegendLookupDummy").show();
			$(".eventLegendLookupLive").hide();
		}
		else {
			$(".eventLegendLookupLive").show();
			$(".eventLegendLookupDummy").hide();
		}


		var endType = $("input[name='EventEndType']:checked").val();
		combo_disable("#EventEndDateID", (endType != "1"));
		combo_disable("#EventEndSessionID", (endType != "1"));
		combo_disable("#EventDurationID", (endType != "2"));

	}

	function changeEventLegendType(type) {

		if (type == "LookupTable") {
			changeLegendEventColumnID();
		}

		refreshCalendarEventDisplay();

	}

	function changeEventTable() {

		// Reload dropdowns from server
		var frmSubmit = $("#frmPostCalendarEvent");
		OpenHR.submitForm(frmSubmit, "divPopupReportDefinition", null, null, "Reports/ChangeEventBaseTable", changeLegendEventColumnID);

	}

	function postThisCalendarEvent() {

		var legendTypeName;

		// Validation
		if ($("#EventName").val() == "") {
			OpenHR.modalMessage("You must give this event a name.");
			return false;
		}

		if ($('input:radio[name=EventEndType]:checked').val() == "1" && $("#EventEndDateID").val() == "0") {
			OpenHR.modalMessage("A valid end date column has not been selected.")
			return false;
		}

		if ($('input:radio[name=EventEndType]:checked').val() == "2" && $("#EventDurationID").val() == "0") {
			OpenHR.modalMessage("A valid duration column has not been selected.")
			return false;
		}

		var legendLookupColumnID = $("#LegendLookupColumnID").val();
		if (legendLookupColumnID == null) { legendLookupColumnID = 0 }

		var legendLookupCodeID = $("#LegendLookupCodeID").val()
		if (legendLookupCodeID == null) { legendLookupCodeID = 0 }

		if ($("input[name='LegendType']:checked").val() == "1") {
			legendTypeName = $("#LegendEventColumnID option:selected").text() + "." + $("#LegendLookupCodeID option:selected").text()
		}
		else {
			legendTypeName = $("#LegendCharacter").val();
		}

		var eventStartSessionName = "";
		if ($("#EventStartSessionID option:selected").val() > 0) {
			eventStartSessionName = $("#EventStartSessionID option:selected").text();
		}

		var eventEndDateName = "";
		if ($("#EventEndDateID option:selected").val() > 0) {
			eventEndDateName = $("#EventEndDateID option:selected").text();
		}

		var eventEndSessionName = "";
		if ($("#EventEndSessionID option:selected").val() > 0) {
			eventEndSessionName = $("#EventEndSessionID option:selected").text();
		}

		var eventDurationName = "";
		if ($("#EventDurationID option:selected").val() > 0) {
			eventDurationName = $("#EventDurationID option:selected").text();
		}

		var description1Text = "";
		if ($("#EventDesc1ColumnID option:selected").val() > 0) {
			description1Text = $("#EventDesc1ColumnID option:selected").text();
		}

		var description2Text = "";
		if ($("#EventDesc2ColumnID option:selected").val() > 0) {
			description2Text = $("#EventDesc2ColumnID option:selected").text();
		}

		var datarow = {
			ID: $("#CalendarEventID").val(),
			EventKey: 			'@Model.EventKey',
			ReportID: 			'@Model.ReportID',
			ReportType: 		'@CInt(Model.ReportType)',
			Name: $("#EventName").val(),
			TableID: $("#EventTableID").val(),
			TableName: $("#EventTableID option:selected").text(),
			FilterID: $("#txtEventFilterID").val(),
			FilterName: $("#txtEventFilter").val(),
			FilterViewAccess: $("#FilterViewAccess").val(),
			EventEndType: $("#EventEndType:checked").val(),
			EventStartDateID: $("#EventStartDateID").val(),
			EventStartSessionID: $("#EventStartSessionID").val(),
			EventEndDateID: $("#EventEndDateID").val(),
			EventEndSessionID: $("#EventEndSessionID").val(),
			EventDurationID: $("#EventDurationID").val(),
			LegendType: $("input:radio[name=LegendType]:checked").val(),
			LegendCharacter: $("#LegendCharacter").val(),
			LegendLookupTableID: $("#LegendLookupTableID").val(),
			LegendLookupColumnID: legendLookupColumnID,
			LegendLookupCodeID: legendLookupCodeID,
			LegendEventColumnID: $("#LegendEventColumnID").val(),
			EventDesc1ColumnID: $("#EventDesc1ColumnID").val(),
			EventDesc2ColumnID: $("#EventDesc2ColumnID").val(),
			EventStartDateName: $("#EventStartDateID option:selected").text(),
			EventStartSessionName: eventStartSessionName,
			EventEndDateName: eventEndDateName,
			EventEndSessionName: eventEndSessionName,
			EventDurationName: eventDurationName,
			LegendTypeName: legendTypeName,
			EventDesc1ColumnName: description1Text,
			EventDesc2ColumnName: description2Text
		};


		// Update client
		var grid = $("#CalendarEvents")
		grid.jqGrid('delRowData', '@Model.EventKey')
		grid.jqGrid('addRowData', '@Model.EventKey', datarow);
		grid.setGridParam({ sortname: 'EventKey', sortorder: "Asc"}).trigger('reloadGrid');
		grid.jqGrid("setSelection", '@Model.EventKey');


		setViewAccess('FILTER', $("#CalendarEventsViewAccess"), $("#FilterViewAccess").val(), $("#EventName").val());

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



