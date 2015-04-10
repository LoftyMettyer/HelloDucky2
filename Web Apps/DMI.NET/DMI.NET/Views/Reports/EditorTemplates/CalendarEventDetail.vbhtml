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
		<div class="displaytable">
			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					<label>
						@Html.DisplayNameFor(Function(model) model.Name)
					</label>
				</div>

				<div class="tablecell">
					@Html.TextBoxFor(Function(model) model.Name, New With {.id = "EventName", .class = ""})
				</div>
			</div>

			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					@Html.LabelFor(Function(model) model.TableID)
				</div>

				<div class="tablecell">
					@Html.TableDropdown("TableID", "EventTableID", Model.TableID, Model.AvailableTables, "changeEventTable();")
				</div>
			</div>
			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					<label class="">Filter :</label>
					<input type="hidden" id="txtEventFilterID" name="FilterID" value="@Model.FilterID" />
				</div>
				<div class="tablecell">
					@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtEventFilter", .readonly = "true"})
					@Html.EllipseButton("cmdBasePicklist", "selectEventFilter()", True)
				</div>
			</div>
		</div>
	</fieldset>

	<fieldset class="">
		<legend class="fontsmalltitle">Event Start :</legend>
		<div class="displaytable">
			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					@Html.LabelFor(Function(model) model.EventStartDateID)
				</div>
				<div class="tablecell ">
					@Html.ColumnDropdownFor(Function(m) m.EventStartDateID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlDate}, Nothing)
				</div>
			</div>
			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					@Html.LabelFor(Function(model) model.EventStartSessionID)
				</div>
				<div class="tablecell">
					@Html.ColumnDropdownFor(Function(m) m.EventStartSessionID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .Size = 2}, Nothing)
				</div>
			</div>
		</div>
	</fieldset>

	<fieldset class="">
		<legend class="fontsmalltitle">Event End :</legend>
		<div class="displaytable">
			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.None), Model.EventEndType = CalendarEventEndType.None, New With {.onclick = "changeEventEndType()"})
					None
				</div>
				<div class="tablecell"></div>
			</div>

			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.EndDate), Model.EventEndType = CalendarEventEndType.EndDate, New With {.onclick = "changeEventEndType()"})
					End Date
				</div>
			</div>

			<div class="tablerow">								
					<div class="tablecell width35 padleft45">@Html.LabelFor(Function(model) model.EventEndDateID)</div>
					<div class="tablecell">@Html.ColumnDropdownFor(Function(m) m.EventEndDateID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlDate, .AddNone = True}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})</div>				
			</div>

			<div class="tablerow">
				<div class="tablecell width35 padleft45">@Html.LabelFor(Function(model) model.EventEndSessionID)</div>
				<div class="tablecell">@Html.ColumnDropdownFor(Function(m) m.EventEndSessionID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .Size = 2}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})</div>
			</div>

			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.Duration), Model.EventEndType = CalendarEventEndType.Duration, New With {.onclick = "changeEventEndType()"})
					Duration
				</div>
			</div>

			<div class="tablerow">
				<div class="tablecell width35 padleft45">Length :</div>
				<div class="tablecell">@Html.ColumnDropdownFor(Function(m) m.EventDurationID, New ColumnFilter() With {.TableID = Model.TableID, .IsNumeric = True, .AddNone = True}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})</div>
			</div>
		</div>
	</fieldset>
</div>

<div class="width50 floatleft ">
	<fieldset class="">
		<legend class="fontsmalltitle">Key :</legend>
		<div class="displaytable">
			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					@Html.RadioButton("LegendType", CInt(CalendarLegendType.Character), Model.LegendType = CalendarLegendType.Character, New With {.id = "legendType_character", .onclick = "changeEventLegendType('Character')"})
					@Html.DisplayNameFor(Function(model) model.LegendCharacter)
				</div>
				<div class="tablecell">
					@Html.TextBoxFor(Function(model) model.LegendCharacter, New With {.maxlength = 2})
				</div>
			</div>

			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					@Html.RadioButton("LegendType", CInt(CalendarLegendType.LookupTable), Model.LegendType = CalendarLegendType.LookupTable, New With {.id = "legendType_lookuptable", .onclick = "changeEventLegendType('LookupTable')"})
					Lookup Table
				</div>
			</div>

			<div class="tablerow">
				<div class="tablecell width35 padleft45">@Html.DisplayNameFor(Function(model) model.LegendEventColumnID)</div>
				<div class="tablecell">
					@Html.ColumnDropdownFor(Function(m) m.LegendEventColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .ColumnType = ColumnType.Lookup}, _
													 New With {.class = "eventLegendLookupLive", .disabled = (Model.EventEndType = CalendarEventEndType.EndDate), .onchange = "changeLegendEventColumnID()"})
					<select disabled class="eventLegendLookupDummy"><option>None</option></select>
				</div>
			</div>
			<div class="tablerow">
				<div class="tablecell width35 padleft45">@Html.DisplayNameFor(Function(model) model.LegendLookupTableID)</div>
				<div class="tablecell">
					@Html.LookupTableDropdown("LegendLookupTableID", "LegendLookupTableID", Model.LegendLookupTableID, "changeEventLookupTable();" _
																	, New With {.class = "eventLegendLookupLive"})
					<select disabled class="eventLegendLookupDummy"><option>None</option></select>
				</div>
			</div>
			<div class="tablerow">
				<div class="tablecell width35 padleft45">@Html.DisplayNameFor(Function(model) model.LegendLookupColumnID)</div>
				<div class="tablecell">
					@Html.ColumnDropdownFor(Function(m) m.LegendLookupColumnID _
																	, New ColumnFilter() With {.TableID = Model.LegendLookupTableID, .DataType = ColumnDataType.sqlVarChar} _
																	, New With {.class = "eventLegendLookupLive"})
					<select disabled class="eventLegendLookupDummy"><option>None</option></select>
				</div>
			</div>
			<div class="tablerow">
				<div class="tablecell width35 padleft45">@Html.DisplayNameFor(Function(model) model.LegendLookupCodeID)</div>
				<div class="tablecell">
					@Html.ColumnDropdownFor(Function(m) m.LegendLookupCodeID _
																	, New ColumnFilter() With {.TableID = Model.LegendLookupTableID, .DataType = ColumnDataType.sqlVarChar} _
																	, New With {.class = "eventLegendLookupLive"})
					<select disabled class="eventLegendLookupDummy"><option>None</option></select>
				</div>
			</div>
		</div>
	</fieldset>

	<fieldset class="">
		<legend class="fontsmalltitle">Event Description :</legend>
		<div class="displaytable">
			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					@Html.DisplayNameFor(Function(model) model.EventDesc1ColumnID)
				</div>
				<div class="tablecell">
					@Html.ColumnDropdownFor(Function(m) m.EventDesc1ColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .AddNone = True, .ShowFullName = True, .IncludeParents = True, .ExcludeOleAndPhoto = True}, Nothing)
				</div>

			</div>
			<div class="tablerow">
				<div class="tablecell width35 padleft20">
					@Html.DisplayNameFor(Function(model) model.EventDesc2ColumnID)
				</div>
				<div class="tablecell">
					@Html.ColumnDropdownFor(Function(m) m.EventDesc2ColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .AddNone = True, .ShowFullName = True, .IncludeParents = True, .ExcludeOleAndPhoto = True}, Nothing)
				</div>
			</div>
		</div>
	</fieldset>
</div>

<div id="divCalendarDetailButtons" class="width100 floatright">
	<input type="button" id="butEventEditCancel" value="Cancel" onclick="closeThisCalendarEvent();" />
	<input type="button" id="butEventEditOK" value="OK" onclick="postThisCalendarEvent();" />
</div>

@code
	Html.EndForm()
End Code

<script type="text/javascript">

	$(function () {
		refreshCalendarEventDisplay();		

		if (isDefinitionReadOnly()) {
			$("#frmPostCalendarEvent input").prop('disabled', "disabled");
			$("#frmPostCalendarEvent select").prop('disabled', "disabled");
			$("#frmPostCalendarEvent :button").prop('disabled', "disabled");
		}

		button_disable($("#butEventEditCancel")[0], false);

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
			if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower') {
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

		}, getPopupWidth(), getPopupHeight());

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

		if ($("#EventStartDateID")[0].length == 0) {
			OpenHR.modalMessage("A valid start date column has not been selected.");
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
		
		var legendEventColumnID = $("#LegendEventColumnID").val()
		if (legendEventColumnID == null) { legendEventColumnID = 0 }

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
			EventKey: '@Model.EventKey',
			ReportID: '@Model.ReportID',
			ReportType: '@CInt(Model.ReportType)',
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
			LegendEventColumnID: legendEventColumnID,
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
		grid.setGridParam({ sortname: 'EventKey', sortorder: "Asc" }).trigger('reloadGrid');
		grid.jqGrid("setSelection", '@Model.EventKey');


		setViewAccess('FILTER', $("#CalendarEventsViewAccess"), $("#FilterViewAccess").val(), $("#EventName").val());

		// Post to server
		OpenHR.postData("Reports/PostCalendarEvent", datarow)

		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();

		// enable save button
		enableSaveButton();

	}

	function closeThisCalendarEvent() {
		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();
	}

</script>



