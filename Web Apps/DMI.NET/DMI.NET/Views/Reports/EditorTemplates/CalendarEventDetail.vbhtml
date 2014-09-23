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
				<td class="" colspan="2">@Html.ColumnDropdownFor(Function(m) m.EventDurationID, New ColumnFilter() With {.TableID = Model.TableID, .IsNumeric = True, .AddNone = True}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})</td>
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
					@Html.RadioButton("LegendType", CInt(CalendarLegendType.Character), Model.LegendType = CalendarLegendType.Character, New With {.onclick = "changeEventLegendType('Character')"})
					@Html.DisplayNameFor(Function(model) model.LegendCharacter)
				</td>
				<td class="width60">
					@Html.TextBoxFor(Function(model) model.LegendCharacter, New With {.maxlength = 2})
				</td>
				<td></td>
			</tr>
			<tr>
				<td>
					@Html.RadioButton("LegendType", CInt(CalendarLegendType.LookupTable), Model.LegendType = CalendarLegendType.LookupTable, New With {.onclick = "changeEventLegendType('LookupTable')"})
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
													 New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate), .onchange = "changeLegendEventColumnID()"})
				</td>
			</tr>
			<tr>
				<td class="padleft20">@Html.DisplayNameFor(Function(model) model.LegendLookupTableID)</td>
				<td class="" colspan="2">@Html.LookupTableDropdown("LegendLookupTableID", "LegendLookupTableID", Model.LegendLookupTableID, "changeEventLookupTable();")</td>
			</tr>
			<tr>
				<td class="padleft20">@Html.DisplayNameFor(Function(model) model.LegendLookupColumnID)</td>
				<td class="" colspan="2">
					@Html.ColumnDropdownFor(Function(m) m.LegendLookupColumnID, New ColumnFilter() _
													 With {.TableID = Model.LegendLookupTableID, .DataType = ColumnDataType.sqlVarChar}, Nothing)
				</td>
			</tr>
			<tr>
				<td class="padleft20">@Html.DisplayNameFor(Function(model) model.LegendLookupCodeID)</td>
				<td class="" colspan="2">
					@Html.ColumnDropdownFor(Function(m) m.LegendLookupCodeID, New ColumnFilter() _
													 With {.TableID = Model.LegendLookupTableID, .DataType = ColumnDataType.sqlVarChar}, Nothing)
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

<div id="eventEventInformationContainer" class="width50 floatleft bgpink hidden">
	<fieldset>
		<legend class="fontsmalltitle">Event :</legend>

		<div class="">
			<label>
				@Html.DisplayNameFor(Function(model) model.Name)
			</label>
			@Html.TextBoxFor(Function(model) model.Name, New With {.id = "EventName"})
		</div>

		<div class="">
			@Html.LabelFor(Function(model) model.TableID)
			@Html.TableDropdown("TableID", "EventTableID", Model.TableID, Model.AvailableTables, "changeEventTable();")
		</div>

		<div class="">
			<input type="hidden" id="txtEventFilterID" name="FilterID" value="@Model.FilterID" />
			<label>Filter :</label>
			@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtEventFilter", .readonly = "true"})
			@Html.EllipseButton("cmdBasePicklist", "selectEventFilter()", True)
		</div>
	</fieldset>

	<fieldset>
		<legend class="fontsmalltitle">Event Start :</legend>

		<div class="">
			@Html.LabelFor(Function(model) model.EventStartDateID)
			@Html.ColumnDropdownFor(Function(m) m.EventStartDateID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlDate}, Nothing)
		</div>

		<div class="">
			@Html.LabelFor(Function(model) model.EventStartSessionID)
			@Html.ColumnDropdownFor(Function(m) m.EventStartSessionID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .Size = 2}, Nothing)
		</div>
	</fieldset>

	<fieldset>
		<legend class="fontsmalltitle">Event End :</legend>
		<div>
			<div class="optiongrouppadding">
				@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.None), Model.EventEndType = CalendarEventEndType.None, New With {.onclick = "changeEventEndType()"})
				None
			</div>
			<div class="optiongrouppadding">
				@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.EndDate), Model.EventEndType = CalendarEventEndType.EndDate, New With {.onclick = "changeEventEndType()"})
				End Date
			</div>

			<div class="">
				<label>
					@Html.LabelFor(Function(model) model.EventEndDateID)
				</label>
				@Html.ColumnDropdownFor(Function(m) m.EventEndDateID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlDate, .AddNone = True}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})
			</div>

			<div class="">
				<label>
					@Html.LabelFor(Function(model) model.EventEndSessionID)
				</label>
				@Html.ColumnDropdownFor(Function(m) m.EventEndSessionID, New ColumnFilter() With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .Size = 2}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})
			</div>

			<div class="optiongrouppadding">
				@Html.RadioButton("EventEndType", CInt(CalendarEventEndType.Duration), Model.EventEndType = CalendarEventEndType.Duration, New With {.onclick = "changeEventEndType()"})
				Duration
			</div>

			<div class="">
				<label>
					Length :
				</label>
				@Html.ColumnDropdownFor(Function(m) m.EventDurationID, New ColumnFilter() With {.TableID = Model.TableID, .IsNumeric = True, .AddNone = True}, New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate)})
			</div>
		</div>
	</fieldset>
</div>

<div id="eventEventInformationContainerRight" class=" width50 floatleft bgaqua hidden">
	<fieldset>
		<legend class="fontsmalltitle">Key</legend>

		<div class="optiongrouppadding">
			@Html.RadioButton("LegendType", CInt(CalendarLegendType.Character), Model.LegendType = CalendarLegendType.Character, New With {.onclick = "changeEventLegendType('Character')"})
			@Html.DisplayNameFor(Function(model) model.LegendCharacter)
			@Html.TextBoxFor(Function(model) model.LegendCharacter, New With {.maxlength = 2})
		</div>

		<div class="optiongrouppadding">
			@Html.RadioButton("LegendType", CInt(CalendarLegendType.LookupTable), Model.LegendType = CalendarLegendType.LookupTable, New With {.onclick = "changeEventLegendType('LookupTable')"})
			Lookup Table
		</div>

		<div class="">
			<label>
				@Html.DisplayNameFor(Function(model) model.LegendEventColumnID)
			</label>
			@Html.ColumnDropdownFor(Function(m) m.LegendEventColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .ColumnType = ColumnType.Lookup}, _
													 New With {.disabled = (Model.EventEndType = CalendarEventEndType.EndDate), .onchange = "changeLegendEventColumnID()"})
		</div>

		<div class="">
			<label>
				@Html.DisplayNameFor(Function(model) model.LegendLookupTableID)
			</label>
			@Html.LookupTableDropdown("LegendLookupTableID", "LegendLookupTableID", Model.LegendLookupTableID, "changeEventLookupTable();")
		</div>

		<div class="">
			<label>
				@Html.DisplayNameFor(Function(model) model.LegendLookupColumnID)
			</label>
			@Html.ColumnDropdownFor(Function(m) m.LegendLookupColumnID, New ColumnFilter() _
													 With {.TableID = Model.LegendLookupTableID, .DataType = ColumnDataType.sqlVarChar}, Nothing)
		</div>

		<div class="">
			<label>
				@Html.DisplayNameFor(Function(model) model.LegendLookupCodeID)
			</label>
			@Html.ColumnDropdownFor(Function(m) m.LegendLookupCodeID, New ColumnFilter() _
													 With {.TableID = Model.LegendLookupTableID, .DataType = ColumnDataType.sqlVarChar}, Nothing)
		</div>
	</fieldset>

	<fieldset>
		<legend class="fontsmalltitle">Event Description</legend>

		<div class="">
			<label>
				@Html.DisplayNameFor(Function(model) model.EventDesc1ColumnID)
			</label>
			@Html.ColumnDropdownFor(Function(m) m.EventDesc1ColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .ShowFullName = True, .IncludeParents = True}, Nothing)
		</div>

		<div class="">
			<label>
				@Html.DisplayNameFor(Function(model) model.EventDesc2ColumnID)
			</label>
			@Html.ColumnDropdownFor(Function(m) m.EventDesc2ColumnID, New ColumnFilter() _
													 With {.TableID = Model.TableID, .DataType = ColumnDataType.sqlVarChar, .AddNone = True, .ShowFullName = True, .IncludeParents = True}, Nothing)
		</div>
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
		changeEventLegendType('@Model.LegendType')
		//some styling - alter with caution		
		$("#frmPostCalendarEvent select").css('width', "100%");
		$('#LegendCharacter').width('30');
		$('fieldset').css('border', '0');
		$('table').attr('border', '0');
		$('table').addClass('padleft20');
		$('fieldset').css('padding-right', '0px');

		if (isDefinitionReadOnly()) {
			$("#frmPostCalendarEvent input").prop('disabled', "disabled");
			$("#frmPostCalendarEvent select").prop('disabled', "disabled");
			$("#frmPostCalendarEvent :button").prop('disabled', "disabled");
		}
	});

	function changeEventLookupTable() {

		$.ajax({
			url: 'Reports/GetAvailableColumnsForTable?TableID=' + $("#LegendLookupTableID").val(),
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
		var iLookupTableID = dropDown.options[dropDown.selectedIndex].attributes["data-lookuptableid"].value;
		$("#LegendLookupTableID").val(iLookupTableID);
		changeEventLookupTable();
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

	function changeEventEndType(endType) {

		var endType = $("input[name='EventEndType']:checked").val()

		combo_disable("#EventEndDateID", true)
		combo_disable("#EventEndSessionID", true)
		combo_disable("#EventDurationID", true)

		switch (endType) {
			case "1":
				combo_disable("#EventEndDateID", false)
				combo_disable("#EventEndSessionID", false)
				break;
			case "2":
				combo_disable("#EventDurationID", false)
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

	function changeEventTable() {

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
			LegendType: $("#LegendType:checked").val(),
			LegendCharacter: $("#LegendCharacter").val(),
			LegendLookupTableID: $("#LegendLookupTableID").val(),
			LegendLookupColumnID: legendLookupColumnID,
			LegendLookupCodeID: legendLookupCodeID,
			LegendEventColumnID: $("#LegendEventColumnID").val(),
			EventDesc1ColumnID: $("#EventDesc1ColumnID").val(),
			EventDesc2ColumnID: $("#EventDesc2ColumnID").val(),
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

	$(function () {
		changeEventEndType();
	});


</script>



