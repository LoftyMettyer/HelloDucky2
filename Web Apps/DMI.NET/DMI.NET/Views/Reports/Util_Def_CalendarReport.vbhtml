@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Inherits System.Web.Mvc.WebViewPage(Of Models.CalendarReportModel)

@Code
	Layout = Nothing
End Code

<div>
	@Using (Html.BeginForm("util_def_calendarreport", "Reports", FormMethod.Post, New With {.id = "frmReportDefintion", .name = "frmReportDefintion"}))
		@Html.HiddenFor(Function(m) m.ID)
		@Html.HiddenFor(Function(m) m.Description3ViewAccess, New With {.class = "ViewAccess"})	  

		@<div id="tabs">
			<ul>
				<li><a href="#tabs-1">Definition</a></li>
				<li><a href="#report_definition_tab_eventdetails">Event Details</a></li>
				<li><a href="#report_definition_tab_reportdetails">Report Details</a></li>
				<li><a href="#report_definition_tab_order">Sort Order</a></li>
				<li><a href="#report_definition_tab_output">Output</a></li>
			</ul>

			<div id="tabs-1">
				@Code
				Html.RenderPartial("_Definition", Model)
				End Code

				<fieldset class="width50">
					<legend class="fontsmalltitle">Report Options :</legend>

					<fieldset>
						@Html.LabelFor(Function(m) m.Description1ID)
						<div class="width70 floatright">
							@Html.ColumnDropdownFor(Function(m) m.Description1ID, New ColumnFilter() With {.TableID = Model.BaseTableID, .AddNone = True}, New With {.class = "enableSaveButtonOnComboChange", .onchange = "validateDescriptions()"})
						</div>
					</fieldset>

					<fieldset>
						@Html.LabelFor(Function(m) m.Description2ID)
						<div class="width70 floatright">
							@Html.ColumnDropdownFor(Function(m) m.Description2ID, New ColumnFilter() With {.TableID = Model.BaseTableID, .AddNone = True}, New With {.class = "enableSaveButtonOnComboChange", .onchange = "validateDescriptions()"})
						</div>
					</fieldset>

					<fieldset>
						<div id="" class="floatleft">
							@Html.LabelFor(Function(m) m.Description3ID)
							@Html.HiddenFor(Function(m) m.Description3ID)
						</div>
						<div class="width70 floatright">
							<input class="floatright" type="button" id="cmdDescription3" value="..." onclick="selectDescription3()" />
							<div class="ellipsistextbox">
								<input class="floatleft" type="text" id="txtDescription3" value="@Model.Description3Name" disabled />
							</div>
						</div>
						<input type="hidden" id="txtBasePicklistID" name="picklistID" value="@Model.PicklistID" />
					</fieldset>


					<fieldset>					
						@Html.LabelFor(Function(m) m.RegionID)
						<div class="width70 floatright">
							@Html.ColumnDropdownFor(Function(m) m.RegionID, New ColumnFilter() With {.TableID = Model.BaseTableID, .AddDefault = True, .DataType = ColumnDataType.sqlVarChar}, New With {.id = "cboRegionID", .class = "selectRegionID width100 floatright enableSaveButtonOnComboChange", .onchange = "regionChange()"})
						</div>
					</fieldset>

					<fieldset>
						@Html.LabelFor(Function(m) m.Separator)
						<div class="width70 floatright">
							@Html.DropDownList("Separator", New SelectList(New List(Of String)() From {"None", "Space", ",", ".", "-", ":", ";", "/", "\", "#", "~", "^"}), New With {.class = "enableSaveButtonOnComboChange"})
							@Html.HiddenFor(Function(m) m.Separator, New With {.id = "ddlSeparator"})
							@Html.CheckBoxFor(Function(m) m.GroupByDescription, New With {.id = "chkGroupByDescription", .onclick = "selectGroupByDescription()"})
							@Html.LabelFor(Function(m) m.GroupByDescription, New With {.id = "label_GroupByDescription"})
						</div>
					</fieldset>
				</fieldset>
			</div>

			<div id="report_definition_tab_eventdetails">
				@Code
				Html.RenderPartial("_EventDetails", Model)
				End Code
			</div>

			<div id="report_definition_tab_reportdetails">
				@Code
				Html.RenderPartial("_ReportDetails", Model)
				End Code
			</div>

			<div id="report_definition_tab_order">
				@Code
				Html.RenderPartial("_SortOrder", Model)
				End Code
			</div>

			<div id="report_definition_tab_output">
				@Code
				Html.RenderPartial("_Output", Model.Output)
				End Code
			</div>		 	
		</div>
		@Html.AntiForgeryToken()
	End Using
</div>


<script type="text/javascript">

	function selectGroupByDescription() {
		var bGroupBySelected = $("#chkGroupByDescription").prop('checked');
		combo_disable((".selectRegionID"), bGroupBySelected);

		//If  the Base Table value is anything other than Primary table value then 'Include Bank Holidays', 'Working Days Only' and 'Show Bank Holidays' should disable.
		var bIsPersonnelRecords = ($("#BaseTableID option:selected").val() == '@SettingsConfig.Personnel_EmpTableID');
		if (bIsPersonnelRecords && bGroupBySelected) {
			enableDisableWorkingDaysOrHolidays(bGroupBySelected);
		}
		else if (bIsPersonnelRecords && bGroupBySelected == false) {
			enableDisableWorkingDaysOrHolidays(bGroupBySelected);
		}
		else if (bIsPersonnelRecords == false) {
			enableDisableWorkingDaysOrHolidays(!bIsPersonnelRecords);
		}
	}

	function regionChange() {
		selectWorkingDaysOrHolidays();
	}

	//'Seperator' should only enable if at least 2 descriptions have been entered.
	function validateDescriptions() {
		if (($("#Description1ID").val() == 0) || ($("#Description2ID").val() == 0)) {
			$("#Separator").prop('disabled', true);
			$("#Separator").val("None");
			$('#ddlSeparator').val("None");
		}
		else {
			$("#Separator").prop('disabled', false);
		}
	}

	function selectDescription3() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#Description3ID").val();

		OpenHR.modalExpressionSelect("CALC", tableID, currentID, function (id, name, access) {
			if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower') {
				$("#Description3ID").val(0);
				$("#txtDescription3").val('None');
				OpenHR.modalMessage("The report description calculation will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#Description3ID").val(id);
				$("#txtDescription3").val(name);
				setViewAccess('CALC', $("#Description3ViewAccess"), access, "report description");
				enableSaveButton();
			}
		}, getPopupWidth(), getPopupHeight());

	}

	$(function () {
		$("#tabs").tabs({
			activate: function (event, ui) {
				//Tab click event fired
				if (ui.newTab.text() == "Event Details") {
					//resize the Event Details grid to fit
					var workPageHeight = $('#workframeset').height();
					var gridTopPos = $('#divEventDetails').position().top;
					var tabHeight = $('#tabs>.ui-tabs-nav').outerHeight();
					var marginHeight = 40;
					var gridHeight = workPageHeight - gridTopPos - tabHeight - marginHeight;
					$("#CalendarEvents").jqGrid('setGridHeight', gridHeight);
				}
				if (ui.newTab.text() == "Sort Order") {
					//resize the Event Details grid to fit
					var workPageHeight = $('#workframeset').height();
					var gridTopPos = $('#divSortOrderDiv').position().top;
					var tabHeight = $('#tabs>.ui-tabs-nav').outerHeight();
					var marginHeight = 40;
					var gridHeight = workPageHeight - gridTopPos - tabHeight - marginHeight;
					$("#SortOrders").jqGrid('setGridHeight', gridHeight);

					var gridWidth = $('#divSortOrderDiv').width();
					$("#SortOrders").jqGrid('setGridWidth', gridWidth);
				}
			}
		});

		$('input[type=number]').numeric();
		$('#Description2ID,#Description1ID').css({ "width": "100%", "float": "right" });
		$('#description, #Name').css('width', $('#Description1ID').width());

		// If 'Group by Description' is ticked then 'Include Bank Holidays', 'Working Days Only' ,'Show Bank Holidays' and Region  should disable.	Or  If 'Include Bank Holidays', 'Working Days Only' or 'Show Bank Holidays'  are ticked OR 'Region selected index not equal to 0' than 'Group by Description' should disable.
		if (($("#ActionType").val() == '@UtilityActionType.Edit') || ($("#ActionType").val() == '@UtilityActionType.Copy' ) )
		{
			if ($('#chkGroupByDescription').prop('checked')) {
				combo_disable((".selectRegionID"), true);
				$('#IncludeBankHolidays').prop('disabled', true);
				$('#WorkingDaysOnly').prop('disabled', true);
				$('#ShowBankHolidays').prop('disabled', true);
				$("#label_IncludeBankHolidays").css('opacity', '0.5');
				$("#label_WorkingDaysOnly").css('opacity', '0.5');
				$("#label_ShowBankHolidays").css('opacity', '0.5');
			}
			else
			{
				enableDisableGroupByOrWorkingDaysOrHolidays();
			}
		}

		validateDescriptions();
		button_disable($("#btnSortOrderAdd")[0], isDefinitionReadOnly());

	});

	function submitForm() {
		var frmSubmit = $("#frmReportDefintion");
		OpenHR.submitForm(frmSubmit);
	}

	//If the Base Table value is Primary table value then 'Include Bank Holidays', 'Working Days Only' and 'Show Bank Holidays' should enable.
	function enableDisableGroupByOrWorkingDaysOrHolidays() {		
		var regionValue = $("#RegionID").val();
		var IsHolidaySelected = $('#IncludeBankHolidays').prop('checked') || $('#WorkingDaysOnly').prop('checked') || $('#ShowBankHolidays').prop('checked') || (regionValue != 0);
		$('#chkGroupByDescription').prop('checked', false);
		$('#chkGroupByDescription').prop('disabled', IsHolidaySelected);

		if (IsHolidaySelected) {
			$("#label_GroupByDescription").css("color", "#A59393");
		}
		else {
			$("#label_GroupByDescription").css("color", "#000000");
		}

		var bIsPersonnelRecords = ('@Model.BaseTableID' == '@SettingsConfig.Personnel_EmpTableID');
		if (bIsPersonnelRecords == false) {
			$('#IncludeBankHolidays').prop('checked', false);
			$('#WorkingDaysOnly').prop('checked', false);
			$('#ShowBankHolidays').prop('checked', false);
			$('#IncludeBankHolidays').prop('disabled', !bIsPersonnelRecords);
			$('#WorkingDaysOnly').prop('disabled', !bIsPersonnelRecords);
			$('#ShowBankHolidays').prop('disabled', !bIsPersonnelRecords);
			$("#label_IncludeBankHolidays").css('opacity', '0.5');
			$("#label_WorkingDaysOnly").css('opacity', '0.5');
			$("#label_ShowBankHolidays").css('opacity', '0.5');
		}

	}

	function selectWorkingDaysOrHolidays() {
		// If 'Include Bank Holidays', 'Working Days Only' or 'Show Bank Holidays'  are ticked OR 'Region selected index not equal to 0' than 'Group by Description' should disable.
		var regionValue = $("#RegionID").val();
		var bSelected = $('#IncludeBankHolidays').prop('checked') || $('#WorkingDaysOnly').prop('checked') || $('#ShowBankHolidays').prop('checked') || (regionValue != 0);
		$('#chkGroupByDescription').prop('checked', false);
		$('#chkGroupByDescription').prop('disabled', bSelected);
		if (bSelected) {
			$("#label_GroupByDescription").css("color", "#A59393");
		}
		else {
			$("#label_GroupByDescription").css("color", "#000000");
		}
	}

	//If the Base Table value is Primary table value then 'Include Bank Holidays', 'Working Days Only' and 'Show Bank Holidays' should enable.
	function enableDisableWorkingDaysOrHolidays(bDisabled) {
		$('#IncludeBankHolidays').prop('checked', false);
		$('#WorkingDaysOnly').prop('checked', false);
		$('#ShowBankHolidays').prop('checked', false);

		$('#IncludeBankHolidays').prop('disabled', bDisabled);
		$('#WorkingDaysOnly').prop('disabled', bDisabled);
		$('#ShowBankHolidays').prop('disabled', bDisabled);

		if (bDisabled) {
			$("#label_IncludeBankHolidays").css('opacity', '0.5');
			$("#label_WorkingDaysOnly").css('opacity', '0.5');
			$("#label_ShowBankHolidays").css('opacity', '0.5');
		}
		else {
			$("#label_IncludeBankHolidays").css('opacity', '1');
			$("#label_WorkingDaysOnly").css('opacity', '1');
			$("#label_ShowBankHolidays").css('opacity', '1');
		}
	}



	$("#workframe").attr("data-framesource", "UTIL_DEF_CALENDARREPORT");


</script>
