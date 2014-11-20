﻿@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Models
@Inherits System.Web.Mvc.WebViewPage(Of ReportBaseModel)


@Html.HiddenFor(Function(m) m.ID, New With {.id = "txtReportID"})
@Html.HiddenFor(Function(m) m.ReportType, New With {.id = "txtReportType"})
@Html.HiddenFor(Function(m) m.Timestamp)
@Html.HiddenFor(Function(m) m.ValidityStatus)
@Html.HiddenFor(Function(m) m.BaseViewAccess)
@Html.HiddenFor(Function(m) m.IsReadOnly)
@Html.HiddenFor(Function(m) m.ActionType)

<div class="width100">
	<fieldset class="floatleft width50 bordered">
		<legend class="fontsmalltitle">Identification :</legend>

		<fieldset class="">
			@Html.LabelFor(Function(m) m.Name)
			<div class="width70 floatright">
				@Html.TextBoxFor(Function(m) m.Name, New With {.class = "width100 floatright", .maxlength = 50})
				@Html.ValidationMessageFor(Function(m) m.Name)
			</div>
		</fieldset>

		<fieldset class="">
			@Html.LabelFor(Function(m) m.Description)
			<div id="textareadescription" class="width70 floatright">
				@Html.TextArea("description", Model.Description, New With {.class = "width100 floatright"})
				@Html.ValidationMessageFor(Function(m) m.Description)
			</div>
		</fieldset>
	</fieldset>


	<fieldset id="DataRecordsPermissions" class="floatleft overflowhidden width50">
		<legend class="fontsmalltitle">Data :</legend>
		<div class="inner">
			<fieldset class="">
				Base Table :
				<select class="width70 floatright" name="BaseTableID" id="BaseTableID" onchange="requestChangeReportBaseTable(event.target);"></select>
				<input type="hidden" id="OriginalBaseTableID" />
			</fieldset>

			<div>
				<fieldset id="selectiontypeallrecords" class="">
					@Html.RadioButton("selectiontype", RecordSelectionType.AllRecords, Model.SelectionType = RecordSelectionType.AllRecords,
																New With {.id = "selectiontype_All", .onclick = "changeRecordOption('Base','ALL')"})<span> All Records</span>
				</fieldset>

				<fieldset id="selectiontypepicklistgroup" class="">
					<div id="PicklistRadioDiv" class="floatleft">
						@Html.RadioButton("selectiontype", RecordSelectionType.Picklist, Model.SelectionType = RecordSelectionType.Picklist,
																	New With {.id = "selectiontype_Picklist", .onclick = "changeRecordOption('Base','PICKLIST')"})
						<span>Picklist</span>
					</div>
					<div class="width70 floatright">
						@Html.EllipseButton("cmdBasePicklist", "selectBaseTablePicklist()", Model.SelectionType = RecordSelectionType.Picklist)
						<div class="ellipsistextbox">
							@Html.TextBoxFor(Function(m) m.PicklistName, New With {.id = "txtBasePicklist", .readonly = "true"})
						</div>
					</div>
					<input type="hidden" id="txtBasePicklistID" name="picklistID" value="@Model.PicklistID" />
				</fieldset>

				<fieldset id="selectiontypefiltergroup" class="">
					<div id="FilterRadioDiv" class="floatleft">
						@Html.RadioButton("selectiontype", RecordSelectionType.Filter, Model.SelectionType = RecordSelectionType.Filter,
																	New With {.id = "selectiontype_Filter", .onclick = "changeRecordOption('Base','FILTER')"})
						<span>Filter</span>
					</div>
					<div class="width70  floatright">
						@Html.EllipseButton("cmdBaseFilter", "selectBaseTableFilter()", Model.SelectionType = RecordSelectionType.Filter)
						<div class="ellipsistextbox">
							@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtBaseFilter", .readonly = "true", .class = "width80"})
						</div>
					</div>
					<input type="hidden" id="txtBaseFilterID" name="filterID" value="@Model.FilterID" />
					@Html.ValidationMessageFor(Function(m) m.PicklistID)
					@Html.ValidationMessageFor(Function(m) m.FilterID)
				</fieldset>

				<fieldset>
					<div class="width100  height25 displayTitleInReportHeader" style="display:block">
						@Html.CheckBoxFor(Function(m) m.DisplayTitleInReportHeader)
						@Html.LabelFor(Function(m) m.DisplayTitleInReportHeader, New With {.id = "label_DisplayTitleInReportHeader"})
					</div>
				</fieldset>
			</div>

			<input type="hidden" id="ctl_DefinitionChanged" name="HasChanged" value="false" />
			<input type="hidden" id="baseHidden" name="baseHidden">
		</div>
	</fieldset>

	<fieldset id="AccessPermissions" class="table">
		<legend class="fontsmalltitle">Group Access :</legend>

		<fieldset>
			<div class="nowrap tablelayout" style="margin-top: 0;">
				<div class="tablerow">
					<label>Owner :</label>
					@Html.TextBoxFor(Function(m) m.Owner, New With {.readonly = "true"})
				</div>
				<br />
				<div class="tablerow">
					<label>Access :</label>
					@Html.AccessGrid("GroupAccess", Model.GroupAccess, New With {.id = "tblGroupAccess"})
					<input type="hidden" id="IsForcedHidden" />
				</div>
			</div>

		</fieldset>
	</fieldset>
</div>

<script type="text/javascript">

	$(function () {

		if ($("#txtReportType").val() == '@UtilityType.utlMailMerge') {
			$(".displayTitleInReportHeader").hide();
		}

		$('fieldset').css("border", "0");
		$('table').css("border", "0");

		getBaseTableList();
		refreshViewAccess();

		tableToGrid('#tblGroupAccess', { autoWidth: true, height: 150, cmTemplate: { sortable: false } });

		if ($('#selectiontype_All').prop('checked')) {
			$('#DisplayTitleInReportHeader').prop('disabled', true);
			$("#label_DisplayTitleInReportHeader").css("color", "#A59393");
		}

		menu_toolbarEnableItem('mnutoolSaveReport', false);
		if ($("#ActionType").val() == '@UtilityActionType.Copy') {
			enableSaveButton()
		}

		if (isDefinitionReadOnly())
		{
			$("#frmReportDefintion input").prop('disabled', "disabled");
			$("#frmReportDefintion textarea").prop('disabled', "disabled");
			$("#frmReportDefintion select").prop('disabled', "disabled");
			$("#frmReportDefintion :button").prop('disabled', "disabled");
		}
		else
		{
			$("#frmReportDefintion input").on("keydown", function () { enableSaveButton(); });
			$("#frmReportDefintion textarea").on("keydown", function () { enableSaveButton(); });
			$("#frmReportDefintion input").on("change", function () { enableSaveButton(); });

			//bind click event on the css class for the button and change event for the dropdown to enable the save button
			$("#frmReportDefintion .enableSaveButtonOnClick").on("click", function () { enableSaveButton(); });
			$("#frmReportDefintion .enableSaveButtonOnComboChange").on("change", function () { enableSaveButton(); });
		}
	});

	function isDefinitionReadOnly() {
		return ($("#IsReadOnly").val() == "True");
	}

	function getBaseTableList() {

		$.ajax({
			url: '@Url.Action("GetBaseTables", "Reports", New With {.ReportType = CInt(Model.ReportType)})',
			type: 'GET',
			dataType: 'json',
			success: function (json) {
				$.each(json, function (i, table) {
					var optionHtml = '<option value=' + table.id + '>' + table.Name + '</option>'
					$('#BaseTableID').append(optionHtml);
				});

				$('#BaseTableID').val("@Model.BaseTableID");
				$("#OriginalBaseTableID").val($('#BaseTableID')[0].selectedIndex);

				if ('@CInt(Model.ReportType)' == '2' || '@CInt(Model.ReportType)' == '9') {
					loadAvailableTablesForReport(false);
					attachGridToSelectedColumns();
				}

			}
		});
	}

	function setAllSecurityGroups() {

		var setTo = $("#drpSetAllSecurityGroups").val();
		if (setTo.length > 0) $(".reportViewAccessGroup").val(setTo);

	}

	function changeRecordOption(psTable, psType) {		
		if (psType == "ALL") {
			button_disable($("#cmd" + psTable + "Picklist")[0], true);
			button_disable($("#cmd" + psTable + "Filter")[0], true);
			$("#txt" + psTable + "Filter").val("");
			$("#txt" + psTable + "Picklist").val("");
			$("#txt" + psTable + "PicklistID").val(0);
			$("#txt" + psTable + "FilterID").val(0);
			$('#DisplayTitleInReportHeader').prop('disabled', true);
			$("#DisplayTitleInReportHeader").prop('checked', false);
			$("#label_DisplayTitleInReportHeader").css("color", "#A59393");
		}
		else {
			$('#DisplayTitleInReportHeader').prop('disabled', false);
			$("#label_DisplayTitleInReportHeader").css("color", "#000000");
		}

		if (psType == "PICKLIST") {
			button_disable($("#cmd" + psTable + "Picklist")[0], false)
			button_disable($("#cmd" + psTable + "Filter")[0], true)
			$("#txt" + psTable + "Filter").val("");
			$("#txt" + psTable + "FilterID").val(0);

			if ($("#txt" + psTable + "PicklistID").val() == 0) {
				$("#txt" + psTable + "Picklist").val("None");
			}

		}

		if (psType == "FILTER") {
			button_disable($("#cmd" + psTable + "Picklist")[0], true)
			button_disable($("#cmd" + psTable + "Filter")[0], false)
			$("#txt" + psTable + "Picklist").val("");
			$("#txt" + psTable + "PicklistID").val(0);

			if ($("#txt" + psTable + "FilterID").val() == 0) {
				$("#txt" + psTable + "Filter").val("None");
			}

		}

		setViewAccess(psType, $("#" + psTable + "ViewAccess"), 'RW', '');

	}

		function refreshViewAccess() {

		var bViewAccessEnabled = true;
		var list;

		$("#tblGroupAccess").removeAttr("disabled");

		if ($("#BaseViewAccess").val() == 'HD') { bViewAccessEnabled = false; }
		if ($("#Parent1ViewAccess").val() == 'HD') { bViewAccessEnabled = false; }
		if ($("#Parent2ViewAccess").val() == 'HD') { bViewAccessEnabled = false; }
		if ($("#StartCustomViewAccess").val() == 'HD') { bViewAccessEnabled = false; }
		if ($("#EndCustomViewAccess").val() == 'HD') { bViewAccessEnabled = false; }
		if ($("#Description3ViewAccess").val() == 'HD') { bViewAccessEnabled = false; }

		list = $("#ChildTables").getDataIDs();
		for (i = 0; i < list.length; i++) {
			rowData = $("#ChildTables").getRowData(list[i]);
			if (rowData.FilterViewAccess == "HD") { bViewAccessEnabled = false; }
		}

		list = $("#CalendarEvents").getDataIDs();
		for (i = 0; i < list.length; i++) {
			rowData = $("#CalendarEvents").getRowData(list[i]);
			if (rowData.FilterViewAccess == "HD") { bViewAccessEnabled = false; }
		}

		if (!bViewAccessEnabled) {
			$("#IsForcedHidden").val(true);
			$("#tblGroupAccess").attr("disabled", "disabled");
		}
	}

	function setViewAccess(type, accessControl, newAccess, tableName) {

		var bResetGroupsToHidden = false;
		var iObjectsHidden = 0;
		var displayType;

		if (accessControl.val() != newAccess && newAccess == "HD") {
			bResetGroupsToHidden = true;
		}

		switch (type) {
			case "FILTER":
				displayType = "filter";
				break;

			case "PICKLIST":
				displayType = "picklist";
				break;

			default:
				displayType = "calculation";

		}

		accessControl.val(newAccess);

		if (bResetGroupsToHidden && $("#IsForcedHidden").val() != "true") {
			OpenHR.modalPrompt("This definition will now be made hidden as the " + tableName + " table " + displayType + " is hidden.", 0, "Information").then(function (answer) {
				$(".reportViewAccessGroup").val("HD");
				$("#IsForcedHidden").val(true);
			});
		}
		else {

			$("[id$='ViewAccess']").each(function (index) {
				if ((this).value == "HD") {
					iObjectsHidden += 1;
				}
			});

			if (iObjectsHidden == 0 && $("#IsForcedHidden").val() == "true") {
				OpenHR.modalPrompt("This definition no longer has to be hidden.", 0, "Information").then(function (answer) {
					$("#IsForcedHidden").val(false);
				});
			}

		}

		refreshViewAccess();

	}

	function selectBaseTableFilter() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#txtBaseFilterID").val();
		var tableName = $("#BaseTableID option:selected").text();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name, access) {

			if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower') {
				$("#txtBaseFilterID").val(0);
				$("#txtBaseFilter").val('None');
				OpenHR.modalMessage("The " + tableName + " table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#txtBaseFilterID").val(id);
				$("#txtBaseFilter").val(name);
				setViewAccess('FILTER', $("#BaseViewAccess"), access, tableName);
				enableSaveButton();
			}
		}, 400, 400)
	}

	function selectBaseTablePicklist() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#txtBasePicklistID").val();
		var tableName = $("#BaseTableID option:selected").text();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name, access) {
			if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower') {
				$("#txtBasePicklistID").val(0);
				$("#txtBasePicklist").val('None');
				OpenHR.modalMessage("The " + tableName + " table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#txtBasePicklistID").val(id);
				$("#txtBasePicklist").val(name);
				setViewAccess('PICKLIST', $("#BaseViewAccess"), access, tableName);
				enableSaveButton();
			}

		}, 400, 400);

	}

	function resetParentDetails() {

		$("#RelatedTableParent1").attr("disabled", "disabled");
		disableParent1RadioButtons();
		$("#Parent1_SelectionTypeAll").prop('checked', 'checked');
		changeRecordOption('Parent1', 'ALL');

		$("#RelatedTableParent2").attr("disabled", "disabled");
		disableParent2RadioButtons();
		$("#Parent2_SelectionTypeAll").prop('checked', 'checked');
		changeRecordOption('Parent2', 'ALL');

	}


	function loadAvailableTablesForCalendarReport(baseTableChanged) {

		$.ajax({
			url: 'Reports/GetAvailableColumnsForTable?TableID=' + $("#BaseTableID").val(),
			type: 'GET',
			dataType: 'json',
			cache: false,
			success: function (json) {
				// Clear the Parent1 and Parent2 table names only if the base table changed
				if (baseTableChanged) {
					$("#Description1ID").val(0);
					$("#Description2ID").val(0);
					$("#Description3ID").val(0);
					$("#txtDescription3").val('');
					$('#RegionID').val(0);
					$('#Separator').val('None');
					$('#chkGroupByDescription').prop('checked', false);

					$('#IncludeBankHolidays').prop('checked', false);
					$('#WorkingDaysOnly').prop('checked', false);
					$('#ShowBankHolidays').prop('checked', false);
					$('#ShowCaptions').prop('checked', 'checked');
					$('#ShowWeekends').prop('checked', 'checked');
					$('#StartOnCurrentMonth').prop('checked', 'checked');
				}
				var OptionNone = '<option value=0 data-datatype=0 data-decimals=0 selected>None</option>';
				var optionDescription1 = "<option value='0'>None</option>";
				var optionDescription2 = "<option value='0'>None</option>";
				var optionDescription3 = "<option value='0'>None</option>";
				var optionRegionID = "<option value='0'>None</option>";

				var options = '';
				for (var i = 0; i < json.length; i++) {
					optionDescription1 += "<option value='" + json[i].ID + "'>" + json[i].Name + "</option>";
					optionDescription2 += "<option value='" + json[i].ID + "'>" + json[i].Name + "</option>";
					optionDescription3 += "<option value='" + json[i].ID + "'>" + json[i].Name + "</option>";
					optionRegionID += "<option value='" + json[i].ID + "'>" + json[i].Name + "</option>";
				}
				$("select#Description1ID").html(optionDescription1);
				$("select#Description2ID").html(optionDescription2);
				$("select#Description3ID").html(optionDescription3);
				$("select#RegionID").html(optionRegionID);
			}
		});
	}

	function loadAvailableTablesForReport(baseTableChanged) {

		$.ajax({
			url: '@Html.Raw(Url.Action("GetAllTablesInReport", "Reports", New With {.ReportID = Model.ID, .ReportType = CInt(Model.ReportType)}))',
			type: 'GET',
			dataType: 'json',
			cache: false,
			success: function (json) {

				$('#SelectedTableID').empty()

				// Clear the Parent1 and Parent2 table names only if the base table changed
				if (baseTableChanged) {
					$("#txtParent1ID").val(0);
					$("#txtParent2ID").val(0);
					$("#Parent1_Name").val("");
					$("#Parent2_Name").val("");
				}

				$.each(json, function (i, table) {
					var optionHtml = '<option value=' + table.id + '>' + table.Name + '</option>'
					$('#SelectedTableID').append(optionHtml);

					if (table.Relation == 1 && baseTableChanged) {
						$("#RelatedTableParent1").removeAttr("disabled");
						enableParent1RadioButtons();
						$("#txtParent1ID").val(table.id);
						$("#Parent1_Name").val(table.Name);
					}

					if (table.Relation == 2 && baseTableChanged) {
						$("#RelatedTableParent2").removeAttr("disabled");
						enableParent2RadioButtons();
						$("#txtParent2ID").val(table.id);
						$("#Parent2_Name").val(table.Name);
					}

				});

				$("#SelectedTableID").val($("#BaseTableID").val());
				getAvailableTableColumnsCalcs();

			}
		});
	}

	function requestChangeReportBaseTable(target) {
		
		var tableCount = $("#ChildTables").getGridParam("reccount");
		var columnCount = $("#SelectedColumns").getGridParam("reccount");
		var eventCount = $("#CalendarEvents").getGridParam("reccount");
		var sortOrderCount = $("#SortOrders").getGridParam("reccount");

		if (tableCount > 0 || columnCount > 0 || eventCount > 0 || sortOrderCount > 0) {
			OpenHR.modalPrompt("Changing the base table will result in all table/column specific aspects of this definition being cleared. <br/><br/>Are you sure you wish to continue ?", 4, "").then(function (answer) {
				if (answer == 6) { // Yes
					changeReportBaseTable();
				}
				else {
					$('#BaseTableID')[0].selectedIndex = $("#OriginalBaseTableID").val();
				}
			});
		}
		else {
			changeReportBaseTable();
		}
	}

	function changeReportBaseTable() {
		// Post base table change to server
		var dataSend = {
			ReportID: '@Model.ID',
			ReportType: '@Model.ReportType',
			BaseTableID: $("#BaseTableID option:selected").val()
		};

		OpenHR.postData("Reports/ChangeBaseTable", dataSend, changeReportBaseTableCompleted);
		$("#OriginalBaseTableID").val($('#BaseTableID')[0].selectedIndex);

	}

	function changeReportBaseTableCompleted(json) {

		$("#selectiontype_All").prop('checked', 'checked');
		$("#ChildTablesAvailable").val(parseInt(json.childTablesAvailable));
		changeRecordOption('Base', 'ALL');

		if ($("#txtReportType").val() != '@UtilityType.utlCrossTab' && $("#txtReportType").val() != '@UtilityType.utlNineBoxGrid') {
			removeAllSortOrders();
		}

		if ($("#txtReportType").val() == '@UtilityType.utlCustomReport') {
			removeAllChildTables(false);
			resetParentDetails();
		}

		if ($("#txtReportType").val() == '@UtilityType.utlCustomReport' || $("#txtReportType").val() == '@UtilityType.utlMailMerge') {			
			removeAllSelectedColumns(false);
			loadAvailableTablesForReport(true);
		}

		if ($("#txtReportType").val() == '@UtilityType.utlCalendarReport') {
			button_disable($("#btnSortOrderAdd")[0], false);
			$("#SortOrdersAvailable").val(parseInt(json.sortOrdersAvailable));
			$('#CalendarEvents').jqGrid('clearGridData');
			loadAvailableTablesForCalendarReport(true);
		}

		if ($("#txtReportType").val() == '@UtilityType.utlCrossTab' || $("#txtReportType").val() == '@UtilityType.utlNineBoxGrid') {
			refreshCrossTabColumnsAvailable();
		}

	}

	function removeAllChildTablesCompleted() {

		var childTables = $("#ChildTables").getDataIDs();

		for (i = 0; i < childTables.length; i++) {
			thisTable = $("#ChildTables").getRowData(childTables[i]);

			var columnList = $("#SelectedColumns").getDataIDs();
			for (j = 0; j < columnList.length; j++) {
				rowData = $("#SelectedColumns").getRowData(columnList[j]);

				if (rowData.TableID == thisTable.TableID) {
					$('#SelectedColumns').jqGrid('delRowData', rowData.ID);
				}

			}

		}

		$('#ChildTables').jqGrid('clearGridData');

		enableSaveButton();

	}

	function removeAllChildTables() {

		var data = { ReportID: "@Model.ID", ReportType: "@Model.ReportType" }
		OpenHR.postData("Reports/RemoveAllChildTables", data, removeAllChildTablesCompleted);
	}

	function enableSaveButton() {

		if (!isDefinitionReadOnly()) {
			$("#ctl_DefinitionChanged").val("true");
			menu_toolbarEnableItem('mnutoolSaveReport', true);
		}
	}

	function saveReportDefinition(prompt) {

		var bHasChanged = $("#ctl_DefinitionChanged").val();

		if (prompt == true) {
			if (bHasChanged == "true") {

				OpenHR.modalPrompt("You have made changes. Click 'OK' to discard your changes, or 'Cancel' to continue editing.", 1, "Confirm").then(function (answer) {
					if (answer == 1) {
						validateReportDefinition();
					}
				});
			}
			else {
				return 6;
			}

		} else {
			validateReportDefinition()
		}

		return 0;

	}

	function validateReportDefinition() {

		var gridData;

		// Columns selected
		gridData = $("#SelectedColumns").getRowData();
		$('#txtCSAAS').val(JSON.stringify(gridData));

		// Related Tables
		gridData = $("#ChildTables").getRowData();
		$('#txtCTAAS').val(JSON.stringify(gridData));

		// Calendar Events
		gridData = $("#CalendarEvents").getRowData();
		$('#txtCEAAS').val(JSON.stringify(gridData));

		// Sort Order columns
		gridData = $("#SortOrders").getRowData();
		$('#txtSOAAS').val(JSON.stringify(gridData));

		var $form = $("#frmReportDefintion");
		$("#tblGroupAccess").removeAttr("disabled");

		$.ajax({
			url: $form.attr("action"),
			type: $form.attr("method"),
			data: $form.serialize(),
			async: true,
			success: function (json) {

				switch (json.ErrorCode) {
					case 0:
						submitReportDefinition();
						break;

					case 1:
						OpenHR.modalPrompt(json.ErrorMessage, 0, "OpenHR");
						break;

					case -1:
						OpenHR.modalPrompt(json.ErrorMessage, 0, "OpenHR");
						break;

					default:
						OpenHR.modalPrompt(json.ErrorMessage, 4, "Confirm").then(function (answer) {
							if (answer == 6) {
								submitReportDefinition();
							}
						});
						break;

				}
				refreshViewAccess();
			}
		});
	}

	function submitReportDefinition() {
		$("#ValidityStatus").val('ServerCheckComplete');
		$("#tblGroupAccess").removeAttr("disabled");
		var frmSubmit = $("#frmReportDefintion")[0];
		OpenHR.submitForm(frmSubmit);
	}

	function cancelReportDefinition() {

		var bHasChanged = $("#ctl_DefinitionChanged").val();

		if (bHasChanged == "true") {
			OpenHR.modalPrompt("You have made changes. Click 'OK' to discard your changes, or 'Cancel' to continue editing.", 1, "Confirm").then(function (answer) {
				if (answer == 1) {  // OK
					menu_loadDefSelPage('@CInt(Model.ReportType)', '@Model.ID', $("#BaseTableID option:selected").val(), true);
					return 6;
				}
			})
		}
		else {

			menu_loadDefSelPage('@CInt(Model.ReportType)', '@Session("utilid")', $("#BaseTableID option:selected").val(), true);
		}

		return false;
	}



</script>