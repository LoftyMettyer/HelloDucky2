@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Models
@Inherits System.Web.Mvc.WebViewPage(Of ReportBaseModel)


@Html.HiddenFor(Function(m) m.ID, New With {.id = "txtReportID"})
@Html.HiddenFor(Function(m) m.ReportType, New With {.id = "txtReportType"})
@Html.HiddenFor(Function(m) m.Timestamp)
@Html.HiddenFor(Function(m) m.ValidityStatus)
@Html.HiddenFor(Function(m) m.BaseViewAccess)
@Html.HiddenFor(Function(m) m.IsReadOnly)


<div class="width100">
	<fieldset class="floatleft width50 bordered">
		<legend class="fontsmalltitle">Name :</legend>
			<fieldset>
				<div id="DescriptionItems">
					@Html.LabelFor(Function(m) m.Name)
					@Html.TextBoxFor(Function(m) m.Name, New With {.class = "width70 floatright"})
					@Html.ValidationMessageFor(Function(m) m.Name)
				</div>
			<br />
				<div>
					@Html.LabelFor(Function(m) m.Description)
					@Html.TextArea("description", Model.Description, New With {.class = "width70 floatright"})
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
																New With {.id = "selectiontype_All", .onclick = "changeRecordOption('Base','ALL')"})All Records
					</fieldset>

					<fieldset id="selectiontypepicklistgroup" class="width100">
						<div id="PicklistRadioDiv" class="floatleft">
							@Html.RadioButton("selectiontype", RecordSelectionType.Picklist, Model.SelectionType = RecordSelectionType.Picklist,
																	New With {.id = "selectiontype_Picklist", .onclick = "changeRecordOption('Base','PICKLIST')"})
							<span>Picklist</span>
						</div>
						<div class="width70 floatleft">
							@Html.EllipseButton("cmdBasePicklist", "selectBaseTablePicklist()", Model.SelectionType = RecordSelectionType.Picklist)
							<div class="ellipsistextbox">
								@Html.TextBoxFor(Function(m) m.PicklistName, New With {.id = "txtBasePicklist", .readonly = "true"})
							</div>
						</div>
						<input type="hidden" id="txtBasePicklistID" name="picklistID" value="@Model.PicklistID" />
					</fieldset>

					<fieldset id="selectiontypefiltergroup" class="width100">
						<div id="FilterRadioDiv" class="floatleft">
							@Html.RadioButton("selectiontype", RecordSelectionType.Filter, Model.SelectionType = RecordSelectionType.Filter,
																	New With {.id = "selectiontype_Filter", .onclick = "changeRecordOption('Base','FILTER')"})
							<span>Filter</span>
						</div>
						<div class="width70  floatleft">
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
						<div class="width100  height25" style="display:block">
							@Html.CheckBoxFor(Function(m) m.DisplayTitleInReportHeader)
							@Html.LabelFor(Function(m) m.DisplayTitleInReportHeader)
						</div>
					</fieldset>
				</div>

				<input type="hidden" id="ctl_DefinitionChanged" name="HasChanged" value="false" />
				<input type="hidden" id="baseHidden" name="baseHidden">
			</div>
		</fieldset>

	<fieldset id="AccessPermissions" class="table">
			<legend class="fontsmalltitle">Access :</legend>

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
				</div>
			</div>

			<br />
			<label id="ForcedHiddenMessage" hidden="hidden">The definition access cannot be changed as it contains a hidden picklist, filter or calculation</label>
		</fieldset>
	</fieldset>
</div>

<script type="text/javascript">

	$(function () {

		$('fieldset').css("border", "0");
		$('table').css("border", "0");
	
		getBaseTableList();
		refreshViewAccess();

		tableToGrid('#tblGroupAccess', { autoWidth: true, height: 150, cmTemplate: { sortable: false } });

		if ($('#selectiontype_All').prop('checked')) $('#DisplayTitleInReportHeader').prop('disabled', true);
		menu_toolbarEnableItem('mnutoolSaveReport', false);

		if (isDefinitionReadOnly()) {
			$("#frmReportDefintion input").prop('disabled', "disabled");
			$("#frmReportDefintion textarea").prop('disabled', "disabled");
			$("#frmReportDefintion select").prop('disabled', "disabled");
			$("#frmReportDefintion :button").prop('disabled', "disabled");
		} else {
			$("#frmReportDefintion input").on("keydown", function () { enableSaveButton(); });
			$("#frmReportDefintion textarea").on("keydown", function () { enableSaveButton(); });
			$("#frmReportDefintion input").on("change", function () { enableSaveButton(); });
			$("#frmReportDefintion select").on("change", function () { enableSaveButton(); });
			$("#frmReportDefintion :button").on("click", function () { enableSaveButton(); });
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
		}
		else {
			$('#DisplayTitleInReportHeader').prop('disabled', false);
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

		$("#AccessPermissionsGrid").removeAttr("disabled");
		$("#ForcedHiddenMessage").hide();

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
			$("#AccessPermissionsGrid").attr("disabled", "disabled");
			$("#ForcedHiddenMessage").show();
		}

	}

	function setViewAccess(type, accessControl, newAccess, tableName) {

		var bResetGroupsToHidden = false;
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

		if (bResetGroupsToHidden) {
			OpenHR.modalPrompt("This definition will now be made hidden as the " + tableName + " table " + displayType + " is hidden.", 0, "Information").then(function (answer) {
				$(".reportViewAccessGroup").val("HD");
			});

		}

		accessControl.val(newAccess);

		refreshViewAccess();

	}

	function selectBaseTableFilter() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#txtBaseFilterID").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name, access) {
			$("#txtBaseFilterID").val(id);
			$("#txtBaseFilter").val(name);
			setViewAccess('FILTER', $("#BaseViewAccess"), access, $("#BaseTableID option:selected").text());
		}, 400, 400);

	}

	function selectBaseTablePicklist() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#txtBasePicklistID").val();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name, access) {
			$("#txtBasePicklistID").val(id);
			$("#txtBasePicklist").val(name);
			setViewAccess('PICKLIST', $("#BaseViewAccess"), access, $("#BaseTableID option:selected").text());
		}, 400, 400);

	}

	function resetParentDetails() {

		$("#RelatedTableParent1").attr("disabled", "disabled");
		$("#Parent1_SelectionTypeAll").prop('checked', 'checked');
		changeRecordOption('Parent1', 'ALL');

		$("#RelatedTableParent2").attr("disabled", "disabled");
		$("#Parent2_SelectionTypeAll").prop('checked', 'checked');
		changeRecordOption('Parent2', 'ALL');

	}

	function loadAvailableTablesForReport(baseTableChanged) {

		$.ajax({
			url: '@Html.Raw(Url.Action("GetAllTablesInReport", "Reports", New With {.ReportID = Model.ID, .ReportType = CInt(Model.ReportType)}))',
			type: 'GET',
			dataType: 'json',
			cache: false,
			success: function (json) {

				$('#SelectedTableID').empty()

				$.each(json, function (i, table) {
					var optionHtml = '<option value=' + table.id + '>' + table.Name + '</option>'
					$('#SelectedTableID').append(optionHtml);

					if (table.Relation == 1 && baseTableChanged) {
						$("#RelatedTableParent1").removeAttr("disabled");
						$("#txtParent1ID").val(table.id);
						$("#Parent1_Name").val(table.Name);
					}

					if (table.Relation == 2 && baseTableChanged) {
						$("#RelatedTableParent2").removeAttr("disabled");
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
			OpenHR.modalPrompt("Changing the base table will result in all table/column specific aspects of this report definition being cleared. <br/><br/>Are you sure you wish to continue ?", 4, "").then(function (answer) {
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

		if ($("#txtReportType").val() != '@UtilityType.utlCrossTab') {
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
			$('#CalendarEvents').jqGrid('clearGridData');
		}

		if ($("#txtReportType").val() == '@UtilityType.utlCrossTab') {
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
		$("#AccessPermissionsGrid").removeAttr("disabled");

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
		$("#AccessPermissionsGrid").removeAttr("disabled");
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
			menu_loadDefSelPage('@CInt(Model.ReportType)', '@Model.ID', $("#BaseTableID option:selected").val(), true);
		}

		return false;
	}



</script>

