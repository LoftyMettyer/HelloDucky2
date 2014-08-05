@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Models

@Inherits System.Web.Mvc.WebViewPage(Of ReportBaseModel)


@Html.HiddenFor(Function(m) m.ID, New With {.id = "txtReportID"})
@Html.HiddenFor(Function(m) m.ReportType, New With {.id = "txtReportType"})

<div class="width100">
	<fieldset class="floatleft width50 ">
		<fieldset class="floatleft width99 bordered">
			<legend class="fontsmalltitle">Description :</legend>
			<fieldset>
				<div id="DescriptionItems">
					@Html.LabelFor(Function(m) m.Name)
					@Html.TextBoxFor(Function(m) m.Name, New With {.class = "width70 floatright"})
					@Html.ValidationMessageFor(Function(m) m.Name)
				</div>

				<div>
					@Html.LabelFor(Function(m) m.Description)
					@Html.TextArea("description", Model.Description, New With {.class = "width70 floatright"})
					@Html.ValidationMessageFor(Function(m) m.Description)
				</div>
			</fieldset>
		</fieldset>

		<fieldset id="DataRecordsPermissions" class="overflowhidden width99">
			<legend class="fontsmalltitle">Data :</legend>
			<div class="inner">
				<fieldset class="">
					Base Table :
					<select class="width70 floatright" name="BaseTableID" id="BaseTableID" onchange="requestChangeReportBaseTable(event.target);"></select>
				</fieldset>

				<div>
					<fieldset class="">
						<fieldset id="selectiontypeallrecords" class="">							
							@Html.RadioButton("selectiontype", RecordSelectionType.AllRecords, Model.SelectionType = RecordSelectionType.AllRecords,
																New With {.onclick = "changeRecordOption('Base','all')"})All Records
						</fieldset>

						<fieldset id="selectiontypepicklistgroup" class="width100">
							<div id="PicklistRadioDiv" class="floatleft">
								@Html.RadioButton("selectiontype", RecordSelectionType.Picklist, Model.SelectionType = RecordSelectionType.Picklist,
																	New With {.onclick = "changeRecordOption('Base','picklist')"})
								<span>Picklist</span>
							</div>
							<div class="width70 floatleft">
								@Html.TextBoxFor(Function(m) m.PicklistName, New With {.id = "txtBasePicklist", .readonly = "true", .class = "width80"})
								@Html.EllipseButton("cmdBasePicklist", "selectBaseTablePicklist()", Model.SelectionType = RecordSelectionType.Picklist)
							</div>
							<input type="hidden" id="txtBasePicklistID" name="picklistID" value="@Model.PicklistID" />
						</fieldset>

						<fieldset id="selectiontypefiltergroup" class="width100">
							<div id="FilterRadioDiv" class="floatleft">
								@Html.RadioButton("selectiontype", RecordSelectionType.Filter, Model.SelectionType = RecordSelectionType.Filter, New With {.onclick = "changeRecordOption('Base','filter')"})
								<span>Filter</span>
							</div>
							<div class="width70  floatleft">
								@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtBaseFilter", .readonly = "true", .class = "width80"})
								@Html.EllipseButton("cmdBaseFilter", "selectBaseTableFilter()", Model.SelectionType = RecordSelectionType.Filter)
							</div>
							<input type="hidden" id="txtBaseFilterID" name="filterID" value="@Model.FilterID" />
						</fieldset>
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
	</fieldset>

	<fieldset id="AccessPermissions"  class="width35">
		<fieldset>
			<legend class="fontsmalltitle">Access :</legend>
			<fieldset>
				<span class="floatleft">Owner : </span>
				@Html.TextBoxFor(Function(m) m.Owner, New With {.readonly = "true"})
			</fieldset>
			<fieldset id="AccessPermissionsGrid" >
				@Html.Raw(Html.AccessGrid("GroupAccess", Model.GroupAccess, Nothing))
			</fieldset>
		</fieldset>
	</fieldset>
</div>

<script type="text/javascript">

	$(function () {
		 $('fieldset').css("border", "0");		
		$("#frmReportDefintion :input").on("change", function () { enableSaveButton(this); });
		getBaseTableList();

	});

	function getBaseTableList() {
		$.ajax({
			url: '@Url.Action("GetBaseTables", "Reports")',
			type: 'GET',
			dataType: 'json',
			success: function (json) {
				$.each(json, function (i, table) {
					var optionHtml = '<option value=' + table.id + '>' + table.Name + '</option>'
					$('#BaseTableID').append(optionHtml);
				});

				$('#BaseTableID').val("@Model.BaseTableID");

				if ('@CInt(Model.ReportType)' == '2' || '@CInt(Model.ReportType)' == '9') {
					loadAvailableTablesForReport();
					attachGridToSelectedColumns();
				}

			}
		});
	}

	function changeRecordOption(psTable, psType) {

		if (psType == "all") {
			button_disable($("#cmd" + psTable + "Picklist")[0], true);
			button_disable($("#cmd" + psTable + "Filter")[0], true);
			$("#txt" + psTable + "Filter").val("");
			$("#txt" + psTable + "Picklist").val("");
			$("#txt" + psTable + "PicklistID").val(0);
			$("#txt" + psTable + "FilterID").val(0);
		}

		if (psType == "picklist") {
			button_disable($("#cmd" + psTable + "Picklist")[0], false)
			button_disable($("#cmd" + psTable + "Filter")[0], true)
			$("#txt" + psTable + "Filter").val("");
			$("#txt" + psTable + "FilterID").val(0);

			if ($("#txt" + psTable + "PicklistID").val() == 0) {
				$("#txt" + psTable + "Picklist").val("<None>");
			}

		}

		if (psType == "filter") {
			button_disable($("#cmd" + psTable + "Picklist")[0], true)
			button_disable($("#cmd" + psTable + "Filter")[0], false)
			$("#txt" + psTable + "Picklist").val("");
			$("#txt" + psTable + "PicklistID").val(0);

			if ($("#txt" + psTable + "FilterID").val() == 0) {
				$("#txt" + psTable + "Filter").val("<None>");
			}

		}

	}

	function selectBaseTableFilter() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#txtBaseFilterID").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name) {
			$("#txtBaseFilterID").val(id);
			$("#txtBaseFilter").val(name);
		});

	}

	function selectBaseTablePicklist() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#txtBasePicklistID").val();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name) {
			$("#txtBasePicklistID").val(id);
			$("#txtBasePicklist").val(name);
		});

	}

	function loadAvailableTablesForReport() {

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
			OpenHR.modalPrompt("Changing the base table will result in all table/column specific aspects of this report definition being cleared. <br/><br/> Are you sure you wish to continue ?", 4, "").then(function (answer) {
				if (answer == 6) { // Yes
					changeReportBaseTable();
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

	}

	function changeReportBaseTableCompleted() {

		removeAllSelectedColumns();

		if ($("#txtReportType").val() != '@UtilityType.utlCrossTab') {
			OpenHR.RemoveAllRowsFromGrid(SortOrders, 'Reports/RemoveSortOrder');
		}

		if ($("#txtReportType").val() == '@UtilityType.utlCustomReport') {
			getAvailableTableColumnsCalcs();
			removeAllChildTables();
			loadAvailableTablesForReport();
		}

		if ($("#txtReportType").val() == '@UtilityType.utlCalendarReport') {
			$('#CalendarEvents').jqGrid('clearGridData');
		}


	}

	function removeAllChildTables() {
		$('#ChildTables').jqGrid('clearGridData')
	}

	function removeAllSelectedColumns() {
		$('#SelectedColumns').jqGrid('clearGridData')
	}

	function enableSaveButton() {
		$("#ctl_DefinitionChanged").val("true");
	}

	function saveReportDefinition(prompt) {

		var bHasChanged = $("#ctl_DefinitionChanged").val();

		if (prompt == true) {
			if (bHasChanged == "true") {

				OpenHR.modalPrompt("You have made changes. Click 'OK' to discard your changes, or 'Cancel' to continue editing.", 1, "Confirm").then(function (answer) {
					if (answer == 1) {
						submitReportDefinition();
					}
				});
			}
			else {
				return 6;
			}

		} else {
			submitReportDefinition()
		}

		return 0;

	}

	function submitReportDefinition() {

		// Related Tables
		var gridData = jQuery("#ChildTables").getRowData();
		var postData = JSON.stringify(gridData);
		$('#txtCTAAS').val(postData);

		// Columns selected
		gridData = $("#SelectedColumns").getRowData();
		$('#txtCSAAS').val(JSON.stringify(gridData));

		// Calendar Events
		gridData = $("#CalendarEvents").getRowData();
		$('#txtCEAAS').val(JSON.stringify(gridData));

		// Sort Order columns
		gridData = $("#SortOrders").getRowData();
		$('#txtSOAAS').val(JSON.stringify(gridData));

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

