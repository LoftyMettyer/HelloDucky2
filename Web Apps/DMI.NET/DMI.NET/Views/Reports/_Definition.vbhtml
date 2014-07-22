@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Enums
@Imports DMI.NET.Models
@Imports HR.Intranet.Server.Enums
@Inherits System.Web.Mvc.WebViewPage(Of ReportBaseModel)

@Html.HiddenFor(Function(m) m.ID, New With {.id = "txtReportID"})
@Html.HiddenFor(Function(m) m.ReportType, New With {.id = "txtReportType"})

<fieldset>

	<div class="inner">
		<div class="left">

			@Html.LabelFor(Function(m) m.Name)
			@Html.TextBoxFor(Function(m) m.Name)
			@Html.ValidationMessageFor(Function(m) m.Name)

			<br />
			@Html.LabelFor(Function(m) m.Description)
			@Html.TextBox("description", Model.Description)
			@Html.ValidationMessageFor(Function(m) m.Description)

		</div>

		<div class="right">
			<div class="editor-field-greyed-out">
				Owner: @Html.TextBoxFor(Function(m) m.Owner, New With {.readonly = "true"})
			</div>
			<br />
			Access : @Html.Raw(Html.AccessGrid("GroupAccess", Model.GroupAccess, Nothing))
		</div>
	</div>

</fieldset>

<fieldset>
	<legend>Data :</legend>
	<br />

	<div class="inner">

		<div class="left">
			Base Table :
			<select name="BaseTableID" id="BaseTableID" onchange="requestChangeReportBaseTable(event.target);"></select>
		</div>

		<div class="right">

			Records :
			<br />
			@Html.RadioButton("selectiontype", RecordSelectionType.AllRecords, Model.SelectionType = RecordSelectionType.AllRecords, New With {.onclick = "changeRecordOption('Base','all')"})
			All Records
			<br />

			@Html.RadioButton("selectiontype", RecordSelectionType.Picklist, Model.SelectionType = RecordSelectionType.Picklist, New With {.onclick = "changeRecordOption('Base','picklist')"})
			Picklist
			<input type="hidden" id="txtBasePicklistID" name="picklistID" value="@Model.PicklistID" />
			@Html.TextBoxFor(Function(m) m.PicklistName, New With {.id = "txtBasePicklist", .readonly = "true"})
			@Html.EllipseButton("cmdBasePicklist", "selectRecordOption('base', 'picklist')", Model.SelectionType = RecordSelectionType.Picklist)
			<br />

			@Html.RadioButton("selectiontype", RecordSelectionType.Filter, Model.SelectionType = RecordSelectionType.Filter, New With {.onclick = "changeRecordOption('Base','filter')"})
			Filter
			<input type="hidden" id="txtBaseFilterID" name="filterID" value="@Model.FilterID" />
			@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtBaseFilter", .readonly = "true"})
			@Html.EllipseButton("cmdBaseFilter", "selectRecordOption('base', 'filter')", Model.SelectionType = RecordSelectionType.Filter)
			<br />

			@Html.CheckBoxFor(Function(m) m.DisplayTitleInReportHeader)
			@Html.LabelFor(Function(m) m.DisplayTitleInReportHeader)
			<br />

			@Html.ValidationMessageFor(Function(m) m.PicklistID)
			@Html.ValidationMessageFor(Function(m) m.FilterID)

		</div>

		<input type="hidden" id="ctl_DefinitionChanged" name="HasChanged" value="false" />

		<input type="hidden" id="baseHidden" name="baseHidden">

	</div>

</fieldset>

<script type="text/javascript">

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

  function selectRecordOption(psTable, psType) {

  	var sURL;
  	var frmRecordSelection = $("#frmRecordSelection")[0];
  	var iCurrentID;
  	var iTableID;
  	var dropTable;

    if (psTable == 'base') {

    	dropTable = $("#BaseTableID")[0];
    	iTableID = dropTable.options[dropTable.selectedIndex].value;

      if (psType == 'picklist') {
      	iCurrentID = $("#txtBasePicklistID").val();
      }
      else {
        iCurrentID = $("#txtBaseFilterID").val();
      }
    }
    if (psTable == 'p1') {
      iTableID = $("#txtParent1ID").val();

      if (psType == 'picklist') {
        iCurrentID = $("#txtParent1PicklistID").val();
      }
      else {
        iCurrentID = $("#txtParent1FilterID").val();
      }
    }
    if (psTable == 'p2') {
      iTableID = $("#txtParent2ID").val();

      if (psType == 'picklist') {
        iCurrentID = $("#txtParent2PicklistID").val();
      }
      else {
        iCurrentID = $("#txtParent2FilterID").val();
      }
    }

    if (psTable == 'child') {
    	dropTable = $("#ChildTableID")[0];
    	iTableID = dropTable.options[dropTable.selectedIndex].value;
    	iCurrentID = $("txtChildFilterID").val();
    }

    if (psTable == 'event') {
    	dropTable = $("#EventTableID")[0];
    	iTableID = dropTable.options[dropTable.selectedIndex].value;
    	iCurrentID = $("#txtEventFilterID").val();
    }


    var strDefOwner = $("#Owner").val();
    var strCurrentUser = $("#Owner").val();
    var isOwner;

    strDefOwner = strDefOwner.toLowerCase();
    strCurrentUser = strCurrentUser.toLowerCase();

    if (strDefOwner == strCurrentUser) {
    	isOwner = '1';
    }
    else {
    	isOwner = '0';
    }

    sURL = "util_recordSelection" +
			"?recSelType=" + psType +
				"&recSelTableID=" + iTableID +
					"&recSelCurrentID=" + iCurrentID +
						"&recSelTable=" + psTable +
							"&recSelDefOwner=" + isOwner +
								"&recSelDefType=" + escape("Selection");
    openDialog(sURL, (screen.width) / 3 + 40, (screen.height) / 2, "no", "no");

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

		if (tableCount > 0 || columnCount > 0) {
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
		OpenHR.RemoveAllRowsFromGrid(SortOrders, 'Reports/RemoveSortOrder');

		if ($("#ReportType").val() == '@CInt(UtilityType.utlCustomReport)') {
			getAvailableTableColumnsCalcs();
			removeAllChildTables();
			loadAvailableTablesForReport();
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
			}) }
		else {
			menu_loadDefSelPage('@CInt(Model.ReportType)', '@Model.ID', $("#BaseTableID option:selected").val(), true);
		}

		return false;
	}

	$(function () {

		// tighten up these input selectors?
		$("#frmReportDefintion :input").on("change", function () { enableSaveButton(this); });
		getBaseTableList();

	});

</script>

