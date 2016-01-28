@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.TalentReportModel)
@Html.HiddenFor(Function(m) m.MatchViewAccess, New With {.class = "ViewAccess"})
@Code
	Layout = Nothing
End Code

<div>

	<form action="reports\util_def_mailmerge_downloadtemplate" style="display: none" method="post" id="frmDownloadTemplate" name="frmDownloadTemplate" target="submit-iframe">
		<input type="hidden" id="MailMergeId" name="MailMergeId" value="@Model.ID" />
		<input type="hidden" id="download_token_value_id" name="download_token_value_id" />
		@Html.AntiForgeryToken()
	</form>


	@Using (Html.BeginForm("util_def_talentreport", "Reports", FormMethod.Post, New With {.id = "frmReportDefintion", .name = "frmReportDefintion"}))

		@Html.HiddenFor(Function(m) m.ID)

		@<div id="tabs">
			<ul>
				<li><a href="#tabs-1">Definition</a></li>
				<li><a href="#report_definition_tab_Match">Match Tables</a></li>
				<li><a href="#report_definition_tab_columns">Columns</a></li>
				<li><a href="#report_definition_tab_order">Sort Order</a></li>
				<li><a href="#report_definition_tab_output">Output</a></li>
			</ul>

			<div id="tabs-1">
				@Code
				Html.RenderPartial("_Definition", Model)
				End Code

				<fieldset class="floatleft overflowhidden width50">
					<div class="inner">
						<fieldset class="">
							Person Table : <select class="width70 floatright" name="MatchTableID" id="MatchTableID" onchange="requestChangeReportPersonTable(event.target);"></select>							
							<input type="hidden" id="OriginalRoleTableID" />
							<input type="hidden" id="OriginalRoleTableText" />							
							<input type="hidden" id="OriginalPersonTableID" />
							<input type="hidden" id="OriginalPersonTableText" />							
							<input type="hidden" id="IsPersonTableChange" value="False" />
						</fieldset>

						<div>
							<fieldset id="MatchTableAllRecordsDiv" class="">

								@Html.RadioButton("matchselectiontype", RecordSelectionType.AllRecords, Model.MatchSelectionType = RecordSelectionType.AllRecords,
																						New With {.id = "matchselectiontype_All", .onclick = "changeRecordOption('Match','ALL')"})<span> All Records</span>

							</fieldset>

							<fieldset id="matchselectiontypepicklistgroup" class="">
								<div id="MatchPicklistRadioDiv" class="floatleft">
									@Html.RadioButton("matchselectiontype", RecordSelectionType.Picklist, Model.MatchSelectionType = RecordSelectionType.Picklist,
															 New With {.id = "matchselectiontype_Picklist", .onclick = "changeRecordOption('Match','PICKLIST')"})
									<span>Picklist</span>
								</div>
								<div class="width70 floatright">
									@Html.EllipseButton("cmdMatchPicklist", "selectMatchTablePicklist()", Model.SelectionType = RecordSelectionType.Picklist)
									<div class="ellipsistextbox">
										@Html.TextBoxFor(Function(m) m.MatchPicklistName, New With {.id = "txtMatchPicklist", .readonly = "true"})
										@Html.ValidationMessageFor(Function(m) m.MatchPicklistID)
									</div>
								</div>
								<input type="hidden" id="txtMatchPicklistID" name="MatchPicklistID" value="@Model.MatchPicklistID" />
							</fieldset>

							<fieldset id="matchselectiontypefiltergroup" class="">
								<div id="MatchFilterRadioDiv" class="floatleft">
									@Html.RadioButton("matchselectiontype", RecordSelectionType.Filter, Model.MatchSelectionType = RecordSelectionType.Filter,
																	 New With {.id = "matchselectiontype_Filter", .onclick = "changeRecordOption('Match','FILTER')"})
									<span>Filter</span>
								</div>

								<div class="width70 floatright">
									@Html.EllipseButton("cmdMatchFilter", "selectMatchTableFilter()", Model.SelectionType = RecordSelectionType.Filter)
									<div class="ellipsistextbox">
										@Html.TextBoxFor(Function(m) m.MatchFilterName, New With {.id = "txtMatchFilter", .readonly = "true"})
										@Html.ValidationMessageFor(Function(m) m.MatchFilterID)
									</div>
								</div>
								<input type="hidden" id="txtMatchFilterID" name="MatchFilterID" value="@Model.MatchFilterID" />
							</fieldset>

						</div>
					</div>
				</fieldset>

			</div>

			<div id="report_definition_tab_Match">
				@Code
				Html.RenderPartial("_MatchTables", Model)
				End Code
			</div>

			<div id="report_definition_tab_columns">
				@Code
				Html.RenderPartial("_TalentManagementColumnSelection", Model)
				End Code
			</div>

			<div id="report_definition_tab_order">
				@Code
				Html.RenderPartial("_SortOrder", Model)
				End Code
			</div>

			<div id="report_definition_tab_output">
				@Code
				Html.RenderPartial("_OutputTalentReport", Model.Output)
				End Code
			</div>
		</div>
		@Html.AntiForgeryToken()
	End Using

</div>

<script type="text/javascript">

	function selectMatchTablePicklist() {

		var tableID = $("#MatchTableID").val();
		var currentID = $("#txtMatchPicklistID").val();
		var tableName = $("#MatchTableID option:selected").text();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name, access) {
			//If current user is System Manager/Security Manager, we allow them to add or edit the filter/picklist hidden by another user
			if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower' && '@Model.CanEditSecurityGroups.ToString.ToLower' == "false") {
				$("#txtMatchPicklistID").val(0);
				$("#txtMatchPicklist").val('None');
				OpenHR.modalMessage("The " + tableName + " table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#txtMatchPicklistID").val(id);
				$("#txtMatchPicklist").val(name);
				setViewAccess('PICKLIST', $("#MatchViewAccess"), access, tableName);
				enableSaveButton();
			}
		}, getPopupWidth(), getPopupHeight());

	}

	function selectMatchTableFilter() {

		var tableID = $("#MatchTableID").val();
		var currentID = $("#txtMatchFilterID").val();
		var tableName = $("#MatchTableID option:selected").text();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name, access) {
			//If current user is System Manager/Security Manager, we allow them to add or edit the filter/picklist hidden by another user
			if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower' && '@Model.CanEditSecurityGroups.ToString.ToLower' == "false") {
				$("#txtMatchFilterID").val(0);
				$("#txtMatchFilter").val('None');
				OpenHR.modalMessage("The " + tableName + " table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#txtMatchFilterID").val(id);
				$("#txtMatchFilter").val(name);
				setViewAccess('FILTER', $("#MatchViewAccess"), access, tableName);
				enableSaveButton();
			}
		}, getPopupWidth(), getPopupHeight());

	}


	function setTalentDefinitionDetails() {

		$('#MatchTableID').val("@Model.MatchTableID");
		MatchTableClick();

		$("#OriginalRoleTableID").val($('#BaseTableID').val());		
		$("#OriginalRoleTableText").val($("#BaseTableID option:selected").text());
		$("#OriginalPersonTableID").val($('#MatchTableID').val());
		$("#OriginalPersonTableText").val($("#MatchTableID option:selected").text());
		refreshTalentReportChildTables();
		$('#MatchChildTableID').val("@Model.MatchChildTableID");	
		
	}

	function refreshTalentReportChildTables() {

		$.ajax({
			url: 'Reports/GetChildTables?parentTableId=' + $("#BaseTableID").val(),
			datatype: 'json',
			mtype: 'GET',
			cache: false,
			success: function (json) {

				var option = "";
				for (var i = 0; i < json.length; i++) {
					option += "<option value='" + json[i].id + "'>" + json[i].Name + "</option>";
				}
				$("select#BaseChildTableID").html(option);
				$('#BaseChildTableID').val("@Model.BaseChildTableID");
				refreshTalentReportBaseColumns();
			}
		});

		$.ajax({
			url: 'Reports/GetChildTables?parentTableId=' + $("#MatchTableID").val(),
			datatype: 'json',
			mtype: 'GET',
			cache: false,
			success: function (json) {

				var option = "";
				for (var i = 0; i < json.length; i++) {
					option += "<option value='" + json[i].id + "'>" + json[i].Name + "</option>";
				}
				$("select#MatchChildTableID").html(option);
				$('#MatchChildTableID').val("@Model.MatchChildTableID");
				refreshTalentReportMatchColumns();

			}
		});

	}

	function refreshTalentReportBaseColumns() {

		var optionNone = "<option value='0' data-datatype='0' data-size='0' data-decimals='0'>None</option>";

		$.ajax({
			url: 'Reports/GetAvailableColumnsForTable?TableID=' + $("#BaseChildTableID").val(),
			datatype: 'json',
			mtype: 'GET',
			cache: false,
			success: function (json) {

				var option = "";

				for (var i = 0; i < json.length; i++) {
					option += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].ColumnSize + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";
				}

				$("select#BaseChildColumnID").html(option);
				$("select#BaseMinimumRatingColumnID").html(option);
				$("select#BasePreferredRatingColumnID").html(optionNone + option);

				$('#BaseChildColumnID').val("@Model.BaseChildColumnID");
				$('#BaseMinimumRatingColumnID').val("@Model.BaseMinimumRatingColumnID");
				$('#BasePreferredRatingColumnID').val("@Model.BasePreferredRatingColumnID");

			}
		});
	}

	function refreshTalentReportMatchColumns() {

		$.ajax({
			url: 'Reports/GetAvailableColumnsForTable?TableID=' + $("#MatchChildTableID").val(),
			datatype: 'json',
			mtype: 'GET',
			cache: false,
			success: function (json) {

				var option = "";

				for (var i = 0; i < json.length; i++) {
					option += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].ColumnSize + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";
				}

				$("select#MatchChildColumnID").html(option);
				$("select#MatchChildRatingColumnID").html(option);

				$('#MatchChildColumnID').val("@Model.MatchChildColumnID");
				$('#MatchChildRatingColumnID').val("@Model.MatchChildRatingColumnID");
			}
		});

	}


	$(function () {
		$("#tabs").tabs({
			activate: function (event, ui) {
				//Tab click event fired
				if (ui.newTab.text() == "Columns") {
					resizeColumnGrids();
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
		$('#description, #Name').css('width', $('#BaseTableID').width());
	});

	function resizeColumnGrids() {
		var gridWidth = $('#columnsAvailable').width() - 10;
		$("#AvailableColumns").jqGrid('setGridWidth', gridWidth);
		$('#SelectedTableID').width(gridWidth);

		gridWidth = $('#columnsSelected').width() - 10;
		$("#SelectedColumns").jqGrid('setGridWidth', gridWidth);

		//var gridHeight = $('#columnsAvailable').parent().height() - 20;
		var gridHeight = screen.height / 3;
		$("#SelectedColumns").jqGrid('setGridHeight', gridHeight);
		$("#AvailableColumns").jqGrid('setGridHeight', gridHeight);

		//column aggregate widths
		$('.colAggregates').find('.tablecell').css('width', gridWidth / 3);
	}

	function requestChangeReportPersonTable(target)
	{

		var columnCount = 0;
		var previousPersonTableID = $("#OriginalPersonTableID").val();		

		$("#IsPersonTableChange").val("True");
		var gridData = $("#SelectedColumns").jqGrid('getRowData');
		
		for (j = 0; j < gridData.length; j++) {
			if (gridData[j].TableID === previousPersonTableID) {
				columnCount = columnCount + 1;
				break;
			}
		}		

		if (columnCount > 0 ) {		
			OpenHR.modalPrompt("Changing the person table will result in all table/column specific aspects of this definition being cleared. <br/><br/>Are you sure you wish to continue ?", 4, "").then(function (answer) {
				if (answer == 6) { // Yes
					changeReportPersonTable();
					MatchTableClick();
				}
				else {
					$('#MatchTableID')[0].selectedIndex = $("#OriginalPersonTableID").val();
				}
			});
		}		
		else {
			changeReportPersonTable();
			MatchTableClick();
		}
	}

	function changeReportPersonTable() {
		// Post Person table change to server
		var dataSend = {
			ReportID: '@Model.ID',
			ReportType: '@Model.ReportType',
			MatchTableID: $("#MatchTableID option:selected").val(),
			BaseTableID: 0,
			__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
		};
		
		OpenHR.postData("Reports/changePersonTable", dataSend, changeReportPersonCompleted);
		
	}

	function changeReportPersonCompleted(json) {

		$("#matchselectiontype_All").prop('checked', 'checked');
		$("#ChildTablesAvailable").val(parseInt(json.childTablesAvailable));
		changeRecordOption('Match', 'ALL');
	
		if ($("#txtReportType").val() === '@UtilityType.TalentReport') {
			removeSelectedTableColumns(true,"personTable", $("#OriginalPersonTableText").val());
			refreshTalentReportChildTables(true);
		}
		
		// Enables save button
		enableSaveButton();
	}

	function MatchTableClick() {

		var BaseTableID = $("#BaseTableID").val();
		var MatchTableID = $("#MatchTableID").val();

		//Reset Base Table so none are disabled/hidden
		$('#BaseTableID option').removeAttr('disabled');

		//Hide/disable matching items in Base Table
		$('#BaseTableID option').filter(function () {
			return $(this).val() == MatchTableID;
		}).attr('disabled', 'disabled');
	}

	$("#workframe").attr("data-framesource", "UTIL_DEF_TALENTREPORT");
</script>
