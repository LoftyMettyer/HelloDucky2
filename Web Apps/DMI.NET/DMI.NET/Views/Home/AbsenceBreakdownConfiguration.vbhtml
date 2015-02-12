@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Models
@Imports HR.Intranet.Server.Enums
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of StandardReportsConfigurationModel)

@Code
	Layout = Nothing
	Html.EnableClientValidation()
End Code

<div>
	@Using (Html.BeginForm("Absence_Breakdown_Configuration", "Home", FormMethod.Post, New With {.id = "frmReportConfiguration",
																																					 .name = "frmReportConfiguration"}))
	@<div id="tabs">
	<ul>
		<li><a href="#tabs-1">Definition</a></li>
		<li><a href="#report_configuration_tab_output">Output</a></li>
	</ul>

	@Html.Hidden("recSelTableID", Model.TableId)
	@Html.HiddenFor(Function(m) m.AbsenceTypesAsString, New With {.id = "txtAbsenceTypes"})
	@Html.HiddenFor(Function(m) m.ReportType, New With {.id = "reportType"})


	<div id="tabs-1">

		<fieldset id="columnsAvailable">
			<legend class="fontsmalltitle">Absence Types :</legend>
			<table id="tblAbsenceTypes"></table>
		</fieldset>

		<fieldset id="selectiontypeallrecords" class="floatleft bordered width100">
			<legend class="fontsmalltitle">Date Range :</legend>
			<div class="inner">
				<fieldset class="floatleft overflowhidden width50">

					@Html.RadioButton("DateRange", True, Model.IsDefaultDate, New With {.id = "rdbDefaultDate", .onclick = "changeDateRange('DEFAULT')"})<span> Default (12 months to end of last month)</span>
				</fieldset>

				<fieldset class="floatleft overflowhidden width50">
					@Html.RadioButton("DateRange", True, Model.IsCustomDate, New With {.id = "rdbCustomDate", .onclick = "changeDateRange('CUSTOM')"})<span>Custom</span>
					@Html.HiddenFor(Function(m) m.IsCustomDate, New With {.id = "CustomEndDate"})
				</fieldset>

				<fieldset class="floatleft overflowhidden width50">
					<div id="CustomStartDate" class="floatleft">
						<span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Start Date :</span>
					</div>
					<div class="width70  floatright">
						@Html.EllipseButton("cmdBaseStartDateCalc", "selectCustomStartDate()", Model.IsCustomDate)
						<div class="ellipsistextbox">
							@Html.TextBoxFor(Function(m) m.StartDate, New With {.id = "txtCustomStartDate", .readonly = "true"})
						</div>
					</div>
					@Html.HiddenFor(Function(m) m.StartDateId, New With {.id = "txtStartDateCalcId"})
				</fieldset>

				<fieldset class="floatleft overflowhidden width50">
					<div id="CustomEndDate" class="floatleft">
						<span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End Date :</span>
					</div>
					<div class="width70  floatright">
						@Html.EllipseButton("cmdBaseEndDateCalc", "selectCustomEndDate()", Model.IsCustomDate)
						<div class="ellipsistextbox">
							@Html.TextBoxFor(Function(m) m.EndDate, New With {.id = "txtCustomEndDate", .readonly = "true"})
						</div>
					</div>
					@Html.HiddenFor(Function(m) m.EndDateId, New With {.id = "txtEndDateCalcId"})
				</fieldset>
			</div>
		</fieldset>

		<fieldset id="selectiontypeallrecords" class="floatleft bordered width100">
			<legend><label class="fontsmalltitle">Record Selection :</label> @Html.CheckBoxFor(Function(m) m.DisplayTitleInReportHeader, New With {.id = "chkDisplayTitleInReportHeader", .readonly = "true"})<label> Display filter or picklist title in the report header</label></legend>
			<div class="inner">

				<fieldset class="floatleft overflowhidden width50">

					@Html.RadioButton("Selectiontype", RecordSelectionType.AllRecords, Model.SelectionType = RecordSelectionType.AllRecords,
								New With {.id = "rdbSelectAll", .onclick = "changeRecordOption('ALL')"})<span> All Records</span>
				</fieldset>

				<fieldset id="selectiontypepicklistgroup" class="floatleft overflowhidden width50">
					<div id="PicklistRadioDiv" class="floatleft">
						@Html.RadioButton("SelectionType", RecordSelectionType.Picklist,
														Model.SelectionType = RecordSelectionType.Picklist,
									New With {.id = "rdpPicklistType", .onclick = "changeRecordOption('PICKLIST')"})
						<span>Picklist</span>
					</div>
					<div class="width70 floatright">
						@Html.EllipseButton("cmdBasePicklist", "selectAbsencePicklist()", Model.SelectionType = RecordSelectionType.Picklist)
						<div class="ellipsistextbox">
							@Html.TextBoxFor(Function(m) m.PicklistName, New With {.id = "txtPicklistName", .readonly = "true"})
						</div>
					</div>
					<input type="hidden" id="txtPicklistId" name="picklistID" value="@Model.PicklistId" />
					@Html.ValidationMessageFor(Function(m) m.PicklistId)
				</fieldset>

				<fieldset id="selectiontypefiltergroup" class="floatleft overflowhidden width50">
					<div id="FilterRadioDiv" class="floatleft">
						@Html.RadioButton("SelectionType", RecordSelectionType.Filter, Model.SelectionType = RecordSelectionType.Filter _
									, New With {.id = "rdbFilterType", .onclick = "changeRecordOption('FILTER')"})
						<span>Filter</span>
					</div>
					<div class="width70  floatright">
						@Html.EllipseButton("cmdBaseFilter", "selectAbsenceFilter()", Model.SelectionType = RecordSelectionType.Filter)
						<div class="ellipsistextbox">
							@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtFilterName", .readonly = "true"})
						</div>
					</div>
					<input type="hidden" id="txtFilterId" name="filterID" value="@Model.FilterId" />
					@Html.ValidationMessageFor(Function(m) m.FilterId)
				</fieldset>
			</div>
		</fieldset>
	</div>

	<div id="report_configuration_tab_output">
		output
	</div>
</div>
	@Html.AntiForgeryToken()
 End Using

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" class="ui-helper-hidden">
		@Code
			Html.RenderPartial("~/Views/Shared/gotoWork.ascx")
			Html.AntiForgeryToken()
		End Code
	</form>

</div>

<script type="text/javascript">

	$(function () {
		$("#tabs").tabs();

		$("#toolbarStandardReportConfig").parent().show();
		$("#toolbarStandardReportConfig").click();
		menu_setVisibleMenuItem("mnutoolSaveStandardReportConfig", true);
		menu_toolbarEnableItem("mnutoolSaveStandardReportConfig", false);

		$('table').attr('border', '0');
		$('fieldset').css("border", '0');

		$('input').on("input", function () { enableSaveButton(); });
		$('select').on("change", function () { enableSaveButton(); });
		$('input').on("change", function () { enableSaveButton(); });

		fillAbenceTypes();

		if('@Model.SelectionType' == 'AllRecords')
		{
			$("#chkDisplayTitleInReportHeader")[0].checked = false;
			$("#chkDisplayTitleInReportHeader")[0].disabled = true;
		}

	});

	function selectAbsencePicklist() {

		var tableID = $("#recSelTableID").val();
		var currentID = $("#txtPicklistId").val();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name) {
			$("#txtPicklistId").val(id);
			$("#txtPicklistName").val(name);
			enableSaveButton();
		}, getPopupWidth(), getPopupHeight());

	}

	function selectAbsenceFilter() {

		var tableID = $("#recSelTableID").val();
		var currentID = $("#txtFilterId").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name) {
			$("#txtFilterId").val(id);
			$("#txtFilterName").val(name);
			enableSaveButton();
		}, getPopupWidth(), getPopupHeight());

	}

	function selectCustomStartDate() {

		var tableID = $("#recSelTableID").val();
		var currentID = $("#txtStartDateCalcId").val();

		OpenHR.modalExpressionSelect("CALC", 0, currentID, function (id, name, access) {
			if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower') {
				$("#txtStartDateCalcId").val(0);
				$("#txtCustomStart").val('None');
				OpenHR.modalMessage("The report start date calculation will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#txtStartDateCalcId").val(id);
				$("#txtCustomStartDate").val(name);
				enableSaveButton();
			}
		}, getPopupWidth(), getPopupHeight());
		}

		function selectCustomEndDate() {

			var tableID = $("#recSelTableID").val();
			var currentID = $("#txtEndDateCalcId").val();

			OpenHR.modalExpressionSelect("CALC", 0, currentID, function (id, name, access) {
				if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower') {
					$("#txtEndDateCalcId").val(0);
					$("#txtCustomEndDate").val('None');
					OpenHR.modalMessage("The report end date calculation will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
				}
				else {
					$("#txtEndDateCalcId").val(id);
					$("#txtCustomEndDate").val(name);
					enableSaveButton();
				}
			}, getPopupWidth(), getPopupHeight());

		}

		function fillAbenceTypes() {

			$("#tblAbsenceTypes").jqGrid({
				datatype: 'jsonstring',
				datastr: '@Model.AbsenceTypes.ToJsonResult',
				jsonReader: {
					root: "rows", //array containing actual data
					page: "page", //current page
					total: "total", //total pages for the query
					records: "records", //total number of records
					repeatitems: false,
					id: "0"
				},
				rowNum:'@Model.AbsenceTypes.Count',
				colNames: ['', 'Name'],
				colModel: [
										{
											name: 'IsSelected', width: 40, align: "center", editable: true, edittype: 'checkbox', editoptions: { value: "True:False" }, formatter: "checkbox", formatoptions: { disabled: false }
										},
										{ name: 'Type', width: 300, formatoptions: { disabled: false } }
				],
				onSelectRow: function (id) {
					//ToDO
				},
				loadComplete: function () {

					$('#tblAbsenceTypes input[type=checkbox]').on('click', function () { enableSaveButton(); });

					var topID = $("#tblAbsenceTypes").getDataIDs()[0]
					$("#tblAbsenceTypes").jqGrid("setSelection", topID);
				}
			});

		}

	function changeRecordOption(psType) {

			if (psType == "ALL") {
				$("#txtFilterName").val("");
				$("#txtPicklistName").val("");
				$("#txtPicklistId").val(0);
				$("#txtFilterId").val(0);
				$("#chkDisplayTitleInReportHeader")[0].checked = false;
				$("#chkDisplayTitleInReportHeader")[0].disabled = true;
			}

			if (psType == "PICKLIST") {
				button_disable($("#cmdBasePicklist")[0], false)
				button_disable($("#cmdBaseFilter")[0], true)
				$("#txtFilterName").val("");
				$("#txtFilterId").val(0);

				if ($("#txtPicklistId").val() == 0) {
					$("#txtPicklistName").val("None");
				}

				$("#chkDisplayTitleInReportHeader")[0].disabled = false;

			}

			if (psType == "FILTER") {
				button_disable($("#cmdBasePicklist")[0], true)
				button_disable($("#cmdBaseFilter")[0], false)
				$("#txtPicklistName").val("");
				$("#txtPicklistId").val(0);

				if ($("#txtFilterId").val() == 0) {
					$("#txtFilterName").val("None");
				}

				$("#chkDisplayTitleInReportHeader")[0].disabled = false;
			}
		}

	function changeDateRange(psType) {
		if (psType == "DEFAULT") {
			$("#CustomEndDate").val(false);
				$("#txtStartDateCalcId").val(0);
				$("#txtEndDateCalcId").val(0);
				$("#txtCustomStartDate").val("");
				$("#txtCustomEndDate").val("");
				button_disable($("#cmdBaseEndDateCalc")[0], true)
				button_disable($("#cmdBaseStartDateCalc")[0], true)
			}

			if (psType == "CUSTOM") {
				$("#CustomEndDate").val(true);
				button_disable($("#cmdBaseEndDateCalc")[0], false)
				button_disable($("#cmdBaseStartDateCalc")[0], false)
				$("#txtCustomStartDate").val("None");
				$("#txtCustomEndDate").val("None");
			}
		}


		function enableSaveButton() {
			menu_toolbarEnableItem('mnutoolSaveStandardReportConfig', true);
		}

		function ReportConfiguration_okClick() {
			if ($("#mnutoolSaveStandardReportConfig")[0].className == 'button ui-corner-all ui-state-default ui-state-hover') {
				saveConfiguration();
			}
		}

		function saveChanges(psAction, pfPrompt, pfTBOverride) {
			if ($("#mnutoolSaveStandardReportConfig")[0].className == 'button ui-corner-all ui-state-default disabled' ||
				$("#mnutoolSaveStandardReportConfig")[0].className == 'button ui-corner-all ui-state-default ui-state-hover disabled') {
				return 6; //No to saving the changes, as none have been made.
			} else
				return 0;
		}

		function saveConfiguration() {
			var frmReportConfiguration = OpenHR.getForm("workframe", "frmReportConfiguration");

			var gridData = $("#tblAbsenceTypes").jqGrid('getRowData');
			$('#txtAbsenceTypes').val(JSON.stringify(gridData));

			var flagOk = true;
			if ($("#CustomEndDate").val().toLowerCase() == "true") {
				if ($("#txtStartDateCalcId").val() == "0" && $("#txtEndDateCalcId").val() == "0") {
					OpenHR.modalPrompt("You must select Start and End Date calculation");
					flagOk = false;
				}
				else if ($("#txtStartDateCalcId").val() != "0" && $("#txtEndDateCalcId").val() == "0") {
					OpenHR.modalPrompt("You must select an End Date calculation");
					flagOk = false;
				}
				else if ($("#txtStartDateCalcId").val() == "0" && $("#txtEndDateCalcId").val() != "0") {
					OpenHR.modalPrompt("You must select Start Date calculation");
					flagOk = false;
				}
			}

			if (flagOk) {
				OpenHR.submitForm(frmReportConfiguration);
			}
		}

		$("#workframe").attr("data-framesource", "ABSENCE_BREAKDOWN_CONFIGURATION");
</script>
