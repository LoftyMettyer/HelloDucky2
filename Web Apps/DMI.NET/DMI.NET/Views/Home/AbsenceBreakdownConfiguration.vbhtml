@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Models
@Imports HR.Intranet.Server.Enums
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.StandardReportsConfigurationModel)

@Code
	Layout = Nothing
	Html.EnableClientValidation()
End Code

<div>
	@Using (Html.BeginForm("ABSENSE_BREAKDOWN_CONFIGURATION", "Home", New With {.id = "frmReportConfiguration",
																																					 .name = "frmReportConfiguration"}))
	@<div id="tabs">
		<ul>
			<li><a href="#tabs-1">Definition</a></li>
			<li><a href="#report_configuration_tab_output">Output</a></li>
		</ul>

	 	@Html.Hidden("recSelTableID", Model.TableId)

		<div id="tabs-1">

			<fieldset id="columnsAvailable">
				<legend class="fontsmalltitle">Absence Types :</legend>
				<table id="tblAbsenceTypes"></table>
			</fieldset>

			<fieldset id="selectiontypeallrecords" class="floatleft bordered width100">
				<legend class="fontsmalltitle">Date Range :</legend>
				<div class="inner">
					<fieldset class="floatleft overflowhidden width50">

						@Html.RadioButton("DateRange", Model.IsDefaultDate, Model.IsDefaultDate,
							New With {.id = "rdbDefaultDate", .onclick = "changeDateRange('DEFAULT')"})<span> Default (12 months to end of last month)</span>
					</fieldset>

					<fieldset class="floatleft overflowhidden width50">
						@Html.RadioButton("DateRange", Model.IsCustomDate, Model.IsCustomDate,
							New With {.id = "rdbCustomDate", .onclick = "changeDateRange('CUSTOM')"})<span>Custom</span>
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

End Using

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" class="ui-helper-hidden">
		@Code
			Html.RenderPartial("~/Views/Shared/gotoWork.ascx")
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

	});

	function selectAbsencePicklist() {

		var tableID = $("#recSelTableID").val();
		var currentID = $("#txtPicklistId").val();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name) {
			$("#txtPicklistId").val(id);
			$("#txtPicklistName").val(name);
		}, 400, 400);

	}

	function selectAbsenceFilter() {

		var tableID = $("#recSelTableID").val();
		var currentID = $("#txtPicklistId").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name) {
			$("#txtFilterId").val(id);
			$("#txtFilterName").val(name);
		}, 400, 400);

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
				//setViewAccess('CALC', $("#StartCustomViewAccess"), access, "report start date");
			}
		}, 400, 400);
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
					//setViewAccess('CALC', $("#EndCustomViewAccess"), access, "report end date");
				}
			}, 400, 400);

		}

		function fillAbenceTypes() {

			$("#tblAbsenceTypes").jqGrid({
				datatype: 'jsonstring',
				datastr: '@Model.AbsenceTypes.ToJsonResult',
				mtype: 'GET',
				jsonReader: {
					root: "rows", //array containing actual data
					page: "page", //current page
					total: "total", //total pages for the query
					records: "records", //total number of records
					repeatitems: false,
					id: "ID"
				},
				colNames: ['ID', '', 'Name'],
				colModel: [
										{ name: 'ID', width: 50, key: true, hidden: true, sorttype: 'integer' },
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
				$("#txtStartDateCalcId").val(0);
				$("#txtEndDateCalcId").val(0);
				$("#txtCustomStartDate").val("");
				$("#txtCustomEndDate").val("");
				button_disable($("#cmdBaseEndDateCalc")[0], true)
				button_disable($("#cmdBaseStartDateCalc")[0], true)
			}

			if (psType == "CUSTOM") {
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
			saveConfiguration();
		}

		function saveConfiguration() {

			var frmReportConfiguration = OpenHR.getForm("workframe", "frmReportConfiguration");

			//menu_refreshMenu();
			//var menuForm = OpenHR.getForm("menuframe", "frmMenuInfo");


			//if (frmConfiguration.chkOwner_Calculations.checked == true) frmConfiguration.txtOwner_Calculations.value = 1;
			//if (frmConfiguration.chkOwner_CrossTabs.checked == true) frmConfiguration.txtOwner_CrossTabs.value = 1;
			//if (frmConfiguration.chkOwner_NineBoxGrid.checked == true) frmConfiguration.txtOwner_NineBoxGrid.value = 1;
			//if (frmConfiguration.chkOwner_CustomReports.checked == true) frmConfiguration.txtOwner_CustomReports.value = 1;
			//if (frmConfiguration.chkOwner_Filters.checked == true) frmConfiguration.txtOwner_Filters.value = 1;
			//if (frmConfiguration.chkOwner_MailMerge.checked == true) frmConfiguration.txtOwner_MailMerge.value = 1;
			//if (frmConfiguration.chkOwner_Picklists.checked == true) frmConfiguration.txtOwner_Picklists.value = 1;
			//if (frmConfiguration.chkOwner_CalendarReports.checked == true) frmConfiguration.txtOwner_CalendarReports.value = 1;

			OpenHR.submitForm(frmReportConfiguration);
		}

		$("#workframe").attr("data-framesource", "ABSENCE_BREAKDOWN_CONFIGURATION");
</script>
