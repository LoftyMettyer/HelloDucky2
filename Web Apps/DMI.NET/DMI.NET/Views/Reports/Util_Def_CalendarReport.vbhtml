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
		@Html.HiddenFor(Function(m) m.Description3ViewAccess)		 

		@<div id="tabs">
	<ul>
		<li><a href="#tabs-1">Definition</a></li>
		<li><a href="#report_definition_tab_eventdetails">Event Details</a></li>
		<li><a href="#report_definition_tab_reportdetails">Report Details</a></li>
		<li><a href="#report_definition_tab_order">Order</a></li>
		<li><a href="#report_definition_tab_output">Output</a></li>
	</ul>

	<div id="tabs-1" class="width100">
		@Code
		Html.RenderPartial("_Definition", Model)
		End Code

		<fieldset class="width100">
			<legend class="fontsmalltitle">Report Options:</legend>

			@Html.LabelFor(Function(m) m.Description1ID)
			@Html.ColumnDropdownFor(Function(m) m.Description1ID, New ColumnFilter() With {.TableID = Model.BaseTableID, .AddNone = True}, Nothing)

			<br />
			@Html.LabelFor(Function(m) m.Description2ID)
			@Html.ColumnDropdownFor(Function(m) m.Description2ID, New ColumnFilter() With {.TableID = Model.BaseTableID, .AddNone = True}, Nothing)

			<br />
			@Html.LabelFor(Function(m) m.Description3ID)
			@Html.HiddenFor(Function(m) m.Description3ID)
			<input type="text" id="txtDescription3" value="@Model.Description3Name" disabled />
			<input type="button" id="cmdDescription3" value="..." onclick="selectDescription3()" />

			<br/>
			@Html.LabelFor(Function(m) m.RegionID)
			@Html.ColumnDropdownFor(Function(m) m.Description2ID, New ColumnFilter() With {.TableID = Model.BaseTableID, .AddNone = True, .DataType = ColumnDataType.sqlVarChar}, New With {.id = "cboRegionID"})

			<br/>
			@Html.CheckBoxFor(Function(m) m.GroupByDescription, New With {.id = "chkGroupByDescription", .onclick = "selectGroupByDescription()"})
			@Html.LabelFor(Function(m) m.GroupByDescription)

			<br/>
			@Html.LabelFor(Function(m) m.Separator)
			@Html.DropDownList("Separator", New SelectList(New List(Of String)() From {"None", "Space", ",", ".", "-", ":", ";", "/", "\", "#", "~", "^"}))

		</fieldset>

	</div>

	<div id="report_definition_tab_eventdetails">
		<fieldset>
			@Code
			 Html.RenderPartial("_EventDetails", Model)
			End Code
		</fieldset>
	</div>

	<div id="report_definition_tab_reportdetails">
		<fieldset>
			@Code
			 Html.RenderPartial("_ReportDetails", Model)
			End Code
		</fieldset>
	</div>

	<div id="report_definition_tab_order">
		<fieldset>
			@Code
			Html.RenderPartial("_SortOrder", Model)
			End Code
		</fieldset>
	</div>

	<div id="report_definition_tab_output">
		<fieldset>
			@Code
				Html.RenderPartial("_Output", Model.Output)
			End Code
		</fieldset>
	</div>

</div>
	End Using
</div>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
	@Code
			Html.RenderPartial("~/Views/Shared/gotoWork.ascx")
	End Code
</form>

<script type="text/javascript">

	function selectGroupByDescription() {

		var bSelected = $("#chkGroupByDescription").prop('checked');
		combo_disable($("#cboRegionID")[0], bSelected);

	}

	function selectDescription3() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#Description3ID").val();

		OpenHR.modalExpressionSelect("CALC", tableID, currentID, function (id, name, access) {
			$("#Description3ID").val(id);
			$("#txtDescription3").val(name);
			setViewAccess('CALC', $("#Description3ViewAccess"), access, "report description");
		});

	}

	$(function () {
		$("#tabs").tabs();
		$('input[type=number]').numeric();

		if ($("#IsReadOnly").val() == "True") {
			$("#frmReportDefintion :input").prop("disabled", true);
		}

		selectGroupByDescription();
		button_disable($("#btnSortOrderAdd")[0], false);
	});

	function submitForm() {
		var frmSubmit = $("#frmReportDefintion");
		OpenHR.submitForm(frmSubmit);
	}

	$("#workframe").attr("data-framesource", "UTIL_DEF_CALENDARREPORT");


</script>
