﻿@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.CalendarReportModel)

@Code
	Layout = Nothing
End Code

<div>
	@Using (Html.BeginForm("util_def_calendarreport", "Reports", FormMethod.Post, New With {.id = "frmReportDefintion", .name = "frmReportDefintion"}))
		@Html.HiddenFor(Function(m) m.ID)

		@<div id="tabs">
			<ul>
				<li><a href="#tabs-1">Definition</a></li>
				<li><a href="#report_definition_tab_eventdetails">Event Details</a></li>
				<li><a href="#report_definition_tab_reportdetails">Report Details</a></li>
				<li><a href="#report_definition_tab_order">Order</a></li>
				<li><a href="#report_definition_tab_output">Output</a></li>
			</ul>

			<div id="tabs-1">
 		<div class="width100">
		 	<fieldset>
				@Code
				Html.RenderPartial("_Definition", Model)
				End Code

				@Html.LabelFor(Function(m) m.Description1ID)
				@Html.TextBoxFor(Function(m) m.Description1ID)
		 		<br />
				@Html.LabelFor(Function(m) m.Description2ID)
				@Html.TextBoxFor(Function(m) m.Description2ID)
				<br />
				@Html.LabelFor(Function(m) m.Description2ID)
				@Html.TextBoxFor(Function(m) m.Description3ID)

				@Html.HiddenFor(Function(m) m.RegionID)
				@Html.HiddenFor(Function(m) m.GroupByDescription)
				@Html.HiddenFor(Function(m) m.Separator)
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


	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
		@Code
			Html.RenderPartial("~/Views/Shared/gotoWork.ascx")
		End Code
	</form>
</div>
	End Using
</div>
<script type="text/javascript">

	$(function () {
		$("#tabs").tabs();
		$('input[type=number]').numeric();

		if ($("#IsReadOnly").val() == "True") {
			$("#frmReportDefintion :input").prop("disabled", true);
		}

		button_disable($("#btnSortOrderAdd")[0], false);
	});

	function submitForm() {
		var frmSubmit = $("#frmReportDefintion");
		OpenHR.submitForm(frmSubmit);
	}

	$("#workframe").attr("data-framesource", "UTIL_DEF_CALENDARREPORT");


</script>
