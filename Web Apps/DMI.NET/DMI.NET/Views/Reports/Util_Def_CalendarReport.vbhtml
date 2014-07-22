@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.CalendarReportModel)

@Code
	Layout = Nothing
End Code

<style>
	.wrapper {
		width: 100%;
		overflow-x: auto;
		overflow-y: hidden;
	}

	.inner {
		width: 100%;
	}

	.left {
		width: 50%;
		float: left;
	}

	.right {
		width: 50%;
		float: left;
	}

	input[readonly="true"] {
		background-color: #F2F2F2 !important;
		color: #826D82;
		border-color: #ddd;
		pointer-events: none;
		cursor: default;
	}

	input[readonly="readonly"] {
		background-color: #F2F2F2 !important;
		color: #826D82;
		border-color: #ddd;
		pointer-events: none;
		cursor: default;
	}

</style>

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
				@Code
				Html.RenderPartial("_Definition", Model)
				End Code

				@Html.LabelFor(Function(m) m.Description1ID)
				@Html.TextBoxFor(Function(m) m.Description1ID)
				<br/>
				@Html.LabelFor(Function(m) m.Description2ID)
				@Html.TextBoxFor(Function(m) m.Description2ID)
				<br />
				@Html.LabelFor(Function(m) m.Description2ID)
				@Html.TextBoxFor(Function(m) m.Description3ID)

				@Html.HiddenFor(Function(m) m.RegionID)
				@Html.HiddenFor(Function(m) m.GroupByDescription)
				@Html.HiddenFor(Function(m) m.Separator)

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

	End Using

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
		@Code
			Html.RenderPartial("~/Views/Shared/gotoWork.ascx")
		End Code
	</form>




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
