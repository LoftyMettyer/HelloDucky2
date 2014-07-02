@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.CrossTabModel)


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
</style>


<div>

	@Using (Html.BeginForm("util_def_crosstabs_submit", "Reports", FormMethod.Get, New With {.id = "frmReportDefintion", .name = "frmReportDefintion"}))

		@Html.HiddenFor(Function(m) m.ID)

		@<div id="tabs">

			<ul>
				<li><a href="#tabs-1">Definition</a></li>
				<li><a href="#report_definition_tab_columns">Columns</a></li>
				<li><a href="#report_definition_tab_output">Output</a></li>
			</ul>

			<div id="tabs-1">
				@Code
				Html.RenderPartial("Definition", Model)
				End Code
			</div>

			<div id="report_definition_tab_columns">
				@Code
				Html.RenderPartial("CrossTabsColumnSelection", Model)
				End Code
			</div>

			<div id="report_definition_tab_output">
				@Code
				Html.RenderPartial("Output", Model.Output)
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
	});


	function submitForm() {
		var frmSubmit = $("#frmReportDefintion");
		OpenHR.submitForm(frmSubmit);
	}

	$("#workframe").attr("data-framesource", "UTIL_DEF_CROSSTABS");

</script>
