﻿@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.CrossTabModel)

@Code
	Layout = Nothing
	Html.EnableClientValidation()
End Code

<div>
	@Using (Html.BeginForm("util_def_crosstab", "Reports", FormMethod.Post, New With {.id = "frmReportDefintion", 
																																									 .name = "frmReportDefintion"}))

 		@Html.HiddenFor(Function(m) m.IsReadOnly)
		@Html.HiddenFor(Function(m) m.ID)

		@<div id="tabs">
			<ul>
				<li><a href="#tabs-1">Definition</a></li>
				<li><a href="#report_definition_tab_columns">Columns</a></li>
				<li><a href="#report_definition_tab_output">Output</a></li>
			</ul>

			<div id="tabs-1">
				
					@Code				
					Html.RenderPartial("_Definition", Model)
					End Code
				
			</div>

			<div id="report_definition_tab_columns">
				@Code
				Html.RenderPartial("_CrossTabsColumnSelection", Model)
				End Code
			</div>

			<div id="report_definition_tab_output">
				@Code
				Html.RenderPartial("_Output", Model.Output)
				End Code
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
		$('input[type=number]').numeric();

		if ($("#IsReadOnly").val() == "True") {
			$("#frmReportDefintion :input").prop("disabled", true);
		}	

	});

	$("#workframe").attr("data-framesource", "UTIL_DEF_CROSSTABS");

</script>
