@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.MailMergeModel)

@Code
	Layout = Nothing
End Code

<div>
  @Using (Html.BeginForm("util_def_mailmerge", "Reports", FormMethod.Post, New With {.id = "frmReportDefintion", .name = "frmReportDefintion"}))

    @Html.HiddenFor(Function(m) m.ID)

    @<div id="tabs">
      <ul>
				<li><a href="#tabs-1">Definition</a></li>
        <li><a href="#report_definition_tab_columns">Columns</a></li>
        <li><a href="#report_definition_tab_order">Sort Order</a></li>
				<li><a href="#report_definition_tab_output">Output</a></li>
      </ul>

      <div id="tabs-1" class="">	
				<fieldset>
					@Code
					Html.RenderPartial("_Definition", Model)
					End Code
				</fieldset>							
      </div>

			<div id="report_definition_tab_columns">
				
					@Code
					Html.RenderPartial("_ColumnSelection", Model)
					End Code
				
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
					Html.RenderPartial("_MergeOutput", Model)
					End Code
				</fieldset>
      </div>
    </div>
  End Using

  <form action="default_Submit" method="post" id="frmGoto" name="frmGoto"class="ui-helper-hidden">
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

  $("#workframe").attr("data-framesource", "UTIL_DEF_MAILMERGE");
</script>
