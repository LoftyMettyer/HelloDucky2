@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Inherits System.Web.Mvc.WebViewPage(Of Models.CustomReportModel)

@Code
	Layout = Nothing
	Html.EnableClientValidation(True)
End Code

<div>
	@Code
		Html.EnableClientValidation(True)
	End Code
	@Using (Html.BeginForm("util_def_customreport", "Reports", FormMethod.Post, New With {.id = "frmReportDefintion", .name = "frmReportDefintion"}))
	
  @Html.HiddenFor(Function(m) m.ID)
                     
  @<div id="tabs">
        <ul>
            <li><a href="#tabs-1">Definition</a></li>
            <li><a href="#tabs-2">Related Tables</a></li>
            <li><a href="#report_definition_tab_columns">Columns</a></li>
            <li><a href="#report_definition_tab_order">Sort Order</a></li>
            <li><a href="#report_definition_tab_output">Output</a></li>
        </ul>

        <div id="tabs-1">
		<div class="width100">
          @Code                        
						Html.RenderPartial("_Definition", Model)
          End Code
			<fieldset class="width100">
				<fieldset>
					<legend class="fontsmalltitle">Report Options: (MOVED FROM Output tab (discuss))</legend>
						@Html.CheckBox("IsSummary", Model.IsSummary)Summary Report
					<br />
						@Html.CheckBox("IgnoreZerosForAggregates", Model.IgnoreZerosForAggregates)Ignore zeros when calculating aggregates

				</fieldset>
			</fieldset>
					</div>
        </div>

        <div id="tabs-2">
          @Code
						Html.RenderPartial("_RelatedTables", Model)
          End Code
        </div>

        <div id="report_definition_tab_columns">
          @Code
						Html.RenderPartial("_ColumnSelection", Model)
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

	<div id="divEmailGroupSelection">

	</div>
    </div>


  


  End Using

  <form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
    @Code
      Html.RenderPartial("~/Views/Shared/gotoWork.ascx")
    End Code
  </form>

	<div id="eventdetail">
	</div>
	
</div>


<script type="text/javascript">

  $(function () {
  	$("#tabs").tabs();
  	$('input[type=number]').numeric();

  	if ($("#IsReadOnly").val() == "True") {
  		$("#frmReportDefintion :input").prop("disabled", true);
  	}

			$(function () {

  });




		});


    $("#workframe").attr("data-framesource", "UTIL_DEF_CUSTOMREPORTS");

    $('#tabs').bind('tabsshow', function (event, ui) {

      var tabPage;

      if (ui.index == "0") {
        tabPage = $("#frmCustomReportsTab1");
      }
      
      if (ui.index == "1") {
        tabPage = $("#frmCustomReportsTab2");
      }

      if (ui.index == "2") {
        tabPage = $("#frmCustomReportsTab3");
      }
    })

</script>
