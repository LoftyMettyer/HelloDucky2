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

      <div id="tabs-1">
					@Code
					Html.RenderPartial("_Definition", Model)
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
					Html.RenderPartial("_MergeOutput", Model)
					End Code				
      </div>
    </div>
		@Html.AntiForgeryToken()
  End Using

  <form action="default_Submit" method="post" id="frmGoto" name="frmGoto"class="ui-helper-hidden">
    @Code
      Html.RenderPartial("~/Views/Shared/gotoWork.ascx")
    End Code
	@Html.AntiForgeryToken()
  </form>
</div>

<script type="text/javascript">

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

  $("#workframe").attr("data-framesource", "UTIL_DEF_MAILMERGE");
</script>
