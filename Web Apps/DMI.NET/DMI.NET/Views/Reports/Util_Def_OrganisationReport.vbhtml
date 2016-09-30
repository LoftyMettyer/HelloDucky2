@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Inherits System.Web.Mvc.WebViewPage(Of Models.OrganisationReportModel)

@Code
   Layout = Nothing
   Html.EnableClientValidation(True)
End Code

<div>
   @Code
      Html.EnableClientValidation(True)
   End Code

   @Using (Html.BeginForm("Util_Def_OrganisationReport", "Reports", FormMethod.Post, New With {.id = "frmReportDefintion", .name = "frmReportDefintion"}))
   @Html.HiddenFor(Function(m) m.ID)

      @<div id="tabs">
         <ul>
            <li><a href="#tabs-1">Definition</a></li>
            <li><a href="#tabs-2">Filter</a></li>
            <li><a href="#report_definition_tab_columns">Columns</a></li>
         </ul>

         <div id="tabs-1">
            @Code
               Html.RenderPartial("_OrgDefinition", Model)
            End Code
         </div>

         <div id="tabs-2">
            @Code
               Html.RenderPartial("_OrgFilterSelect", Model)
            End Code
         </div>

         <div id="report_definition_tab_columns">
            @Code
               Html.RenderPartial("_OrgColumnSelection", Model)
            End Code
         </div>
          <div id="divPopupPreview" style="display: none;"></div>
      </div>

      @Html.AntiForgeryToken()
               End Using

</div>

<script type="text/javascript">

   // A $( document ).ready() block.
   $(document).ready(function () {
      getAvailableTableViewColumns();
      attachGridToSelectedColumns();
      getUnauthorisedColumns();
      validateBaseView();
   });

   $(function () {
      $("#tabs").tabs({
         activate: function (event, ui) {
            //Tab click event fired
            if (ui.newTab.text() == "Columns") {
               var topID = $("#SelectedColumns").getDataIDs()[0]
               $('#SelectedColumns').jqGrid('resetSelection');
               $("#SelectedColumns").jqGrid('setSelection', topID);
               resizeColumnGrids();
            }
            if (ui.newTab.text() == "Filter") {
               refreshSelectColumnCombo();
            }

         }
      });

      $('input[type=number]').numeric();
      $('#description, #Name').css('width', $('#BaseTableID').width());

   });

   $("#workframe").attr("data-framesource", "UTIL_DEF_ORGANISATIONREPORT");

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

</script>
