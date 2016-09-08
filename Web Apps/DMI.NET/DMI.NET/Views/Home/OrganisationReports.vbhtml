@Imports DMI.NET
@Imports DMI.NET.Classes

@Inherits System.Web.Mvc.WebViewPage(Of OrganisationReportChartModel)

<link href=@Url.LatestContent("~/Scripts/jquery/jOrgChart/css/jquery.jOrgChart.css") rel="stylesheet" />
<link href=@Url.LatestContent("~/Scripts/jquery/jOrgChart/css/custom.css") rel="stylesheet" />
<link href=@Url.LatestContent("~/Scripts/jquery/jOrgChart/css/prettify.css") rel="stylesheet" />

<style>
   .truncate {
      width: 95%;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
   }

   .jOrgChart .node {
      height: 200px !important;
      width: auto;
      min-width: 200px;
      border: 1px solid gray;
      padding: 5px 0px 0px 0px;
      overflow: auto;
      font-weight: bold !important;
   }

   .expandNode {
      bottom: 4px;
      right: 4px;
   }

   .ui-state-active, .ui-widget-content .ui-state-active, .ui-widget-header .ui-state-active {
      /*background: none !important;*/
      /*color: black !important;*/
   }

   /*margin: 0px !important;*/
</style>
<script>

   $(document).ready(function () {

      // Common logic to show desired ribbon and menu
      $("#workframe").attr("data-framesource", "ORGREPORTS");
      showDefaultRibbon();
      menu_refreshMenu();

      if ('@Model.OrgReportChartNodeList.Any()' == 'False') {
         $('#noData').show();
         menu_toolbarEnableItem('divBtnPrintOrgChart', false);
         menu_toolbarEnableItem('divBtnPrintPreviewOrgChart', false);
         menu_toolbarEnableItem('mnutoolOrgReportsExpand', false);
         menu_toolbarEnableItem('divBtnSelectOrgReports', false);
         $('.mnuBtnPrintOrgChart>span').prop('disabled', true);
         $('.mnuBtnPrintOrgChart').prop('disabled', true);
         $('.mnuBtnPrintPreviewOrgChart>span').prop('disabled', true);
         $('.mnuBtnPrintPreviewOrgChart').prop('disabled', true);
         $('.mnuBtnSelectOrgChart>span').prop('disabled', true);
         $('.mnuBtnSelectOrgChart').prop('disabled', true);

      } else {
         $("#tempList").find("li").each(function () {
            var lineManagerStaffNo = $(this).attr("id");
            var hierarchyLevel = $(this).attr("hierarchyLevel");
            var parentNode = hierarchyLevel == "0" ? 'org' : lineManagerStaffNo;
            $('#' + parentNode).append($(this));
         });

         //Add a class to collapse all peer trees.
         $("#org li.ui-state-highlight").siblings().addClass("collapsed");
         $("#org li.ui-state-highlight").parents('li').siblings().addClass("collapsed");

         $('#workframe').attr('overflow', 'auto');
         $("#org").jOrgChart({
            chartElement: '#chart',
            dragAndDrop: false
         });

         setTimeout(centreMe(true), 500);


         $("#chart").find("#divMainContainer").each(function () {
            var totalEmpDiv = $(this).find("#divPostEmployees").children().length;
            if (totalEmpDiv > 1) {
               $(this).parent().css("max-width", totalEmpDiv * 200);
               $(this).parent().css("min-width", totalEmpDiv * 200);
               $(this).find("#divPostTitle").css("max-width", totalEmpDiv * 170);
               $(this).find("#divPostTitle").css("min-width", totalEmpDiv * 170);
               $(this).find("#divPostColumns div").css("width", totalEmpDiv * 180);
               $(this).find("#divPostColumns div").css("display", "inline-block");
               $(this).find("#divPostEmployees").css("padding-left", 8);
               $(this).find("#divPostEmployees").css("padding-right", 8);
            }
            else {
               $(this).parent().css("max-width", 200);
               $(this).parent().css("min-width", 200);
               $(this).find("#divPostTitle").css("max-width", 170);
               $(this).find("#divPostTitle").css("min-width", 170);
               $(this).find("#divPostColumns").css("display", "inline-table");
               $(this).find("#divPostColumns div").css("width", 170);
               $(this).find("#divPostColumns div").css("margin-left", 5);
               $(this).find("#divPostColumns div").css("margin-top", 2);
               $(this).find("#divPostColumns div").css("margin-bottom", 2);
            }
         });

         $('.printSelect').toggle().prop('checked', true);


         //($("#divPostEmployees").children().length * 195) + "px !important"
         $("#optionframe").hide();
         $("#workframe").show();

         // Print selected nodes and kill checkbox bubbling (so the nodes don't expand aswell)
         $(document).off('click', '.printSelect').on('click', '.printSelect', function (event) { event.stopPropagation(); printSelectClick(this); });

         //Set up print options on ribbon
         $(document).off('click', '.mnuBtnPrintOrgChart').on('click', '.mnuBtnPrintOrgChart', function () { printOrgChart(); });	// print all nodes
         $(document).off('click', '.mnuBtnPrintPreviewOrgChart').on('click', '.mnuBtnPrintPreviewOrgChart', function () { printOrgChart(true); });	// print preview all nodes

         //Enable org chart nodes to be selected for printing.
         $(document).off('click', '.mnuBtnSelectOrgChart').on('click', '.mnuBtnSelectOrgChart', function () {
            $('.printSelect').toggle();
         });

         $(document).off('click', 'div.node').on('click', 'div.node', function () {
            $('div.node.ui-state-active').removeClass('ui-state-active').addClass('ui-state-default');
            $(this).removeClass('ui-state-default').addClass('ui-state-active');
            centreMe(false);
         });

         //Show the click to expand plus/minus icon
         showExpandNodeIcons();

         //enable/disable expand all nodes button
         menu_toolbarEnableItem("mnutoolOrgReportsExpand", ($('.contracted').length > 0));
      }

      setTimeout(function () {
         $(".divMultiline").dotdotdot({ wrap: 'letter', fallbackToLetter: true });
      }, 1);
   }); //--------------End Ready ---------------

   function showExpandNodeIcons() {

      $('.node').each(function () {
         if ($(this).parent().parent().siblings().length > 0) {
            $(this).find('.expandNode').show();
         }
      });

      //set all contracted nodes expand icon to a +
      $('.contracted .expandNode').attr('src', window.ROOT + 'Content/images/plus.gif');

   }

   function centreMe(fSelf) {
      try {
         var classToCentre = (fSelf ? '.node.ui-state-highlight' : '.node.ui-state-active');
         var menuWidth = 0;
         if (!window.menu_isSSIMode()) menuWidth = $('#menuframe').width();
         var workframe = $('#workframeset');
         if (window.currentLayout == "tiles") workframe = $('#chart');

         var myNodePos = $(classToCentre).offset().left;
         var workframeWidth = workframe.width();
         workframeWidth += menuWidth;

         if ((myNodePos > workframeWidth) || (myNodePos < menuWidth)) {
            workframe.animate({ scrollLeft: 0 }, 0);
            myNodePos = $(classToCentre).offset().left;
            workframeWidth = workframe.width();

            var scrollLeftNewPos = myNodePos - ((workframeWidth / 2) + menuWidth) + 48;
            workframe.animate({ scrollLeft: scrollLeftNewPos }, 2000);

         }


      } catch (e) { }
   }

   function printSelectClick(clickObj, event) {

      var fChecked = $(clickObj).prop('checked');

      $(clickObj).parent().parent().parent().nextAll("tr").find(".printSelect").prop('checked', fChecked);
      $(clickObj).parent().parent().parent().nextAll("tr").find(".printSelect").prop('disabled', fChecked);

   }

   function printOrgChart(pfPreview) {

      //calculate fPrintAll flag based on selection
      var fPrintAll = ($('.printSelect').css('display') == "none");
      var divToPrint;
      var untickedItemsCount = $('.printSelect:not(:checked)').length;

      if (($('.printSelect:checked:enabled').length === 0) && (!fPrintAll)) {
         OpenHR.modalMessage("No nodes selected to print.");
      } else {

         var winHeight = 1;
         var winWidth = 1;

         if (OpenHR.isChrome() || pfPreview) {
            winHeight = screen.height / 2;
            winWidth = screen.width / 2;
         }

         //Creates a new window, copies the required html content to it and send it to printer.
         var newWin = window.open("", "_blank", 'toolbar=' + (pfPreview ? 'yes' : 'no') + ',status=no,menubar=no,scrollbars=yes,resizable=yes, width=' + winWidth + ', height=' + winHeight + ', visible=none', "");
         newWin.document.write('<link href=\"' + window.ROOT + 'Scripts/jquery/jOrgChart/css/jquery.jOrgChart.css" rel="stylesheet" />');
         newWin.document.write('<link href=\"' + window.ROOT + 'Scripts/jquery/jOrgChart/css/custom.css" rel="stylesheet" />');
         newWin.document.write('<link href=\"' + window.ROOT + 'Scripts/jquery/jOrgChart/css/prettify.css" rel="stylesheet" />');
         newWin.document.write('<link href=\"' + window.ROOT + 'Content/themes/redmond-segoe/jquery-ui.min.css" rel="stylesheet" />');
         newWin.document.write('<sty');
         newWin.document.write('le>');
         newWin.document.write('body {font-family: "Segoe UI", Verdana; }');
         newWin.document.write('h2 {page-break-before: always;}'); //adds page breaks as required.
         newWin.document.write('.jOrgChart .node {height: 200px !important;width: auto;min-width: 200px;border: 1px solid gray;padding: 5px 0px 0px 0px;overflow: auto;font-weight: bold !important;}');
         newWin.document.write('#divPostEmployees {margin-top:5px !important;}');
         newWin.document.write('</sty');
         newWin.document.write('le>');
         newWin.document.write('<h1 style="width: 400px;">Organisation Reports</h1>');

         $('.printSelect').hide(); //hide the selection tickboxes.
         $('.expandNode').hide(); //hide the selection tickboxes.

         if ((untickedItemsCount > 0) && (fPrintAll !== true)) {

            //Send only selected items to printer.
            // This is different to normal print - it includes page breaks, and expands hidden, selected nodes.
            var pageNo = 1;

            $('.printSelect:checked:enabled').closest('table').each(function () {
               if ($(this).parent().parent().css('visibility') !== "hidden") {
                  $(this).parent().attr('id', 'currentlyPrinting'); //get a handle on the parent table.

                  newWin.document.write('<div class="orgChart" id="chart">');
                  newWin.document.write('<div class="jOrgChart">');
                  newWin.document.write('<table border="0">');
                  if (pageNo > 1) newWin.document.write('<h2 style="width: 400px;">Organisation Chart</h2>');

                  divToPrint = document.getElementById('currentlyPrinting');
                  newWin.document.write(divToPrint.innerHTML);

                  newWin.document.write('</table>');
                  newWin.document.write('</div>');
                  newWin.document.write('</div>');

                  $(this).parent().attr('id', ''); // remove handle for the next branch

                  pageNo += 1;
               }
            });
            $('.printSelect').show(); // redisplay checkboxes.
         } else {
            //print all - just grab the whole div.
            divToPrint = document.getElementById('chart');
            newWin.document.write(divToPrint.innerHTML);
         }

         newWin.document.write('<scri');
         newWin.document.write('pt type="text/javascript">');
         if (!pfPreview) newWin.document.write('setTimeout("this.print(); this.close();", 500);');
         if (pfPreview) newWin.document.write("alert('Press control+P to print this page...');");
         newWin.document.write('</scri');
         newWin.document.write('pt>');
         newWin.document.close();

         showExpandNodeIcons(); // redisplay expand boxes.
      }
   }

</script>

<div>
   <ul id='org' style="display: none;"></ul>
   <ul id='tempList' style="display: none;">
      @Code For Each item In Model.OrgReportChartNodeList
            If Model.IsPostBasedSystem = False Then
            @<li hierarchyLevel="@item.HierarchyLevel"
                 id="@item.LineManagerStaffNo"
                 class="@item.NodeTypeClass">
               <div style="overflow: auto;max-height: 190px;" id="divMainContainer" class="centered">
                  @For Each childitem In item.ReportColumnItemList
                     Html.RenderPartial("_OrganisationReportColumnNode", childitem)
                  Next
               </div>
               <input type="checkbox" class="printSelect" />
               <img title="expand/contract this node" class="expandNode" src='@Url.Content("~/Content/images/minus.gif")' />
               <ul id="@item.EmployeeStaffNo" />
            </li>
            Else
               ''Post based system goes here...
               @<li hierarchyLevel="@item.HierarchyLevel"
                    id="@item.LineManagerStaffNo"
                    class="ui-corner-all ui-state-default">
                  <div style="overflow-x:hidden;overflow-y: auto;max-height: 185px;padding-right: 0px;padding-left: 0px;" id="divMainContainer" class="centered">
                     <div id="divPostTitle" class="truncate centered" style="min-height:20px;text-align: center;display:inline-block;">
                        <span title="@item.PostTitle">@item.PostTitle</span>
                     </div>
                     <div id="divPostColumns">
                        @For Each colitem In item.ReportColumnItemList.Where(Function(m) m.TableID = Model.Hierarchy_TableID)
                           Html.RenderPartial("_OrganisationReportColumnNode", colitem)
                        Next
                     </div>
                     <div style="display:table;padding: 0px 5px;" id="divPostEmployees">
                        @For Each childitem In item.PostWiseNodeList
                        @<div style="min-width:180px;display:table-cell;" class="centered">
                           @If (childitem.ReportColumnItemList.Where(Function(m) m.TableID <> Model.Hierarchy_TableID).Count > 0) Then
                           @<div Style="margin-right:5px;border:1px solid gray;padding:6px;max-width:180px;width:176px;" Class="@childitem.NodeTypeClass centered" EmployeeID="@childitem.EmployeeID">
                              @For Each nonePostItm In childitem.ReportColumnItemList.Where(Function(m) m.TableID <> Model.Hierarchy_TableID)
                                 Html.RenderPartial("_OrganisationReportColumnNode", nonePostItm)
                              Next
                           </div>
                           End If
                        </div>
                        Next
                     </div>
                  </div>
                  <input type="checkbox" class="printSelect" />
                  <img title="expand/contract this node" class="expandNode" src='@Url.Content("~/Content/images/minus.gif")' />
                  <ul id="@item.EmployeeStaffNo" />
               </li>
            End If
         Next
      End Code
   </ul>
</div>
<div class="absolutefull">
   <div class="pageTitleDiv">
      <a href='javascript:loadPartialView("linksMain", "Home", "workframe", null);' title='Back'>
         <i class='pageTitleIcon icon-circle-arrow-left'></i>
      </a>
      <span style="margin-left: 40px; margin-right: 20px" class="pageTitle" id="RecordEdit_PageTitle">@Session("utilname")</span>
   </div>

   <div id="chart" class="orgChart"></div>

   <div id="noData" class="ui-widget-content" style="width: 50%; margin: 0 auto; padding: 20px; border: none;display: none;">
      <h2 class="centered">Cannot display the Organisation Report</h2>
      <br />
      <p class="centered">@Session("ErrorText")</p>
      <p class="centered">Please contact your system administrator.</p>
   </div>
</div>