@Imports DMI.NET
@Imports DMI.NET.Classes
@Inherits System.Web.Mvc.WebViewPage(Of OrganisationReportChartModel)
<link href=@Url.LatestContent("~/Scripts/jquery/jOrgChart/css/jquery.jOrgChart.css") rel="stylesheet" />
<link href=@Url.LatestContent("~/Scripts/jquery/jOrgChart/css/custom.css") rel="stylesheet" />
<link href=@Url.LatestContent("~/Scripts/jquery/jOrgChart/css/prettify.css") rel="stylesheet" />
<script src=@Url.LatestContent("~/Scripts/html2canvas.js") type="text/javascript"></script>

<style>
   .truncate {
      width: 95%;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
   }

   .jOrgChart .node {
      height: auto;
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
      pointer-events: auto !important;
   }

   .ghostNode > div:first-child {
      background: url("../Content/images/ghostNodeImage.png");
      background-repeat: no-repeat;
      background-position: center;
   }

   .ghostNode img:not(.expandNode), .rootNode img:not(.expandNode) {
      visibility: hidden;
   }

   .jOrgChart .filteredNode > :not([type=checkbox]) {
      opacity: 0.40;
   }

   .jOrgChart .filteredNode > .expandNode {
      opacity: 1;
   }

   .jOrgChart .filteredNode :not(.ui-state-highlight) {
      background-color: rgba(255, 255, 255, 0.5);
   }

   .jOrgChart .filteredNode{
      border: 1px solid lightgray;
   }

</style>
<script>

   $(document).ready(function () {
      // Common logic to show desired ribbon and menu
      $("#workframe").attr("data-framesource", "ORGREPORTS");
      showDefaultRibbon();
      menu_refreshMenu();

      //Disable Utility buttons on first load
      menu_toolbarEnableItem('mnutoolCustomReportsFindForOrgReports', false);
      menu_toolbarEnableItem('mnutoolCalendarReportsFindForOrgReports', false);
      menu_toolbarEnableItem('mnutoolMailMergeFindForOrgReports', false);

      if ('@Model.OrgReportChartNodeList.Any()' == 'False') {
         $('#noData').show();
         menu_toolbarEnableItem('divBtnPrintOrgReports', false);
         menu_toolbarEnableItem('divBtnPrintPreviewOrgReports', false);
         menu_toolbarEnableItem('mnutoolOrgReportsExpand', false);
         menu_toolbarEnableItem('divBtnSelectOrgReports', false);
         menu_toolbarEnableItem('mnutoolSaveRecordOrgReports', false);
         $('.clsSaveRecordOrgReports').prop('disabled', true);
         $('.mnuBtnPrintOrgChart>span').prop('disabled', true);
         $('.mnuBtnPrintOrgChart').prop('disabled', true);
         $('.mnuBtnPrintPreviewOrgChart>span').prop('disabled', true);
         $('.mnuBtnPrintPreviewOrgChart').prop('disabled', true);
         $('.mnuBtnSelectOrgChart>span').prop('disabled', true);
         $('.mnuBtnSelectOrgChart').prop('disabled', true);
      } else {
         menu_toolbarEnableItem('divBtnPrintOrgReports', true);
         menu_toolbarEnableItem('divBtnPrintPreviewOrgReports', true);
         menu_toolbarEnableItem('mnutoolOrgReportsExpand', true);
         menu_toolbarEnableItem('divBtnSelectOrgReports', true);
         menu_toolbarEnableItem('mnutoolSaveRecordOrgReports', true);
         $('.clsSaveRecordOrgReports').prop('disabled', false);
         $('.mnuBtnPrintOrgChart>span').prop('disabled', false);
         $('.mnuBtnPrintOrgChart').prop('disabled', false);
         $('.mnuBtnPrintPreviewOrgChart>span').prop('disabled', false);
         $('.mnuBtnPrintPreviewOrgChart').prop('disabled', false);
         $('.mnuBtnSelectOrgChart>span').prop('disabled', false);
         $('.mnuBtnSelectOrgChart').prop('disabled', false);

         //Generate treebased li-ul structure.
         $("#tempList").find("li").each(function () {
            var lineManagerStaffNo = $(this).attr("id");
            var hierarchyLevel = $(this).attr("hierarchyLevel");
            var parentNode = hierarchyLevel == "0" ? 'org' : lineManagerStaffNo;
            $('#' + parentNode).append($(this));
         });

         //Add a class to collapse all peer trees.
         $("#org li.currentNode").siblings().addClass("collapsed");
         $("#org li.currentNode").parents('li').siblings().addClass("collapsed");

         $('#workframe').attr('overflow', 'auto');

         //Get #org element and generate chart in #chart element
         $("#org").jOrgChart({
            chartElement: '#chart',
            dragAndDrop: false
         });

         setTimeout('centreMe(true)', 500);

         $("#chart").find(".divMainContainer").each(function () {

            //Get total number of employees in one post.
            var totalEmpDiv = $(this).find("#divPostEmployees").children().length;

            //Calculate and set width of parant and respective elements.
            if (totalEmpDiv > 1) {
               //If there are more then one employess in post then set width on multiply by no of employees.
               $(this).parent().css("max-width", totalEmpDiv * 200);
               $(this).parent().css("min-width", totalEmpDiv * 200);
               $(this).find("#divPostTitle").css("max-width", totalEmpDiv * 170);
               $(this).find("#divPostTitle").css("min-width", totalEmpDiv * 170);
               $(this).find("#divPostColumns").css("display", "inline-table");
               $(this).find("#divPostColumns div").css("width", totalEmpDiv * 180);
               $(this).find("#divPostColumns div").css("display", "inline-block");
               $(this).find("#divPostColumns div").css("margin-top", 1);
               $(this).find("#divPostColumns div").css("margin-bottom", 0);
               $(this).find("#divPostEmployees div div").css("width", 176 + totalEmpDiv);
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

         $("#optionframe").hide();
         $("#workframe").show();

         // Print selected nodes and kill checkbox bubbling (so the nodes don't expand aswell)
         $(document).off('click', '.printSelect').on('click', '.printSelect', function (event) { event.stopPropagation(); printSelectClick(this); });

         //Set up print options on ribbon
         $(document).off('click', '.mnuBtnPrintOrgChart').on('click', '.mnuBtnPrintOrgChart', function () { printOrgReport(); });	// print all nodes
         $(document).off('click', '.mnuBtnPrintPreviewOrgChart').on('click', '.mnuBtnPrintPreviewOrgChart', function () { printOrgReport(true); });	// print preview all nodes

         //Enable org chart nodes to be selected for printing.
         $(document).off('click', '.mnuBtnSelectOrgChart').on('click', '.mnuBtnSelectOrgChart', function () {
            $('.printSelect').toggle();
         });

         //Set up Save To File option on ribbon
         $(document).off('click', '.clsSaveRecordOrgReports').on('click', '.clsSaveRecordOrgReports', function () { SaveRecordOrgReports(true); });

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

         //If not highlighted node found  then return.
         if($("#chart .currentNode").length ==0){
            return;
         }

         var classToCentre = '#chart .currentNode';
         var menuWidth = 0;
         if (!window.menu_isSSIMode()) menuWidth = $('#menuframe').width();
         var workframe = $('#workframeset');
         if (window.currentLayout == "tiles") workframe = $('#chart');

         //Find the closest div with .node class
         var nodeDive = $(classToCentre).closest('.node');

         var CurrentNodeLeftPos = nodeDive.offset().left;
         var CurrentNodeTopPos = nodeDive.offset().top;

         var workframeWidth = workframe.width();
         var workframeHeight = workframe.height();
         workframeWidth += menuWidth;

         if ((CurrentNodeLeftPos > workframeWidth) || (CurrentNodeLeftPos < menuWidth)) {
            workframe.animate({ scrollLeft: 0, scrollTop: 0 }, 0);
            workframeWidth = workframe.width();
            workframeHeight = workframe.height();

            CurrentNodeLeftPos = nodeDive.offset().left;
            CurrentNodeTopPos = nodeDive.offset().top;
            //Calculate the top position of highlighted node.
            var scrollTopNewPos = CurrentNodeTopPos -((workframeHeight/2) + $("#banner").height() +$("#fixedlinks").height()) +100;

            //Calculate the left position of highlighted node.
            var scrollLeftNewPos = CurrentNodeLeftPos - ((workframeWidth / 2) + 150) ;

            //Reposition the main containter div according to left and top position of highlighted node.
            workframe.animate({ scrollLeft: scrollLeftNewPos, scrollTop: scrollTopNewPos }, 2000);
         }
      } catch (e) { }
   }
   var divTop;
   var divLeft;
   function GetSelectedNodesFromOrgChart(pfPreview) {

      $("#divSaveToFileContainer").empty();

      //calculate fPrintAll flag based on selection
      var fPrintAll = ($('#chart .printSelect').css('display') == "none");
      var untickedItemsCount = $('#chart .printSelect:not(:checked)').length;

      if (($('#chart .printSelect:checked:enabled').length === 0) && (!fPrintAll)) {
         OpenHR.modalMessage("No nodes selected to Save.");
         return false;
      } else {

         //Get chart position.
         divTop = $('#workframeset').scrollTop();
         divLeft = $('#workframeset').scrollLeft();

         var printcheckboxstatus = true;
         if ( $('#chart .printSelect').css('display') == 'none' ) {
            printcheckboxstatus = false;
         }

         $('#chart .printSelect').hide(); //hide the selection tickboxes.
         $('#chart .expandNode').hide(); //hide the selection tickboxes.

         //If any individual nodes are selected then save only those.
         if ((untickedItemsCount > 0) && (fPrintAll !== true)) {

            //Send only selected items to save.
            $('#chart .printSelect:checked:enabled').closest('table').each(function () {
               if ($(this).parent().parent().css('visibility') !== "hidden") {

                  //get a handle on the parent table.
                  $(this).parent().attr('id', 'currentlyPrinting');

                  var divToPrint = document.getElementById('currentlyPrinting');

                  var divnode = $('<div class="jOrgChart"></div>');
                  divnode.append(divToPrint.innerHTML);
                  divnode.wrap('<div class="orgChart"></div>');

                  $("#divSaveToFileContainer").append(divnode);

                  // remove handle for the next branch.
                  $(this).parent().attr('id', '');
               }
            });
         } else {

            //Save whole chart at onces.
            $("#divSaveToFileContainer").append($("#chart").clone());
         }

         $('#divSaveToFileContainer .printSelect').hide(); //hide the selection tickboxes.
         $('#divSaveToFileContainer .expandNode').hide(); //hide the expand buttons.

         if(printcheckboxstatus){
            $('#chart .printSelect').show(); // redisplay checkboxes.
            menu_toolbarEnableItem('mnutoolCustomReportsFindForOrgReports', true);
            menu_toolbarEnableItem('mnutoolCalendarReportsFindForOrgReports', true);
            menu_toolbarEnableItem('mnutoolMailMergeFindForOrgReports', true);
         }
         else{
            menu_toolbarEnableItem('mnutoolCustomReportsFindForOrgReports', false);
            menu_toolbarEnableItem('mnutoolCalendarReportsFindForOrgReports', false);
            menu_toolbarEnableItem('mnutoolMailMergeFindForOrgReports', false);
         }

         showExpandNodeIcons(); // redisplay expand boxes.
         return true;
      }
   }

   function isIEOrEdgeBrowser(userAgent) {
      userAgent = userAgent || navigator.userAgent;
      return userAgent.indexOf("MSIE ") > -1 || userAgent.indexOf("Trident/") > -1 || userAgent.indexOf("Edge/") > -1;
   }

   function SaveRecordOrgReports() {

      $("body").addClass("loading");
      menu_ShowWait('Please wait...');

      if (GetSelectedNodesFromOrgChart() == false) {
         $("body").removeClass("loading");
         return;
      }

      setTimeout(function(){

         $('#divSaveToFileParent').show();
         if ((isIEOrEdgeBrowser()==true && $('#divSaveToFileContainer').prop('scrollWidth')>=8000) ||
             ($('#divSaveToFileContainer').prop('scrollWidth')>=32000)) {

            OpenHR.modalMessage("Selected nodes exceeds the limit of file, Please select lesser number of nodes.");
            window.console.log("Canvas IndexSizeError: Index or size is negative or greater than the allowed amount.");

            $("body").removeClass("loading");
            $('#divSaveToFileParent').hide();

            return false;
         }

         $("#divSaveToFileContainer .printSelect").hide();
         $('#divSaveToFileContainer').scrollTop(0).scrollLeft(0);
         $('#workframeset').scrollTop(0).scrollLeft(0);

         var useWidth = $('#divSaveToFileParent').prop('scrollWidth') + 500;
         var useHeight = $('#divSaveToFileParent').prop('scrollHeight') + 300;

         html2canvas($("#divSaveToFileParent"), {
            useCORS: true,
            logging: true,
            onrendered: function (canvas) {
               //For IE/Edge browser the image will download in local folder.
               if (canvas.msToBlob) {

                  //In IE/Edge save canvas as Blob object.
                  window.navigator.msSaveBlob(new Blob([canvas.msToBlob()],{type:"image/png"}), '@Session("utilname")' + ".png");

                  $('#divSaveToFileParent').hide();
                  $("body").removeClass("loading");

                  //Set the chart position according to previous state.
                  $('#workframeset').scrollTop(divTop).scrollLeft(divLeft);
               } else {

                  //For all other browser the image will open in new tab window.
                  $("body").removeClass("loading");

                  var data = canvas.toDataURL("image/png");

                  //Create div element and add image into it.
                  var div = document.createElement('div');
                  var img = document.createElement('img');
                  img.src = data;

                  //Create window object to open in new tab.
                  var newWin = window.open();

                  //First Checking Condition Works For IE & Firefox
                  //Second Checking Condition Works For Chrome
                  if (newWin == null || typeof(newWin)=='undefined') {
                     OpenHR.modalMessage("Please disable your pop-up blocker and click the 'Save To File' button again.");
                     return;
                  }

                  //Open image in new tab window.
                  newWin.location.href = canvas.toDataURL();

                  $('#divSaveToFileParent').hide();
                  $('#workframeset').scrollTop(divTop).scrollLeft(divLeft);
               }
               $("#divSaveToFileContainer").empty();
            },
            width: useWidth,
            height: useHeight,
            allowTaint: true
         });
      },10);
   }

   function printSelectClick(clickObj, event) {
      //Disable Utility Buttons if no record selected
      if ($('.printSelect:checked:enabled').length === 0 ) {
         menu_toolbarEnableItem('mnutoolCustomReportsFindForOrgReports', false);
         menu_toolbarEnableItem('mnutoolCalendarReportsFindForOrgReports', false);
         menu_toolbarEnableItem('mnutoolMailMergeFindForOrgReports', false);
      }
      else{
         menu_toolbarEnableItem('mnutoolCustomReportsFindForOrgReports', true);
         menu_toolbarEnableItem('mnutoolCalendarReportsFindForOrgReports', true);
         menu_toolbarEnableItem('mnutoolMailMergeFindForOrgReports', true);
      }
      var fChecked = $(clickObj).prop('checked');

      $(clickObj).parent().parent().parent().nextAll("tr").find(".printSelect").prop('checked', fChecked);
      $(clickObj).parent().parent().parent().nextAll("tr").find(".printSelect").prop('disabled', fChecked);
      GetSelectedEmployeeIDs();
   }

   function printOrgReport(pfPreview) {

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
         newWin.document.write('<link href=\"' + window.ROOT + 'Content/themes/oneAdvanced/jquery-ui.min.css" rel="stylesheet" />');
         newWin.document.write('<sty');
         newWin.document.write('le>');
         newWin.document.write('body {font-family: "Segoe UI", Verdana; }');
         newWin.document.write('h2 {page-break-before: always;}'); //adds page breaks as required.
         newWin.document.write('.jOrgChart .node {height: auto;width: auto;min-width: 200px;border: 1px solid gray;padding: 5px 0px 0px 0px;overflow: auto;font-weight: bold !important;}');
         newWin.document.write('#divPostColumns div {margin-top:1px !important;}');
         newWin.document.write('#divPostColumns div {margin-bottom:0px !important;}');
         newWin.document.write('#divPostColumns div {height:auto !important;}');
         newWin.document.write('.ghostNode > div:first-child { background: url("../Content/images/ghostNodeImage.png"); background-repeat: no-repeat; background-position: center;}');
         newWin.document.write('.ghostNode img, .rootNode img { visibility: hidden; }');
         newWin.document.write('.jOrgChart .ui-state-disabled { opacity: 0.20 }');
         newWin.document.write('.truncate {white-space: nowrap;overflow: hidden;text-overflow: ellipsis;}');
         newWin.document.write('.expandNode {bottom: 4px;right: 4px;}');
         newWin.document.write('.jOrgChart .filteredNode > :not([type=checkbox]) { opacity: 0.20; }');
         newWin.document.write('.jOrgChart .filteredNode > .expandNode { opacity: 1; }');
         newWin.document.write('.jOrgChart .filteredNode { border: 1px solid lightgray; }');
         newWin.document.write('</sty');
         newWin.document.write('le>');
         newWin.document.write('<h1>Organisation Reports : ' + '@Session("utilname")' +' </h1>');

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
                  if (pageNo > 1) newWin.document.write('<h2>Organisation Reports : ' + '@Session("utilname")' + '</h2>');

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

   //Get id's of selected records
   function GetSelectedEmployeeIDs() {
      var SelectedIds=[];
      if ('@Model.IsPostBasedSystem'=='True') {
         $('.printSelect:checked').each(function() {
            SelectedIds.push($(this).attr("postid"));
         });
         $("#txtSelectedRecordsInFindGrid")[0].value = SelectedIds.sort(function(a, b){return b-a});
         $("#txtOrgReportTableID")[0].value = @Model.Hierarchy_TableID;
      }
      else{
         $('.printSelect:checked').each(function () {
            SelectedIds.push($(this).attr("employeeid"));
         });
         $("#txtSelectedRecordsInFindGrid")[0].value = SelectedIds.sort(function(a, b){return b-a});
         $("#txtOrgReportTableID")[0].value = @Model.Hierarchy_TableID;
      }
   }

   function refreshData()
   {
      $("#toolbarReportFind").parent().hide();
   }

   //Disable Utility buttons(Custom,Calender,Mail-Merge) for SSI mode.
   if(menu_isSSIMode() == true){
      $("#mnuSectionReportsAndUtilityForOrgReports").hide();
   }else{
      $("#mnuSectionReportsAndUtilityForOrgReports").show();
   }
</script>

<div>
   <ul id='org' style="display: none;"></ul>
   <ul id='tempList' style="display: none;">
      @Code For Each item In Model.OrgReportChartNodeList
            'Commercial Based System
            If Model.IsPostBasedSystem = False Then
      @<li hierarchyLevel="@item.HierarchyLevel"
           id="@item.LineManagerStaffNo"
           class="@item.NodeTypeClass">
         <div style="overflow-x:hidden;overflow-y: hidden;" class="divMainContainer centered">
            @For Each childitem In item.ReportColumnItemList        'Render all columns form defination.
               Html.RenderPartial("_OrganisationReportColumnNode", childitem)
            Next
         </div>
         <input type="checkbox" class="printSelect" employeeid="@item.EmployeeID" hierarchyLevel="@item.HierarchyLevel" />
         <img title="expand/contract this node" class="expandNode" src='@Url.Content("~/Content/images/minus.gif")' hierarchyLevel="@item.HierarchyLevel" />
         <ul id="@item.EmployeeStaffNo" />
      </li>
            Else
               ''Post based system goes here...   
               Dim ParentNodeClass = item.PostWiseNodeList.FirstOrDefault.NodeTypeClass.ToString
               Dim ChildNodeStyle As String = ""
      @<li hierarchyLevel="@item.HierarchyLevel"
           id="@item.LineManagerStaffNo"
           class="@ParentNodeClass ui-corner-all ui-state-default">
         <div style="overflow-x:hidden;overflow-y: hidden;padding-right: 0px;padding-left: 0px;" class="divMainContainer centered">
            <div id="divPostColumns">
               @For Each colitem In item.ReportColumnItemList.Where(Function(m) m.TableID = Model.Hierarchy_TableID)  'Render only basedview columns.
                  Html.RenderPartial("_OrganisationReportColumnNode", colitem)
               Next
            </div>

            <div style="display:table;padding: 0px 5px;margin-bottom:15px;" id="divPostEmployees">
               @For Each childitem In item.PostWiseNodeList  'Create internal boxes for each employee.
                  If childitem.IsVacantPost Then
                     ChildNodeStyle = "background:url('../Content/images/vacancytext.png');"
                  Else
                     ChildNodeStyle = ""
                  End If
               @<div style="min-width:180px;display:table-cell;" class="centered">
                  @If (childitem.ReportColumnItemList.Where(Function(m) m.TableID <> Model.Hierarchy_TableID).Count > 0) Then
               @<div Style="@ChildNodeStyle;margin-right:5px;border:1px solid gray;padding:6px;max-width:180px;width:176px;position: relative; overflow: hidden;" Class="centered" EmployeeID="@childitem.EmployeeID">
                  @For Each nonePostItm In childitem.ReportColumnItemList.Where(Function(m) m.TableID <> Model.Hierarchy_TableID)
                     Html.RenderPartial("_OrganisationReportColumnNode", nonePostItm)
                  Next
               </div>
                  End If
               </div>
               Next
            </div>

         </div>
         @If item.IsVacantPost = False Then       'Show Select checkbox option only for non-vacant post.
      @<input type="checkbox" Class="printSelect" postid="@item.PostID" hierarchyLevel="@item.HierarchyLevel" />
         End If
         <img title="expand/contract this node" Class="expandNode" hierarchyLevel="@item.HierarchyLevel" src='@Url.Content("~/Content/images/minus.gif")' />
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
<div id="divSaveToFileParent" style="display:none;position:absolute;z-index:-10000;background-color:white;">
   <h2 style="margin-left:20px;">Organisation Report : @Session("utilname")</h2>
   <div id="divSaveToFileContainer">
   </div>
</div>
