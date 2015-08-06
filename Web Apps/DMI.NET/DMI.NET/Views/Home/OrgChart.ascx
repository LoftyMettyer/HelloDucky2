<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(Of IEnumerable (Of DMI.NET.Models.OrgChart))" %>
<%@Import namespace="DMI.NET" %>

<link href="<%: Url.LatestContent("~/Scripts/jquery/jOrgChart/css/jquery.jOrgChart.css")%>" rel="stylesheet" />
<link href="<%: Url.LatestContent("~/Scripts/jquery/jOrgChart/css/custom.css")%>" rel="stylesheet" />
<link href="<%: Url.LatestContent("~/Scripts/jquery/jOrgChart/css/prettify.css")%>" rel="stylesheet" />

<style>
	.MAT, .SICK, .HOLS, .COMP {
		background-color: lightgray!important;
		background-image: none!important;
	}
</style>

<script>

	$(document).ready(function () {


		// Common logic to show desired ribbon and menu
		$("#workframe").attr("data-framesource", "ORGCHART");
		showDefaultRibbon();
		menu_refreshMenu();

		if ('<%=Model.any()%>' == 'False') {
			$('#noData').show();
			menu_toolbarEnableItem('divBtnPrintOrgChart', false);
			menu_toolbarEnableItem('divBtnPrintPreviewOrgChart', false);
			menu_toolbarEnableItem('mnutoolOrgChartExpand', false);
			menu_toolbarEnableItem('divBtnSelectOrgChart', false);
			$('.mnuBtnPrintOrgChart>span').prop('disabled', true);
			$('.mnuBtnPrintOrgChart').prop('disabled', true);
			$('.mnuBtnPrintPreviewOrgChart>span').prop('disabled', true);
			$('.mnuBtnPrintPreviewOrgChart').prop('disabled', true);
			$('.mnuBtnSelectOrgChart>span').prop('disabled', true);
			$('.mnuBtnSelectOrgChart').prop('disabled', true);

		} else {
			//process the results into unordered list.		
			$("#hiddenItems").find(":hidden").not("script").each(function () {
				var props = $(this).val().split("\t");
				var employeeID = props[0],
					employeeForenames = props[1],
					employeeSurname = props[2],
					employeeStaffNo = props[3],
					lineManagerStaffNo = props[4],
					employeeJobTitle = props[5],
					hierarchyLevel = props[6],
					photoPath = props[7],
					absenceTypeClass = props[8];

				//If hierarchy level = 0 add to root (#org), otherwise append to previous manager's staff_number
				var parentNode = hierarchyLevel == "0" ? 'org' : lineManagerStaffNo;
				var nodeHTML = '<li class="' + absenceTypeClass + '">';	//this is converted to a div at runtime
				nodeHTML += '<input type="checkbox" class="printSelect"/>';
				nodeHTML += '<img title="expand/contract this node" class="expandNode" src="' + window.ROOT + 'Content/images/minus.gif"/>';
				nodeHTML += '<div class="jobTitle">' + employeeJobTitle + '</div>';
				nodeHTML += '<img style="width: 48px; height: 48px;" src="' + photoPath + '"/>';
				nodeHTML += '<p>' + employeeForenames + ' ' + employeeSurname + '</p>';
				nodeHTML += '<ul id="' + employeeStaffNo + '">';
				nodeHTML += '</li>';
				$('#' + parentNode).append(nodeHTML);
			});

			//Add a class to collapse all peer trees.
			$("#org li.ui-state-highlight").siblings().addClass("collapsed");
			$("#org li.ui-state-highlight").parents('li').siblings().addClass("collapsed");

			$('#workframe').attr('overflow', 'auto');
			$("#org").jOrgChart({
				chartElement: '#chart',
				dragAndDrop: false
			});

			setTimeout('centreMe(true)', 500);

			//Set up tool tip for absentees...
			$("div[class*='REASON']").each(function () {
				try {
					var classString = $(this).attr('class');

					if (OpenHR.nullsafeString(classString).length > 0) {
						var absReason = classString.substr(classString.indexOf('#') + 1);
						absReason = absReason.substr(0, absReason.indexOf('#'));

						$(this).attr('title', absReason);
					}
				}
				catch (e) { }
			});

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
			menu_toolbarEnableItem("mnutoolOrgChartExpand", ($('.contracted').length > 0));
			

		}
	});

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


		} catch(e) {}
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
			newWin.document.write('</sty');
			newWin.document.write('le>');
			newWin.document.write('<h1 style="width: 400px;">Organisation Chart</h1>');			

			$('.printSelect').hide(); //hide the selection tickboxes.
			$('.expandNode').hide(); //hide the selection tickboxes.

			if ((untickedItemsCount > 0) && (fPrintAll !== true)) {

				//Send only selected items to printer. 
				// This is different to normal print - it includes page breaks, and expands hidden, selected nodes.
				var pageNo = 1;

				$('.printSelect:checked:enabled').closest('table').each(function() {
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


<div id="hiddenItems">
	<ul id='org' style="display: none;">
	</ul>
	<%	For Each item In Model%>
	<%Dim inputString As String
		inputString = (item.EmployeeID & vbTab &
									 item.EmployeeForenames & vbTab &
									 item.EmployeeSurname & vbTab &
									 item.EmployeeStaffNo & vbTab &
									 item.LineManagerStaffNo & vbTab &
									 item.EmployeeJobTitle & vbTab &
									 item.HierarchyLevel & vbTab &
									 item.PhotoPath & vbTab &
									 item.AbsenceTypeClass)%>
	<input type='hidden' value='<%=InputString%>' />
	<%	Next%>
</div>

<div class="absolutefull">
	<div class="pageTitleDiv">
		<a href='javascript:loadPartialView("linksMain", "Home", "workframe", null);' title='Back'>
			<i class='pageTitleIcon icon-circle-arrow-left'></i>
		</a>
		<span style="margin-left: 40px; margin-right: 20px" class="pageTitle" id="RecordEdit_PageTitle">Organisation Chart</span>
	</div>

	<div id="chart" class="orgChart"></div>

	<div id="noData" class="ui-widget-content" style="width: 50%; margin: 0 auto; padding: 20px; border: none;display: none;">
		<h2 class="centered">Cannot display the Organisation Chart</h2>
		<br />
		<p class="centered"><%=Session("ErrorText")%></p>		
		<p class="centered">Please contact your system administrator.</p>
	</div>

</div>


