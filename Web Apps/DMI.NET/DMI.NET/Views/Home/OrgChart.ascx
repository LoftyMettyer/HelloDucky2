<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(Of IEnumerable (Of DMI.NET.Models.OrgChart))" %>
<%@Import namespace="DMI.NET" %>

<link href="<%= Url.LatestContent("~/Scripts/jquery/jOrgChart/css/jquery.jOrgChart.css")%>" rel="stylesheet" />
<link href="<%= Url.LatestContent("~/Scripts/jquery/jOrgChart/css/custom.css")%>" rel="stylesheet" />
<link href="<%= Url.LatestContent("~/Scripts/jquery/jOrgChart/css/prettify.css")%>" rel="stylesheet" />

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
		$("#toolbarHome").click();

		if ('<%=Model.any()%>' == 'False') {
				$('#noData').show();
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
					$('#' + parentNode).append('<li class="' + absenceTypeClass + '">' + employeeJobTitle + '<img style="width: 48px; height: 48px;" src="' + photoPath + '"/><p>' + employeeForenames + ' ' + employeeSurname + '</p><ul id="' + employeeStaffNo + '"></li>');
				});

				//Add a class to collapse all peer trees.
				$("#org li.ui-state-active").siblings().addClass("collapsed");

				$('#workframe').attr('overflow', 'auto');
				$("#org").jOrgChart({
					chartElement: '#chart',
					dragAndDrop: true
				});

				setTimeout('centreMe()', 500);

				$("#optionframe").hide();
				$("#workframe").show();
			}
	});

		function centreMe() {
			var myNodePos = $('.node.ui-state-active').offset().left;
			var workframeWidth = $('#workframeset').width();
			var scrollLeftNewPos = myNodePos - workframeWidth + 380 + 48;

			$('#workframeset').animate({ scrollLeft: scrollLeftNewPos }, 2000);

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
		<span style="margin-left: 40px; margin-right: 20px" class="pageTitle" id="RecordEdit_PageTitle">Organisation Chart
		</span>
	</div>

	<div id="chart" class="orgChart"></div>

	<div id="noData" class="ui-widget-content" style="width: 50%; margin: 0 auto; padding: 20px; border: none;display: none;">
		<h2 class="centered">Cannot display the Organisation Chart</h2>
		<br />
		<p class="centered"><%=Session("ErrorText")%></p>		
		<p class="centered">Please contact your system administrator.</p>
	</div>

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>

</div>


