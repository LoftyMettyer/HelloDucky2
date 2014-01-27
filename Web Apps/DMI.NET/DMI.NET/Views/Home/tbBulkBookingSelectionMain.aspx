<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<%
	Response.Expires = -1
	
	Session("selectionType") = Request("selectionType")
	Session("tbSelectionDataLoading") = True
%>
<html>
<head>
		
	<title>OpenHR Intranet</title>
	<script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>


	<%--Here's the stylesheets for the font-icons displayed on the dashboard for wireframe and tile layouts--%>
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />

	<%--Base stylesheets--%>
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />

	<%--stylesheet for slide-out dmi menu--%>
	<link href="<%: Url.LatestContent("~/Content/contextmenustyle.css")%>" rel="stylesheet" type="text/css" />

	<%--ThemeRoller stylesheet--%>
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />

	<%--jQuery Grid Stylesheet--%>
	<link href="<%: Url.LatestContent("~/Content/ui.jqgrid.css")%>" rel="stylesheet" type="text/css" />

	<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
	<meta http-equiv="refresh" content="<%=session("TimeoutSecs")%>;URL=dialogtimeout">
	<title>OpenHR Intranet</title>
	<object classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
		id="Microsoft_Licensed_Class_Manager_1_0"
		viewastext>
		<param name="LPKPath" value="<%: Url.Content("~/lpks/ssmain.lpk")%>">
	</object>

		<script type="text/javascript">
			function loadAddRecords() {
				var iCount;
				iCount = new Number(document.getElementById("txtLoadCount").value);
				document.getElementById("txtLoadCount").value = iCount + 1;

				if (iCount > 0) {

					var dataForm = document.getElementById("frmGetData");

					dataForm.txtTableID.value = document.getElementById("txtTableID").value;
					dataForm.txtViewID.value = document.getElementById("txtViewID").value;
					dataForm.txtOrderID.value = document.getElementById("txtOrderID").value;
					dataForm.txtFirstRecPos.value = 1;
					dataForm.txtCurrentRecCount.value = 0;
					dataForm.txtPageAction.value = "LOAD";

					refreshData();

				}
			}
	</script>

</head>

<body>
	
	
	<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">
	<input type='hidden' id="txtTableID" name="txtTableID" value="0">
	<input type='hidden' id="txtViewID" name="txtViewID" value="0">
	<input type='hidden' id="txtOrderID" name="txtOrderID" value="0">
	<input type='hidden' id="txtSelectionType" name="txtSelectionType" value="<%=Request("selectionType")%>">

	

	<div id="mainframeset" name="mainframeset">
		<div data-framesource="tbBulkBookingSelection" name="workframe" id="workframe"><%Html.RenderPartial("~/Views/Home/tbBulkBookingSelection.ascx")%></div>
		<div data-framesource="tbBulkBookingSelectionData" name="dataframe" id="dataframe" style="display: none;"><%Html.RenderPartial("~/Views/Home/tbBulkBookingSelectionData.ascx")%></div>
	</div>


</body>

</html>



