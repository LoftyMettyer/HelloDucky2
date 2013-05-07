<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<!DOCTYPE html>

<%
	Response.Expires = -1
	
	Session("selectionType") = Request("selectionType")
	Session("tbSelectionDataLoading") = True
%>
<html>
<head>
    
    <script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/bundles/jQueryUI")%>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>           
    <link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css"/>

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

	<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">
	<input type='hidden' id="txtTableID" name="txtTableID" value="0">
	<input type='hidden' id="txtViewID" name="txtViewID" value="0">
	<input type='hidden' id="txtOrderID" name="txtOrderID" value="0">
	<input type='hidden' id="txtSelectionType" name="txtSelectionType" value="<%=Request("selectionType")%>">
</head>
<body>
	<div id="mainframeset" name="mainframeset">
		<div data-framesource="tbBulkBookingSelection" name="workframe" id="workframe"><%Html.RenderPartial("~/Views/Home/tbBulkBookingSelection.ascx")%></div>
		<div data-framesource="tbBulkBookingSelectionData" name="dataframe" id="dataframe" style="display: none;"><%Html.RenderPartial("~/Views/Home/tbBulkBookingSelectionData.ascx")%></div>
	</div>

</body>

</html>


