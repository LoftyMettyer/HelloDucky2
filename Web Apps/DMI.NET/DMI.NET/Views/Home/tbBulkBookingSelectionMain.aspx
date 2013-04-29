<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<!DOCTYPE html>

<%
	Response.Expires = -1
	
	session("selectionType") = Request("selectionType")
	session("tbSelectionDataLoading") = true
%>
<html>
<head>
	
<script src="<%: Url.Content("~/Scripts/jquery/jquery-1.8.3.js")%>"></script>
<script src="<%: Url.Content("~/Scripts/jquery/jquery-ui-1.9.2.custom.js")%>"></script>
<script src="<%: Url.Content("~/Scripts/OpenHR.js")%>" type="text/javascript"></script>	

<meta NAME="GENERATOR" CONTENT="Microsoft Visual Studio 6.0">
<meta http-equiv="refresh" content="<%=session("TimeoutSecs")%>;URL=dialogtimeout">
<title>OpenHR Intranet</title>

<link href="~/Content/OpenHR.css" rel="stylesheet" />

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

	<INPUT type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">
	<INPUT type='hidden' id="txtTableID" name="txtTableID" value="0">
	<INPUT type='hidden' id="txtViewID" name="txtViewID" value="0">
	<INPUT type='hidden' id="txtOrderID" name="txtOrderID" value="0">
	<INPUT type='hidden' id="txtSelectionType" name="txtSelectionType" value="<%=Request("selectionType")%>">
</head>

<div id="mainframeset" name="mainframeset">
  <div data-framesource="tbBulkBookingSelection" name="workframe" id="workframe"><%Html.RenderPartial("~/Views/Home/tbBulkBookingSelection.ascx")%></div>
  <div data-framesource="tbBulkBookingSelectionData" name="dataframe" id="dataframe" style="display: none;"><%Html.RenderPartial("~/Views/Home/tbBulkBookingSelectionData.ascx")%></div>
</div>

</html>


